#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf2docx.py —— 把 PDF 转成 Word(.docx)，尽量保留段落 / 表格 / 图片的格式与位置。

本文件基于 pdf2word_v2.py，功能与命令行参数保持一致，主要针对「最后一页解析完后，
写 docx 阶段要等好几分钟」的问题做了优化，并补上了各阶段的耗时日志：

  * 表格写入：python-docx 的 Table._cells / row.cells 每次访问都会遍历整张表重建
    全部单元格对象，原实现在「行 x 列」双重循环里反复触发，复杂度是 O(单元格数^2)，
    几百个单元格的大表要花几十秒。这里改成整表网格只取一次、合并后重取一次。
  * 段落写入：Heading / Table Grid 样式对象只查一次，不再每段都走 styles 查找。
  * 阶段日志：解析 / 页眉页脚 / 版面分析 / 跨页合并 / 写 docx / 落盘，各自打印耗时，
    以后慢在哪一步一眼能看到。

依赖:  pip install pdfplumber python-docx pillow pypdf
用法:  python pdf2docx.py input.pdf [-o output.docx] [--pages 1-5,8] [--password xxx]
       [--keep-page-breaks] [--no-columns] [--no-figures] [--no-headings]
       [--no-header-footer] [--cache-page] [--cache-dir DIR] [-v]
       [--remove-page-head-foot]

页眉页脚:
      默认行为是把跨页重复的页眉/页脚/页码从正文里摘出来，放进 Word 的页眉页脚
      (页码写成 PAGE/NUMPAGES 域)。加 --remove-page-head-foot 后，除完整版 docx 外
      再额外保存一份「已删除页眉、页脚、页码」的 docx，并把每一页检测到的内容
      输出到控制台与 <输出名>-page-head-foot.json。

缓存: --cache-page 启用后，每一页的解析结果会以可读格式存到 --cache-dir 指定的目录
      (默认: 输出 docx 同目录下的 _pdf2word_pages)：
        page_XXXX.json  该页结构化结果（行/字符/表格/图片框，UTF-8 可读）
        page_XXXX.txt   该页纯文本视图（逐行文本 + 表格单元格文本）
        page_XXXX_*.jpg/png  该页的图片/矢量图渲染结果
        manifest.json   全部页的摘要索引
      默认不缓存任何逐页内容。

核心思路
--------
1. 逐页解析 (pdfplumber)：
   - 表格 : find_tables() -> 网格 + 合并单元格 + 底色 + 单元格内的文字/图片
   - 图片 : 优先直接取原始图像流 (JPEG 原样 / 无损 RGB/灰度)，不行再按区域渲染
   - 矢量图(图表等): 聚类曲线/色块后整体渲染成图片，并剔除其内部文字，避免重复
   - 文字 : 字符 -> 行 -> 段落 (带字体/字号/粗斜体/颜色)；自动识别双栏
2. 全局处理：
   - 页眉页脚：跨页重复的行识别出来，放入 Word 页眉/页脚，页码用域代码
   - 段落折行合并：根据 "上一行右侧剩余空间能否放下下一行首个词" 判断是软换行还是真换行
     (英文去连字符、中文不加空格)
   - 跨页合并：表格 (列边界对齐 + 重复表头去除) 与段落 (末行写满且下页无缩进)
3. 写 docx：页面尺寸/页边距、对齐、缩进、行距/段间距、表格列宽行高、图片大小/对齐
4. 鲁棒性：加密/损坏 PDF、单页失败整页转图片兜底、非法 XML 字符、超大表格/图片等
"""
from __future__ import annotations

import argparse
import io
import json
import logging
import math
import os
import re
import sys
import time
from collections import Counter, defaultdict
from dataclasses import dataclass, field, replace
from typing import Any, Dict, List, Optional, Set, Tuple

import pdfplumber
from docx import Document
from docx.enum.section import WD_ORIENT, WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor
from PIL import Image

log = logging.getLogger("pdf2word")
_STATS: Counter = Counter()   # 统计无法映射的字形数量，用于提示 "可能需要 OCR"

# ============================================================================
# 配置与数据结构
# ============================================================================


@dataclass
class Config:
    pages: Optional[Set[int]] = None      # 0-based 页码集合，None=全部
    password: Optional[str] = None
    keep_page_breaks: bool = False        # 是否在原 PDF 分页处强制分页
    detect_headings: bool = True          # 大字号短段落 -> Word 标题样式
    detect_columns: bool = True           # 双栏识别
    render_figures: bool = True           # 矢量图整体渲染成图片
    strip_header_footer: bool = True      # 页眉页脚放入 Word 页眉页脚
    remove_page_head_foot: bool = False   # 额外输出一份去掉页眉/页脚/页码的 docx
    keep_fonts: bool = True               # 保留字体名
    image_dpi: int = 200                  # 渲染图片的 dpi
    max_table_cells: int = 6000           # 超过则不当表格处理
    cache_pages: bool = False             # 保存每一页的解析结果 (默认关闭)
    cache_dir: Optional[str] = None       # 逐页缓存目录；None 时用默认目录
    jobs: int = 1                         # 解析进程数；>1 时多进程按页并行（默认单进程）


@dataclass
class Run:
    text: str
    font: Optional[str]
    size: float
    bold: bool
    italic: bool
    color: Optional[Tuple[int, int, int]]

    def style(self):
        return (self.font, self.size, self.bold, self.italic, self.color)


@dataclass
class Line:
    runs: List[Run]
    x0: float
    x1: float
    top: float          # 字符 top 的中位数 (用于行距)
    bottom: float
    y0: float           # 外接范围
    y1: float
    size: float
    first_tok_w: float  # 首个"不可拆分单元"的宽度 (英文=首词, 中文=首字)
    chars: list = field(default_factory=list, repr=False)
    kind: str = "full"  # full / left / right (分栏)
    band: int = 0
    col: int = 0

    @property
    def text(self) -> str:
        return "".join(r.text for r in self.runs)

    @property
    def bold_all(self) -> bool:
        rs = [r for r in self.runs if r.text.strip()]
        return bool(rs) and all(r.bold for r in rs)


@dataclass
class ImageData:
    data: bytes
    w: float
    h: float


@dataclass
class ParaData:
    lines: List[Line]
    L: float            # 用于缩进计算的左基线
    R: float            # 用于折行判断的右边界
    box_L: float        # 对齐判断用的容器左右边界
    box_R: float
    pitch: float
    space_before: float = 0.0
    heading: int = 0
    in_cell: bool = False


@dataclass
class CellData:
    r0: int
    r1: int
    c0: int
    c1: int
    paras: List[ParaData] = field(default_factory=list)
    images: List[Tuple[float, ImageData, float]] = field(default_factory=list)  # (top, img, x0)
    shade: Optional[str] = None
    box: Tuple[float, float, float, float] = (0, 0, 0, 0)

    def text(self) -> str:
        return " ".join(
            "".join(r.text for r in lines_to_runs(p.lines)) for p in self.paras
        ).strip()


@dataclass
class TableData:
    xs: List[float]
    row_h: List[float]
    cells: List[CellData]
    header_rows: int = 0   # 跨页重复出现的表头行数 -> Word "标题行重复"

    @property
    def n_rows(self) -> int:
        return len(self.row_h)


@dataclass
class Block:
    kind: str                 # para / table / image / imggrid
    x0: float
    top: float
    x1: float
    bottom: float
    page: int
    data: Any
    last_page: int = -1
    page_start: bool = False
    band: int = 0
    col: int = 0
    gap_before: float = 0.0

    def __post_init__(self):
        if self.last_page < 0:
            self.last_page = self.page


@dataclass
class ImageGrid:
    """同一水平带（或上下紧邻）的并排图片，按原版面排成网格。

    rows 里每一行是若干图片 Block，行内按 x0 升序；写 docx 时按列边界
    还原左右/四象限等多图排布，而不是每张图各占一段、从上往下堆叠。
    """
    rows: List[List[Block]]
    x0: float
    top: float
    x1: float
    bottom: float


@dataclass
class PageRaw:
    index: int
    W: float
    H: float
    lines: List[Line] = field(default_factory=list)
    tables: List[Block] = field(default_factory=list)
    images: List[Block] = field(default_factory=list)
    fallback: Optional[ImageData] = None   # 整页解析失败时的整页截图
    rules: List[Tuple[float, float, float, float]] = field(default_factory=list)  # 顶部/底部横线 (top,bottom,x0,x1)


@dataclass
class HFPart:
    line: Line
    is_field: bool


@dataclass
class HFItem:
    page: int                 # 0-based 页码
    kind: str                 # header / footer / page_number
    text: str
    y0: float
    y1: float
    x0: float
    x1: float
    contains_page_number: bool = False


# ============================================================================
# 通用小工具
# ============================================================================

_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]")
_CHAR_MAP = {
    "\uf0b7": "•", "\uf0a7": "▪", "\uf0d8": "➢", "\uf0fc": "✓", "\uf06e": "■",
    "\u00a0": " ", "\u00ad": "", "\u200b": "", "\ufeff": "",
}
_CJK_RANGES = (
    (0x2E80, 0x9FFF), (0xF900, 0xFAFF), (0x3000, 0x303F), (0xFF00, 0xFFEF),
    (0xAC00, 0xD7AF), (0x3040, 0x30FF), (0x20000, 0x2FA1F),
)
_LIST_RE = re.compile(
    r"^\s*(?:[•▪■◦●○·➢✓►*]\s?|[\-–—]\s|\(?\d{1,3}[\.\)、）]|[（(]\d{1,3}[)）]"
    r"|[（(][a-zA-Z一二三四五六七八九十]{1,3}[)）]|[a-zA-Z][\.\)]\s|[一二三四五六七八九十]+[、．.]"
    r"|\d+\.\d+[\.\s]|第[一二三四五六七八九十百\d]+[章节条款篇])"
)
_TERMINAL = "。！？.!?；;：:”\"'）)"

_FONT_MAP = {
    "timesnewroman": "Times New Roman", "times": "Times New Roman", "timesroman": "Times New Roman",
    "arial": "Arial", "helvetica": "Arial", "calibri": "Calibri", "cambria": "Cambria",
    "couriernew": "Courier New", "courier": "Courier New", "verdana": "Verdana",
    "tahoma": "Tahoma", "georgia": "Georgia", "garamond": "Garamond", "symbol": "Symbol",
    "simsun": "宋体", "宋体": "宋体", "simhei": "黑体", "黑体": "黑体", "kaiti": "楷体",
    "simkai": "楷体", "fangsong": "仿宋", "simfang": "仿宋", "microsoftyahei": "微软雅黑",
    "yahei": "微软雅黑", "msyh": "微软雅黑", "stsong": "宋体", "stsonglight": "宋体",
    "notosanscjksc": "Noto Sans CJK SC", "notoserifcjksc": "Noto Serif CJK SC",
}


def is_cjk(ch: str) -> bool:
    if not ch:
        return False
    o = ord(ch[0])
    return any(a <= o <= b for a, b in _CJK_RANGES)


def sanitize(text: str) -> str:
    return _ILLEGAL_XML.sub("", text or "")


def rgb_of(color) -> Optional[Tuple[int, int, int]]:
    if color is None:
        return None
    try:
        vals = (float(color),) if isinstance(color, (int, float)) else tuple(float(v) for v in color)
    except Exception:
        return None
    if len(vals) == 1:
        rgb = (vals[0],) * 3
    elif len(vals) == 3:
        rgb = vals
    elif len(vals) == 4:
        c, m, y, k = vals
        rgb = ((1 - c) * (1 - k), (1 - m) * (1 - k), (1 - y) * (1 - k))
    else:
        return None
    return tuple(max(0, min(255, int(round(v * 255)))) for v in rgb)  # type: ignore


def run_color(color) -> Optional[Tuple[int, int, int]]:
    rgb = rgb_of(color)
    if rgb is None or max(rgb) < 40:  # 黑色不必特意设置
        return None
    return rgb


_SUBSET = re.compile(r"^[A-Z]{6}\+")


def parse_font(fontname: str, keep_fonts: bool = True):
    """返回 (族名或None, bold, italic)"""
    n = _SUBSET.sub("", fontname or "")
    low = n.lower()
    bold = any(k in low for k in ("bold", "black", "heavy", "semibold", "demibold", "extrabold", "-bd", "黑"))
    italic = any(k in low for k in ("italic", "oblique", "-it", ",it", "slanted"))
    if not keep_fonts:
        return None, bold, italic
    fam = re.split(r"[-,]", n)[0]
    fam = re.sub(r"(PSMT|MT|PS|Std|Pro)$", "", fam)
    key = fam.lower().replace(" ", "")
    if key in _FONT_MAP:
        return _FONT_MAP[key], bold, italic
    if len(fam) < 3 or re.match(r"^[A-Za-z]{1,4}\d+$", fam) or "cid" in key or "identity" in key:
        return None, bold, italic
    return fam, bold, italic


def cluster_values(vals: List[float], tol: float) -> List[float]:
    vals = sorted(vals)
    groups: List[List[float]] = []
    for v in vals:
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def nearest_idx(edges: List[float], v: float) -> int:
    return min(range(len(edges)), key=lambda i: abs(edges[i] - v))


def center_in(c_x0, c_top, c_x1, c_bottom, box, pad: float = 0.5) -> bool:
    cx, cy = (c_x0 + c_x1) / 2, (c_top + c_bottom) / 2
    return box[0] - pad <= cx <= box[2] + pad and box[1] - pad <= cy <= box[3] + pad


# ============================================================================
# 字符 -> 行 -> 段落
# ============================================================================


def char_style(c: dict, cfg: Config):
    fam, b, i = parse_font(c.get("fontname", ""), cfg.keep_fonts)
    size = round(float(c.get("size") or 10) * 2) / 2
    return (fam, max(size, 1.0), b, i, run_color(c.get("non_stroking_color")))


def make_line(chars: List[dict], cfg: Config) -> Optional[Line]:
    cs = sorted(chars, key=lambda c: c["x0"])
    runs: List[Run] = []
    kept: List[dict] = []
    last_ch, prev = "", None
    for c in cs:
        t = _CHAR_MAP.get(c["text"], c["text"])
        if t.startswith("(cid:"):
            # 字体缺少 Unicode 映射：行首的多半是项目符号，其余丢弃，避免输出 "(cid:127)" 乱码
            _STATS["cid"] += 1
            t = "•" if not kept else ""
        t = sanitize(t)
        if not t:
            continue
        if t.isspace():
            t = " "
        st = char_style(c, cfg)
        if t == " ":
            if runs and last_ch != " ":
                runs[-1].text += " "
                last_ch = " "
            continue
        if prev is not None and last_ch not in ("", " "):
            gap = c["x0"] - prev["x1"]
            if gap > 0.22 * min(c["size"], prev["size"]) and not (is_cjk(last_ch) and is_cjk(t[0])):
                runs[-1].text += " "
        if runs and runs[-1].style() == st:
            runs[-1].text += t
        else:
            runs.append(Run(t, *st))
        last_ch = t[-1]
        prev = c
        kept.append(c)
    if not kept:
        return None
    if runs:
        runs[-1].text = runs[-1].text.rstrip()
    runs = [r for r in runs if r.text]
    if not runs:
        return None
    sizes = Counter(round(float(c["size"]) * 2) / 2 for c in kept)
    size = max(sizes.items(), key=lambda kv: (kv[1], kv[0]))[0]
    tops = sorted(c["top"] for c in kept)
    bots = sorted(c["bottom"] for c in kept)
    # 首个不可拆分单元宽度
    fx0 = kept[0]["x0"]
    end = kept[0]["x1"]
    for a, b in zip([None] + kept, kept):
        if a is not None and b["x0"] - a["x1"] > 0.22 * size:
            break
        end = b["x1"]
        if is_cjk(b["text"][0]):
            break
    return Line(
        runs=runs, x0=min(c["x0"] for c in kept), x1=max(c["x1"] for c in kept),
        top=tops[len(tops) // 2], bottom=bots[len(bots) // 2], y0=tops[0], y1=bots[-1],
        size=size, first_tok_w=max(end - fx0, 0.1), chars=kept,
    )


def build_lines(chars: List[dict], cfg: Config) -> List[Line]:
    seen, uniq = set(), []
    for c in chars:  # 去重 (伪粗体常把同一字符叠印两次)
        k = (c["text"], round(c["x0"]), round(c["top"]))
        if k not in seen:
            seen.add(k)
            uniq.append(c)
    uniq.sort(key=lambda c: (c["top"] + c["bottom"]) / 2)
    groups: List[list] = []
    for c in uniq:
        h = c["bottom"] - c["top"]
        if groups:
            g = groups[-1]
            ov = min(g[1], c["bottom"]) - max(g[0], c["top"])
            if ov >= 0.5 * min(h, g[1] - g[0]):
                g[0], g[1] = min(g[0], c["top"]), max(g[1], c["bottom"])
                g[2].append(c)
                continue
        groups.append([c["top"], c["bottom"], [c]])
    lines = []
    for g in groups:
        ln = make_line(g[2], cfg)
        if ln:
            lines.append(ln)
    lines.sort(key=lambda l: (l.top, l.x0))
    return lines


def lines_to_runs(lines: List[Line]) -> List[Run]:
    """把同一段落的多行拼成 runs：英文去连字符/补空格，中文直接相连。"""
    out: List[Run] = []
    for i, ln in enumerate(lines):
        rs = [replace(r) for r in ln.runs]
        if i > 0 and out and rs:
            pt = out[-1].text
            a, b = pt[-1:], rs[0].text[:1]
            if a == "-" and len(pt) > 1 and pt[-2].isalpha() and b.isalpha() and b.islower():
                out[-1].text = pt[:-1]
            elif a == " " or is_cjk(a) or is_cjk(b):
                pass
            else:
                out[-1].text = pt + " "
        out.extend(rs)
    merged: List[Run] = []
    for r in out:
        if merged and merged[-1].style() == r.style():
            merged[-1].text += r.text
        else:
            merged.append(r)
    return [r for r in merged if r.text]


def _continues(prev: Line, nxt: Line, L: float, R: float, pitch: float, n_cur: int) -> bool:
    """nxt 是否仍属于 prev 所在段落 (即 PDF 里的软换行)。"""
    sz = max(prev.size, nxt.size)
    if abs(prev.size - nxt.size) > 1.0:
        return False
    if prev.bold_all != nxt.bold_all:
        return False
    if _LIST_RE.match(nxt.text):
        return False
    if nxt.top - prev.top > pitch * 1.4 + 1.0:           # 段间距明显更大
        return False
    in_p, in_n = prev.x0 - L, nxt.x0 - L
    if not (n_cur == 1 and _LIST_RE.match(prev.text)):    # 列表项允许悬挂缩进
        if in_n - in_p > 0.8 * sz:                        # 下一行首行缩进
            return False
        if n_cur > 1 and in_p - in_n > 0.8 * sz:
            return False
    # 关键：上一行右侧剩余空间若放得下下一行的首个词，则这里是"硬换行"
    if R - prev.x1 > nxt.first_tok_w + 0.35 * sz + 1.0:
        return False
    return True


def group_paragraphs(lines: List[Line], box: Optional[Tuple[float, float]] = None,
                     in_cell: bool = False) -> List[ParaData]:
    lines = sorted(lines, key=lambda l: (l.top, l.x0))
    if not lines:
        return []
    n = len(lines)
    x0s = sorted(l.x0 for l in lines)
    x1s = sorted(l.x1 for l in lines)
    L = min(x0s) if in_cell else x0s[int(0.1 * (n - 1))]
    R = x1s[min(n - 1, int(math.ceil(0.85 * (n - 1))))]
    box_L, box_R = box if box else (L, max(x1s))
    if in_cell:
        pad = min(max(L - box_L, 1.0), 8.0)
        R = max(max(x1s), box_R - pad)
    pitches = [b.top - a.top for a, b in zip(lines, lines[1:])
               if abs(a.size - b.size) <= 1 and 0 < b.top - a.top < 3 * max(a.size, b.size)]
    pitch = sorted(pitches)[int(0.3 * (len(pitches) - 1))] if pitches else 1.2 * lines[0].size
    paras: List[ParaData] = []
    cur = [lines[0]]
    prev_first_top = lines[0].top

    def finish(ls: List[Line], sb: float) -> ParaData:
        return ParaData(ls, L, R, box_L, box_R, pitch, space_before=sb, in_cell=in_cell)

    sb = 0.0
    for prev, nxt in zip(lines, lines[1:]):
        if _continues(prev, nxt, L, R, pitch, len(cur)):
            cur.append(nxt)
        else:
            paras.append(finish(cur, sb))
            sb = min(max(nxt.top - prev.top - pitch, 0.0), 60.0)
            cur = [nxt]
    paras.append(finish(cur, sb))
    return paras


def decide_align(p: ParaData) -> str:
    ls = p.lines
    sz = ls[0].size
    bl, br = p.box_L, p.box_R
    width = max(br - bl, 1.0)
    cc = (bl + br) / 2
    tol = max(0.6 * sz, 0.03 * width)

    def centered(l: Line) -> bool:
        return abs((l.x0 + l.x1) / 2 - cc) <= tol and (l.x0 - bl) > 1.5 * sz and (br - l.x1) > 1.5 * sz

    if len(ls) == 1:
        l = ls[0]
        if centered(l):
            return "center"
        if (p.R - l.x1) < 0.6 * sz and (l.x0 - p.L) > 6 * sz and l.x1 - l.x0 < 0.6 * width:
            return "right"
        return "left"
    if all(centered(l) for l in ls):
        return "center"
    body = ls[:-1]
    flush = sum(1 for l in body if p.R - l.x1 < 0.6 * sz)
    if body and flush >= max(1, int(0.7 * len(body))) and not p.in_cell:
        return "justify"
    return "left"


# ============================================================================
# 图片提取
# ============================================================================


def _png(im: Image.Image) -> bytes:
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


def render_region(page, bbox, dpi: int) -> bytes:
    x0, t, x1, b = bbox
    px0, pt, px1, pb = page.bbox
    box = (max(x0, px0), max(t, pt), min(x1, px1), min(b, pb))
    if box[2] - box[0] < 1 or box[3] - box[1] < 1:
        raise ValueError("empty region")
    im = page.crop(box).to_image(resolution=dpi).original
    return _png(im.convert("RGB"))


def _raw_image_bytes(img: dict) -> Optional[bytes]:
    """尽量直接取原始图像流；遇到掩码/索引色/CMYK 等复杂情况返回 None 交给渲染兜底。"""
    from pdfminer.pdftypes import resolve1

    stream = img.get("stream")
    if stream is None:
        return None
    attrs = stream.attrs
    for k in ("ImageMask", "SMask", "Mask"):
        v = attrs.get(k)
        if v is not None and resolve1(v) not in (None, False):
            return None
    f = resolve1(attrs.get("Filter"))
    if f is None:
        names = []
    elif isinstance(f, list):
        names = [str(getattr(resolve1(x), "name", resolve1(x))) for x in f]
    else:
        names = [str(getattr(f, "name", f))]
    last = names[-1] if names else ""
    data = stream.get_data()
    w, h = int(resolve1(attrs.get("Width"))), int(resolve1(attrs.get("Height")))
    if last in ("DCTDecode", "DCT"):
        im = Image.open(io.BytesIO(data))
        im.load()
        return data if im.mode in ("RGB", "L") else None
    if last in ("JPXDecode", "CCITTFaxDecode", "CCF", "JBIG2Decode"):
        return None
    if int(resolve1(attrs.get("BitsPerComponent", 8)) or 8) != 8:
        return None
    cs = resolve1(attrs.get("ColorSpace"))
    first = cs[0] if isinstance(cs, (list, tuple)) and cs else cs
    name = str(getattr(resolve1(first), "name", resolve1(first)))
    if name in ("Indexed", "I", "Separation", "DeviceN", "Lab", "Pattern"):
        return None
    n = len(data)
    if n == w * h * 3:
        mode = "RGB"
    elif n == w * h:
        mode = "L"
    else:
        return None
    return _png(Image.frombytes(mode, (w, h), data))


def extract_image(page, img: dict, cfg: Config, force_render: bool = False) -> Optional[ImageData]:
    bbox = (img["x0"], img["top"], img["x1"], img["bottom"])
    w, h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    data = None
    if not force_render:
        try:
            data = _raw_image_bytes(img)
        except Exception as e:  # noqa
            log.debug("raw image extract failed: %s", e)
    if data is None:
        try:
            data = render_region(page, bbox, cfg.image_dpi)
        except Exception as e:  # noqa
            log.warning("image render failed: %s", e)
            return None
    return ImageData(data, w, h)


# ============================================================================
# 表格 / 矢量图
# ============================================================================


def build_table(page, tbl, chars: List[dict], images: List[Tuple[tuple, dict]],
                rects: List[dict], cfg: Config) -> Optional[TableData]:
    cells_bb = [c for c in tbl.cells if c]
    if not cells_bb:
        return None
    xs = cluster_values([v for c in cells_bb for v in (c[0], c[2])], 2.0)
    ys = cluster_values([v for c in cells_bb for v in (c[1], c[3])], 2.0)
    n_r, n_c = len(ys) - 1, len(xs) - 1
    if n_r < 1 or n_c < 1 or n_r * n_c > cfg.max_table_cells:
        return None
    occupied = [[False] * n_c for _ in range(n_r)]
    cells: List[CellData] = []
    tb = tbl.bbox
    tchars = [c for c in chars if center_in(c["x0"], c["top"], c["x1"], c["bottom"], tb, 1)]
    trects = [r for r in rects if r.get("fill") and center_in(r["x0"], r["top"], r["x1"], r["bottom"], tb, 1)]
    for bb in sorted(cells_bb, key=lambda b: (b[1], b[0])):
        c0, c1 = nearest_idx(xs, bb[0]), nearest_idx(xs, bb[2])
        r0, r1 = nearest_idx(ys, bb[1]), nearest_idx(ys, bb[3])
        if c1 <= c0 or r1 <= r0:
            continue
        if any(occupied[r][c] for r in range(r0, r1) for c in range(c0, c1)):
            continue
        for r in range(r0, r1):
            for c in range(c0, c1):
                occupied[r][c] = True
        cell = CellData(r0, r1, c0, c1, box=tuple(bb))
        cch = [c for c in tchars if center_in(c["x0"], c["top"], c["x1"], c["bottom"], bb, 0.5)]
        if cch:
            cell.paras = group_paragraphs(build_lines(cch, cfg), box=(bb[0], bb[2]), in_cell=True)
        for ibox, img in images:
            if center_in(ibox[0], ibox[1], ibox[2], ibox[3], bb, 0):
                d = extract_image(page, img, cfg)
                if d:
                    cell.images.append((ibox[1], d, ibox[0]))
        # 底色：取包含该单元格中心、大小与单元格相近的最小填充矩形
        best, best_area = None, 1e18
        for r in trects:
            if r["x0"] - 2 <= bb[0] + 1 and r["x1"] + 2 >= bb[2] - 1 and r["top"] - 2 <= bb[1] + 1 \
                    and r["bottom"] + 2 >= bb[3] - 1:
                area = (r["x1"] - r["x0"]) * (r["bottom"] - r["top"])
                cell_area = (bb[2] - bb[0]) * (bb[3] - bb[1])
                if area < best_area and area <= cell_area * 1.3 + 4:
                    best, best_area = r, area
        if best is not None:
            rgb = rgb_of(best.get("non_stroking_color"))
            if rgb and min(rgb) < 250:
                cell.shade = "%02X%02X%02X" % rgb
        cells.append(cell)
    if not cells:
        return None
    row_h = [max(ys[i + 1] - ys[i], 1.0) for i in range(n_r)]
    return TableData(xs=xs, row_h=row_h, cells=cells)


def find_figures(page, table_boxes, cfg: Config) -> List[Tuple[float, float, float, float]]:
    """把曲线/色块密集的区域(图表、Logo、轮廓文字等)聚类成"矢量图"区域。"""
    W, H = float(page.width), float(page.height)
    chars = page.chars          # pdfplumber 每访问一次都会重新筛一遍对象，先取一份
    items = []  # [x0, top, x1, bottom, n_curve, n_fillrect]
    for kind, objs in (("curve", page.curves), ("rect", page.rects), ("line", page.lines)):
        for o in objs:
            w, h = o["x1"] - o["x0"], o["bottom"] - o["top"]
            if any(center_in(o["x0"], o["top"], o["x1"], o["bottom"], tb, 1) for tb in table_boxes):
                continue
            if w * h > 0.6 * W * H or (w < 1.5 and h < 1.5):
                continue
            if kind != "curve" and (w < 1.5 or h < 1.5):  # 细线(下划线/分隔线)不参与
                continue
            items.append([o["x0"], o["top"], o["x1"], o["bottom"],
                          1 if kind == "curve" else 0,
                          1 if (kind == "rect" and o.get("fill")) else 0])
    if len(items) < 4 or len(items) > 2500:
        return []
    pad = 4.0
    changed = True
    while changed:
        changed = False
        out: List[list] = []
        for b in items:
            for o in out:
                if b[0] - pad <= o[2] and o[0] - pad <= b[2] and b[1] - pad <= o[3] and o[1] - pad <= b[3]:
                    o[0], o[1], o[2], o[3] = min(o[0], b[0]), min(o[1], b[1]), max(o[2], b[2]), max(o[3], b[3])
                    o[4] += b[4]
                    o[5] += b[5]
                    changed = True
                    break
            else:
                out.append(list(b))
        items = out
    figs = []
    for x0, t, x1, b, ncurve, nfill in items:
        w, h = x1 - x0, b - t
        if w < 30 or h < 30 or w * h > 0.7 * W * H:
            continue
        n_chars = sum(1 for c in chars if center_in(c["x0"], c["top"], c["x1"], c["bottom"], (x0, t, x1, b), 0))
        if ncurve >= 3 or (nfill >= 5 and n_chars <= 80):
            figs.append((max(x0 - 2, 0), max(t - 2, 0), min(x1 + 2, W), min(b + 2, H)))
    return figs


def _grow_figures(figs, chars, cfg: Config, W: float, H: float, pad: float = 12.0):
    """图表周围的坐标轴刻度/标题/图例是独立文字，需并入图片区域，否则会在正文里重复出现。"""
    short = [l for l in build_lines(chars, cfg) if len(l.text) <= 30]
    figs = [list(f) for f in figs]
    for _ in range(5):
        changed = False
        for f in figs:
            for l in short:
                if l.x1 >= f[0] - pad and l.x0 <= f[2] + pad and l.y1 >= f[1] - pad and l.y0 <= f[3] + pad:
                    nf = [min(f[0], l.x0 - 1), min(f[1], l.y0 - 1), max(f[2], l.x1 + 1), max(f[3], l.y1 + 1)]
                    if nf != f:
                        f[:] = nf
                        changed = True
        if not changed:
            break
    return [(max(f[0], 0), max(f[1], 0), min(f[2], W), min(f[3], H)) for f in figs]


# ============================================================================
# 第 1 阶段：解析单页
# ============================================================================


def _collect_rules(page, W: float, H: float) -> List[Tuple[float, float, float, float]]:
    """收集页面上/下 1/4 处的水平细线（页眉下划线、页脚上划线），用于自适应定位页眉页脚区。"""
    out: List[Tuple[float, float, float, float]] = []
    try:
        objs = list(page.lines) + list(page.rects)
    except Exception:  # noqa
        return out
    for o in objs:
        try:
            w = o["x1"] - o["x0"]
            h = o["bottom"] - o["top"]
            top, bottom = float(o["top"]), float(o["bottom"])
        except Exception:
            continue
        if w < 0.45 * W or h > 3.0 or w < 1.0:
            continue
        if bottom <= 0.25 * H or top >= 0.75 * H:
            out.append((top, bottom, float(o["x0"]), float(o["x1"])))
    return out


def parse_page(page, idx: int, cfg: Config) -> PageRaw:
    W, H = float(page.width), float(page.height)
    raw = PageRaw(idx, W, H)
    chars = [c for c in page.chars
             if c.get("upright", True) and c["x1"] > -2 and c["x0"] < W + 2 and c["bottom"] > -2 and c["top"] < H + 2]
    try:
        raw.rules = _collect_rules(page, W, H)
    except Exception as e:  # noqa
        log.debug("rule detect failed: %s", e)

    # --- 表格 ---
    tables: List[Tuple[tuple, TableData]] = []
    try:
        found = sorted(page.find_tables(), key=lambda t: -(t.bbox[2] - t.bbox[0]) * (t.bbox[3] - t.bbox[1]))
        kept = []
        for t in found:
            b = t.bbox
            area = (b[2] - b[0]) * (b[3] - b[1])
            if area > 0.85 * W * H and len(t.cells) <= 2:   # 页面边框，不是表格
                continue
            if any(b[0] >= k[0] - 2 and b[1] >= k[1] - 2 and b[2] <= k[2] + 2 and b[3] <= k[3] + 2 for k in kept):
                continue
            kept.append(b)
            tables.append((b, t))
    except Exception as e:  # noqa
        log.warning("page %d: table detection failed: %s", idx + 1, e)
        tables = []
    table_boxes = [b for b, _ in tables]

    # --- 矢量图 ---
    figs: List[tuple] = []
    if cfg.render_figures:
        try:
            figs = find_figures(page, table_boxes, cfg)
        except Exception as e:  # noqa
            log.debug("figure detection failed: %s", e)

    if figs:
        try:
            figs = _grow_figures(figs, [c for c in chars if not any(
                center_in(c["x0"], c["top"], c["x1"], c["bottom"], tb, 1) for tb in table_boxes)], cfg, W, H)
        except Exception as e:  # noqa
            log.debug("grow figures failed: %s", e)

    # --- 图片 ---
    page_images = []
    for im in page.images:
        try:
            box = (im["x0"], im["top"], im["x1"], im["bottom"])
            w, h = box[2] - box[0], box[3] - box[1]
            if w < 3 or h < 3:
                continue
            if any(center_in(*box, f, 0) for f in figs):
                continue
            if w * h > 0.9 * W * H and len(chars) >= 30:  # 整页背景图/OCR 扫描件文字层：以文字为准
                continue
            page_images.append((box, im))
        except Exception:
            continue

    # 嵌在大图里的图例/局部放大图：不能当成并列的第二栏，丢弃它，并让外层图改为区域渲染，
    # 保证叠加在图上的内容也出现在最终图片里。
    render_idx: Set[int] = set()
    if len(page_images) > 1:
        drop: Set[int] = set()
        for i, (box_i, _im) in enumerate(page_images):
            for j, (box_j, _im2) in enumerate(page_images):
                if i == j or not _box_contains(box_j, box_i):
                    continue
                area_i = max(box_i[2] - box_i[0], 0) * max(box_i[3] - box_i[1], 0)
                area_j = max(box_j[2] - box_j[0], 0) * max(box_j[3] - box_j[1], 0)
                if area_i < 0.9 * area_j:      # 内嵌图明显更小 -> 视为叠加元素
                    drop.add(i)
                    render_idx.add(j)
                    break
        if drop:
            log.debug("page %d: %d nested image(s) merged into container", idx + 1, len(drop))
            page_images = [p for k, p in enumerate(page_images) if k not in drop]
            render_idx = {k - sum(1 for d in drop if d < k) for k in render_idx if k not in drop}

    # --- 表格数据 ---
    rects = list(page.rects)
    for b, t in tables:
        try:
            td = build_table(page, t, chars, page_images, rects, cfg)
        except Exception as e:  # noqa
            log.warning("page %d: table build failed: %s", idx + 1, e)
            td = None
        if td is None:
            table_boxes.remove(b)
            continue
        raw.tables.append(Block("table", b[0], b[1], b[2], b[3], idx, td))

    # --- 独立图片 / 矢量图 ---
    for k, (box, im) in enumerate(page_images):
        if any(center_in(*box, tb, 0) for tb in table_boxes):
            continue
        d = extract_image(page, im, cfg, force_render=k in render_idx)
        if d:
            raw.images.append(Block("image", box[0], box[1], box[2], box[3], idx, d))
    for f in figs:
        try:
            d = ImageData(render_region(page, f, cfg.image_dpi), f[2] - f[0], f[3] - f[1])
            raw.images.append(Block("image", f[0], f[1], f[2], f[3], idx, d))
        except Exception as e:  # noqa
            log.debug("figure render failed: %s", e)
            figs.remove(f)

    # --- 文字(排除表格、矢量图区域) ---
    boxes = table_boxes + figs
    body = [c for c in chars if not any(center_in(c["x0"], c["top"], c["x1"], c["bottom"], b, 1) for b in boxes)]
    raw.lines = build_lines(body, cfg)
    return raw


# ============================================================================
# 逐页缓存 (--cache-page / --cache-dir)
# ============================================================================

_CACHE_SUBDIR = "_pdf2word_pages"


def cache_dir_of(cfg: Config, docx_path: str) -> str:
    """--cache-dir 缺省时，放在输出 docx 同目录下的 _pdf2word_pages。"""
    if cfg.cache_dir:
        return os.path.abspath(cfg.cache_dir)
    return os.path.join(os.path.dirname(os.path.abspath(docx_path)), _CACHE_SUBDIR)


def _page_summary(raw: PageRaw) -> Dict[str, Any]:
    texts = [l.text for l in raw.lines]
    n_cells = sum(len(b.data.cells) for b in raw.tables)
    return {
        "page": raw.index + 1,
        "width": round(raw.W, 2),
        "height": round(raw.H, 2),
        "lines": len(raw.lines),
        "chars": sum(len(t) for t in texts),
        "tables": len(raw.tables),
        "cells": n_cells,
        "images": len(raw.images),
        "fallback": raw.fallback is not None,
        "text_preview": "".join(texts)[:200],
    }


def _img_ext(data: bytes) -> str:
    return "jpg" if data[:2] == b"\xff\xd8" else "png"


def _save_image(data: bytes, cache_dir: str, base: str) -> Optional[str]:
    """缓存里的图片单独落成 jpg/png 文件，JSON 只留文件名引用，不把二进制塞进文本。"""
    fn = "%s.%s" % (base, _img_ext(data))
    try:
        with open(os.path.join(cache_dir, fn), "wb") as f:
            f.write(data)
        return fn
    except Exception as e:  # noqa
        log.warning("缓存图片写入失败 %s: %s", fn, e)
        return None


def _run_json(r: Run) -> dict:
    return {"text": r.text, "font": r.font, "size": r.size, "bold": r.bold,
            "italic": r.italic, "color": list(r.color) if r.color else None}


def _line_json(ln: Line) -> dict:
    return {
        "text": ln.text,
        "x0": round(ln.x0, 2), "x1": round(ln.x1, 2),
        "top": round(ln.top, 2), "bottom": round(ln.bottom, 2),
        "size": ln.size, "bold": ln.bold_all,
        "kind": ln.kind, "col": ln.col, "band": ln.band,
        "runs": [_run_json(r) for r in ln.runs],
        "chars": [{"text": c.get("text", ""), "font": c.get("fontname", ""),
                   "size": round(float(c.get("size") or 0), 2),
                   "x0": round(c["x0"], 2), "x1": round(c["x1"], 2),
                   "top": round(c["top"], 2), "bottom": round(c["bottom"], 2)}
                  for c in ln.chars],
    }


def _para_json(p: ParaData) -> dict:
    return {"text": "".join(l.text for l in p.lines), "lines": len(p.lines),
            "heading": p.heading, "space_before": round(p.space_before, 2),
            "L": round(p.L, 2), "R": round(p.R, 2), "pitch": round(p.pitch, 2)}


def _table_json(td: TableData, pi: int, cache_dir: str) -> dict:
    cells = []
    for k, cd in enumerate(td.cells):
        imgs = []
        for j, (top, im, x0) in enumerate(cd.images):
            fn = _save_image(im.data, cache_dir, "page_%04d_cell%02d_img%02d" % (pi, k, j))
            imgs.append({"file": fn, "w": round(im.w, 2), "h": round(im.h, 2),
                         "top": round(top, 2), "x0": round(x0, 2)})
        cells.append({"r0": cd.r0, "r1": cd.r1, "c0": cd.c0, "c1": cd.c1,
                      "text": cd.text(), "shade": cd.shade,
                      "box": [round(v, 2) for v in cd.box],
                      "paras": [_para_json(p) for p in cd.paras],
                      "images": imgs})
    return {"rows": td.n_rows, "cols": len(td.xs) - 1, "header_rows": td.header_rows,
            "xs": [round(v, 2) for v in td.xs],
            "row_h": [round(v, 2) for v in td.row_h],
            "cells": cells}


def page_cache_json(raw: PageRaw, cache_dir: str) -> dict:
    """把一页的解析结果序列化成可直接阅读的 JSON（图片单独存文件）。"""
    pi = raw.index + 1
    data: Dict[str, Any] = {
        "page": pi, "width": round(raw.W, 2), "height": round(raw.H, 2),
        "lines": [_line_json(l) for l in raw.lines],
        "tables": [_table_json(b.data, pi, cache_dir) for b in raw.tables],
        "images": [],
        "fallback": None,
    }
    for j, b in enumerate(raw.images, 1):
        d: ImageData = b.data
        data["images"].append({
            "file": _save_image(d.data, cache_dir, "page_%04d_img%02d" % (pi, j)),
            "w": round(d.w, 2), "h": round(d.h, 2),
            "box": [round(b.x0, 2), round(b.top, 2), round(b.x1, 2), round(b.bottom, 2)]})
    if raw.fallback is not None:
        data["fallback"] = {
            "file": _save_image(raw.fallback.data, cache_dir, "page_%04d_fallback" % pi),
            "w": round(raw.fallback.w, 2), "h": round(raw.fallback.h, 2)}
    return data


def page_cache_text(raw: PageRaw) -> str:
    """同一页的纯文本视图：逐行文本 + 表格单元格文本，便于直接肉眼核对。"""
    out = ["# page %d  W=%.2f H=%.2f  lines=%d tables=%d images=%d fallback=%s"
           % (raw.index + 1, raw.W, raw.H, len(raw.lines), len(raw.tables),
              len(raw.images), raw.fallback is not None), ""]
    for i, ln in enumerate(raw.lines, 1):
        out.append("L%03d | %s" % (i, ln.text))
    for t, b in enumerate(raw.tables, 1):
        td: TableData = b.data
        out.append("")
        out.append("[table %d] %d 行 x %d 列" % (t, td.n_rows, len(td.xs) - 1))
        for cd in td.cells:
            out.append("  r%d-%d c%d-%d | %s" % (cd.r0, cd.r1 - 1, cd.c0, cd.c1 - 1, cd.text()))
    for i, b in enumerate(raw.images, 1):
        out.append("")
        out.append("[image %d] box=(%.1f,%.1f)-(%.1f,%.1f)" % (i, b.x0, b.top, b.x1, b.bottom))
    return "\n".join(out) + "\n"


def save_page_cache(raw: PageRaw, cache_dir: str, manifest: dict) -> Optional[dict]:
    """把单页解析结果写成 page_XXXX.json + page_XXXX.txt（图片另存 jpg/png）。
    失败只告警，不影响主流程。"""
    pi = raw.index + 1
    name = "page_%04d.json" % pi
    txt_name = "page_%04d.txt" % pi
    try:
        os.makedirs(cache_dir, exist_ok=True)
        data = page_cache_json(raw, cache_dir)
        path = os.path.join(cache_dir, name)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        with open(os.path.join(cache_dir, txt_name), "w", encoding="utf-8") as f:
            f.write(page_cache_text(raw))
        entry = _page_summary(raw)
        entry["file"] = name
        entry["text_file"] = txt_name
        entry["bytes"] = os.path.getsize(path)
    except Exception as e:  # noqa
        log.warning("第 %d 页缓存写入失败: %s", pi, e)
        return None
    manifest["pages"].append(entry)
    return entry


def write_cache_manifest(cache_dir: str, manifest: dict) -> Optional[str]:
    path = os.path.join(cache_dir, "manifest.json")
    try:
        os.makedirs(cache_dir, exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return path
    except Exception as e:  # noqa
        log.warning("缓存 manifest 写入失败: %s", e)
        return None


# ============================================================================
# 第 2 阶段：页眉页脚
# ============================================================================


def _split_parts(line: Line) -> List[List[dict]]:
    cs = sorted(line.chars, key=lambda c: c["x0"])
    parts = [[cs[0]]]
    for a, b in zip(cs, cs[1:]):
        if b["x0"] - a["x1"] > max(2.5 * line.size, 18):
            parts.append([b])
        else:
            parts[-1].append(b)
    return parts


def _norm(text: str) -> str:
    """归一化：数字统一成 #，空白压成一个空格。用于跨页比对同一段页眉/页脚。"""
    return re.sub(r"\d+", "#", re.sub(r"\s+", " ", text.strip()))


_KIND_CN = {"header": "页眉", "footer": "页脚", "page_number": "页码"}

# 纯页码形态："5" / "- 5 -" / "（5）" / "第 5 页" / "第 5 页 共 20 页" / "5/20" / "Page 5 of 20" / "V"
_PN_RES = [
    re.compile(r"^[-—–_·\s]*[#\d]{1,4}[-—–_·\s]*$"),
    re.compile(r"^[（(\[【]\s*[#\d]{1,4}\s*[)）\]】]$"),
    re.compile(r"^第\s*[#\d]{1,4}\s*页(?:\s*[，,/／]?\s*共\s*[#\d]{1,4}\s*页)?$"),
    re.compile(r"^共\s*[#\d]{1,4}\s*页(?:\s*第\s*[#\d]{1,4}\s*页)?$"),
    re.compile(r"^[#\d]{1,4}\s*[/／]\s*[#\d]{1,4}$"),
    re.compile(r"^(?:page|p\.)\s*[#\d]{1,4}(?:\s*(?:of|/)\s*[#\d]{1,4})?$", re.I),
    re.compile(r"^[-—–\s]*[IVXLCDM]{2,7}[-—–\s]*$"),
]

# 页码片段（用于从 "某单位95" 这类"单位名+页码"挤在一起的页脚里把页码剥掉）
_PN_TOKEN = (
    r"(?:第\s*\d{1,4}\s*页(?:\s*[，,/／]?\s*共\s*\d{1,4}\s*页)?"
    r"|共\s*\d{1,4}\s*页(?:\s*第\s*\d{1,4}\s*页)?"
    r"|(?:page|p\.)\s*\d{1,4}(?:\s*(?:of|/)\s*\d{1,4})?"
    r"|\d{1,4}\s*[/／]\s*\d{1,4}"
    r"|[（(]\s*\d{1,4}\s*[)）]"
    r"|\d{1,4})"
)
_PN_TAIL_RE = re.compile(r"[\s\-—–_·|]*" + _PN_TOKEN + r"[\s\-—–_·|]*$", re.I)
_PN_HEAD_RE = re.compile(r"^[\s\-—–_·|]*" + _PN_TOKEN + r"[\s\-—–_·|]*", re.I)

# 页眉/页脚里常见的无意义标签，出现即视为噪声（不依赖跨页重复）
_NOISE_LABELS = {
    "机密", "秘密", "机密文件", "注意保密", "商业秘密", "内部资料", "内部文件",
    "内部使用", "仅供内部使用", "请勿外传", "confidential", "draft",
}


def _is_page_number(text: str) -> bool:
    t = _norm(text).strip()
    return bool(t) and any(r.match(t) for r in _PN_RES)


def _strip_pn_edges(text: str) -> str:
    """去掉首尾的页码片段，返回剩下的正文（页码本身就是全部内容时返回空串）。"""
    s = text.strip()
    for _ in range(2):
        prev = s
        s = _PN_TAIL_RE.sub("", s).strip()
        s = _PN_HEAD_RE.sub("", s).strip()
        if s == prev:
            break
    return s


def _contains_page_number(text: str) -> bool:
    return _is_page_number(text) or _strip_pn_edges(text) != text.strip()


def _is_noise_label(text: str) -> bool:
    t = re.sub(r"[\s:：*·\-—–_|（）()【】\[\]]", "", text or "").lower()
    return bool(t) and t in _NOISE_LABELS


def _hf_keys(text: str) -> List[str]:
    """一行片段的比对键：完整归一化文本；若它带着页码，再用去掉页码后的正文做第二个键。
    这样 "中国电建…有限公司 95" 与 "中国电建…有限公司" + "95" 两种排版能归到同一组。"""
    t = re.sub(r"\s+", " ", text.strip())
    if not t:
        return []
    keys = [_norm(t)]
    core = _strip_pn_edges(t)
    if len(core) >= 2 and core != t:
        k = _norm(core)
        if k not in keys:
            keys.append(k)
    return keys


def _hf_bounds(p: PageRaw) -> Tuple[float, float]:
    """页眉区下界 / 页脚区上界。
    优先用页面上/下缘紧贴文字的水平分隔线定位（比固定 12% 更准，可避开正文首行）；
    找不到分隔线时退回页面高度的 12%。"""
    H = p.H
    top_lim, bot_lim = 0.12 * H, 0.88 * H
    if not p.rules:
        return top_lim, bot_lim
    head_b: Optional[float] = None
    foot_t: Optional[float] = None
    for t, b, _x0, _x1 in p.rules:
        if b <= 0.25 * H and (head_b is None or b < head_b):
            for ln in p.lines:
                if b - 2.0 <= ln.y1 <= b + 8.0:   # 文字紧贴横线上方 -> 页眉下划线
                    head_b = b
                    break
        if t >= 0.75 * H and (foot_t is None or t > foot_t):
            for ln in p.lines:
                if t - 2.0 <= ln.y0 <= t + 14.0:  # 文字紧贴横线下方 -> 页脚上划线
                    foot_t = t
                    break
    if head_b is not None:
        top_lim = min(head_b + 1.5, 0.30 * H)
    if foot_t is not None:
        bot_lim = max(foot_t - 1.5, 0.70 * H)
    return top_lim, bot_lim


def _rows_of(items: List[Tuple[float, float, HFPart]]) -> List[List[HFPart]]:
    """把同一水平上的多个片段（左/中/右）合成一行，避免页脚被拆成两段堆叠。"""
    rows: List[List] = []
    for y0, y1, part in sorted(items, key=lambda t: (t[0], t[2].line.x0)):
        for r in rows:
            ov = min(y1, r[1]) - max(y0, r[0])
            if ov > 0.5 * min(y1 - y0, r[1] - r[0]):
                r[0], r[1] = min(r[0], y0), max(r[1], y1)
                r[2].append(part)
                break
        else:
            rows.append([y0, y1, [part]])
    return [sorted(r[2], key=lambda x: x.line.x0) for r in rows]


def strip_headers_footers(pages: List[PageRaw], cfg: Config):
    """把跨页重复的页眉/页脚/页码从正文行里摘出来。

    返回 (Word 页眉页脚内容, 逐页检测报告)；同时就地修改 pages[*].lines。
    判定依据：同一归一化文本在 >=50% 页面（纯页码 >=35%）的页眉/页脚区重复出现；
    带页码的片段额外用「去掉页码后的正文」做键，兼容 "某单位 95" 与 "某单位"/"95" 两种排版。
    """
    result: Dict[str, List[List[HFPart]]] = {"head": [], "foot": []}
    report: Dict[str, Any] = {
        "pages": [], "threshold": 0, "detected_pages": 0,
        "summary": {"header": [], "footer": [], "page_number": []},
    }
    n = len(pages)
    if not pages:
        return result, report

    # 1) 逐页收集候选片段：落在页眉/页脚区的行，按大间隔切成左/中/右多个 part
    cands: List[List[Tuple[str, Line]]] = []
    for p in pages:
        top_lim, bot_lim = _hf_bounds(p)
        items: List[Tuple[str, Line]] = []
        for ln in p.lines:
            if ln.y1 <= top_lim:
                zone = "head"
            elif ln.y0 >= bot_lim:
                zone = "foot"
            else:
                continue
            for pc in _split_parts(ln):
                pl = make_line(pc, cfg)
                if pl is not None and pl.text.strip():
                    items.append((zone, pl))
        cands.append(items)

    # 2) 跨页统计
    thr = max(2, math.ceil(0.5 * n)) if n > 1 else 1
    thr_pn = max(2, math.ceil(0.35 * n)) if n > 1 else 1
    report["threshold"] = thr
    pages_of: Dict[str, Set[int]] = defaultdict(set)
    text_of: Dict[str, str] = {}
    variants: Dict[str, Set[str]] = defaultdict(set)
    for i, items in enumerate(cands):
        for zone, pl in items:
            for k in _hf_keys(pl.text):
                kk = zone + "\x00" + k
                pages_of[kk].add(i)
                text_of.setdefault(kk, pl.text.strip())
            variants[zone + "\x00" + _norm(pl.text)].add(pl.text.strip())

    accepted: Set[str] = set()
    for kk, ps in pages_of.items():
        zone, key = kk.split("\x00", 1)
        txt = text_of[kk]
        if _is_noise_label(txt):        # 机密/内部资料 之类，单页也删
            accepted.add(kk)
            continue
        pure_pn = _is_page_number(txt)
        if n < 2:                        # 单页无从比对，只认纯页码
            if pure_pn:
                accepted.add(kk)
            continue
        if key and len(ps) >= (thr_pn if pure_pn else thr):
            accepted.add(kk)

    # 3) 按页摘除 + 记录报告（accepted 为空时仍逐页留空记录，便于确认“确实没检测到”）
    best: Dict[str, Tuple[int, List[List[HFPart]]]] = {"head": (0, []), "foot": (0, [])}
    for i, p in enumerate(pages):
        top_lim, bot_lim = _hf_bounds(p)
        new_lines: List[Line] = []
        removed: Dict[str, List[Tuple[float, float, HFPart]]] = {"head": [], "foot": []}
        page_items: List[HFItem] = []
        for ln in p.lines:
            if ln.y1 <= top_lim:
                zone = "head"
            elif ln.y0 >= bot_lim:
                zone = "foot"
            else:
                new_lines.append(ln)
                continue
            parts = [make_line(pc, cfg) for pc in _split_parts(ln)]
            parts = [x for x in parts if x is not None and x.text.strip()]
            if not parts:
                new_lines.append(ln)
                continue
            keep: List[Line] = []
            hit: List[Line] = []
            for x in parts:
                keys = [zone + "\x00" + k for k in _hf_keys(x.text)]
                (hit if any(k in accepted for k in keys) else keep).append(x)
            if not hit:
                new_lines.append(ln)
                continue
            new_lines.extend(keep)
            for x in hit:
                is_field = len(variants.get(zone + "\x00" + _norm(x.text), ())) > 1
                removed[zone].append((x.y0, x.y1, HFPart(x, is_field)))
                page_items.append(HFItem(i, "page_number" if _is_page_number(x.text)
                                        else ("header" if zone == "head" else "footer"),
                                        x.text.strip(), round(x.y0, 2), round(x.y1, 2),
                                        round(x.x0, 2), round(x.x1, 2), _contains_page_number(x.text)))
        p.lines = new_lines
        for zone in ("head", "foot"):
            rows = _rows_of(removed[zone])
            cnt = sum(len(r) for r in rows)
            if cnt > best[zone][0]:
                best[zone] = (cnt, rows)
        report["pages"].append({
            "page": p.index + 1,
            "items": [{"type": it.kind, "text": it.text, "y0": it.y0, "y1": it.y1,
                       "x0": it.x0, "x1": it.x1,
                       "contains_page_number": it.contains_page_number}
                      for it in sorted(page_items, key=lambda t: (t.y0, t.x0))],
        })

    for entry in report["pages"]:
        for it in entry["items"]:
            lst = report["summary"][it["type"]]
            if it["text"] not in lst:
                lst.append(it["text"])
    report["detected_pages"] = sum(1 for e in report["pages"] if e["items"])
    return {"head": best["head"][1], "foot": best["foot"][1]}, report


# ============================================================================
# 第 3 阶段：分栏、段落、排序
# ============================================================================


def detect_gutter(lines: List[Line]) -> Optional[float]:
    if len(lines) < 8:
        return None
    cx0, cx1 = min(l.x0 for l in lines), max(l.x1 for l in lines)
    cw = cx1 - cx0
    if cw < 150:
        return None
    cover = [0] * (int(cw) + 3)
    for l in lines:
        s = set()
        for c in l.chars:
            s.update(range(max(int(c["x0"] - cx0), 0), min(int(c["x1"] - cx0) + 1, len(cover))))
        for i in s:
            cover[i] += 1
    thr = max(1, int(0.1 * len(lines)))
    lo, hi = int(0.25 * cw), int(0.75 * cw)
    best, run_s = (0, 0), None
    for i in range(lo, hi + 1):
        if cover[i] <= thr:
            run_s = i if run_s is None else run_s
            if i - run_s + 1 > best[1] - best[0]:
                best = (run_s, i + 1)
        else:
            run_s = None
    if best[1] - best[0] < 10:
        return None
    xg = cx0 + (best[0] + best[1]) / 2
    left, right, cross = [], [], 0
    for l in lines:
        lc = [c for c in l.chars if c["x1"] <= xg]
        rc = [c for c in l.chars if c["x0"] >= xg]
        if len(lc) + len(rc) < len(l.chars):
            cross += 1
            continue
        if lc:
            left.append(max(c["x1"] for c in lc) - min(c["x0"] for c in lc))
        if rc:
            right.append(max(c["x1"] for c in rc) - min(c["x0"] for c in rc))
    if cross > 0.15 * len(lines) or len(left) < 4 or len(right) < 4:
        return None
    med = lambda v: sorted(v)[len(v) // 2]  # noqa
    if med(left) < 0.3 * (xg - cx0) or med(right) < 0.3 * (cx1 - xg):
        return None
    return xg


def can_merge_paras(a: Block, b: Block) -> bool:
    pa, pb = a.data, b.data
    la, fb = pa.lines[-1], pb.lines[0]
    if pa.heading or pb.heading or abs(la.size - fb.size) > 1:
        return False
    if _LIST_RE.match(fb.text) or la.bold_all != fb.bold_all:
        return False
    if fb.x0 - pb.L > 0.8 * fb.size:
        return False
    return pa.R - la.x1 <= fb.first_tok_w + 0.35 * fb.size + 1.0


def _box_contains(outer, inner, tol: float = 1.5) -> bool:
    """inner 是否基本落在 outer 内部（图例/局部图嵌在大图里）。"""
    return (inner[0] >= outer[0] - tol and inner[1] >= outer[1] - tol
            and inner[2] <= outer[2] + tol and inner[3] <= outer[3] + tol)


def _img_row_span(row: List[Block]) -> Tuple[float, float]:
    return min(b.top for b in row), max(b.bottom for b in row)


def _text_between(lines: List[Line], y0: float, y1: float, x0: float, x1: float) -> bool:
    """两行图片之间的垂直缝隙里是否有正文（图注/说明），有则不能并成一张表。"""
    if y1 - y0 <= 0.5:
        return False
    for ln in lines:
        if ln.bottom <= y0 + 0.5 or ln.top >= y1 - 0.5:
            continue
        if min(ln.x1, x1) - max(ln.x0, x0) > 1.0:
            return True
    return False


def group_image_blocks(imgs: List[Block], lines: List[Line], W: float) -> List[Block]:
    """把同一水平带上的并排图片合成 imggrid 块。

    原始实现里每张图片都是一个独立 block，写 docx 时各占一段、按 top 从上往下堆，
    于是「左右两栏」被抖成上下两张、「四象限」被抖成一列。这里按 top/bottom 重合度
    分行，行内再按横向间距切段，最后把列边界对齐的相邻行合成一张网格表，
    交给 write_image_grid 用无边框表格还原排布。
    """
    if len(imgs) < 2:
        return list(imgs)
    # 安全兜底：仍被其它图包住的小图（图例/局部图）不参与成行
    kept: List[Block] = []
    for b in sorted(imgs, key=lambda x: -((x.x1 - x.x0) * (x.bottom - x.top))):
        if not any(_box_contains((o.x0, o.top, o.x1, o.bottom), (b.x0, b.top, b.x1, b.bottom))
                   and (o.x1 - o.x0) * (o.bottom - o.top) > 0.9 * (b.x1 - b.x0) * (b.bottom - b.top)
                   for o in kept):
            kept.append(b)
    imgs = kept
    if len(imgs) < 2:
        return list(imgs)
    items = sorted(imgs, key=lambda b: (b.top, b.x0))
    rows: List[List[Block]] = []
    for b in items:
        best, best_ov = None, 0.0
        for row in rows:
            rtop, rbot = _img_row_span(row)
            ov = min(rbot, b.bottom) - max(rtop, b.top)
            h = min(rbot - rtop, b.bottom - b.top)
            if h > 0 and ov > 0 and ov / h > best_ov:
                best, best_ov = row, ov / h
        if best is not None and best_ov >= 0.4:   # 垂直方向明显重叠 -> 同行
            best.append(b)
        else:
            rows.append([b])

    rows.sort(key=lambda r: (_img_row_span(r)[0], min(b.x0 for b in r)))
    # 行内按 x 排序；横向空隙过大（分属两栏的独立插图）时切成两段
    seg_rows: List[List[Block]] = []
    for row in rows:
        row.sort(key=lambda b: b.x0)
        cur = [row[0]]
        for b in row[1:]:
            if b.x0 - cur[-1].x1 > max(0.18 * W, 40.0):
                seg_rows.append(cur)
                cur = [b]
            else:
                cur.append(b)
        seg_rows.append(cur)

    tol = max(4.0, 0.02 * W)
    groups: List[List[List[Block]]] = []      # 每个网格 = 若干行，每行若干图片
    for row in seg_rows:
        if groups:
            prev = groups[-1][-1]
            prtop, prbot = _img_row_span(prev)
            rtop, rbot = _img_row_span(row)
            vgap = rtop - prbot
            if (0 <= vgap <= max(12.0, 0.3 * min(prbot - prtop, rbot - rtop))
                    and len(prev) == len(row)
                    and all(abs(p.x0 - q.x0) <= tol and abs(p.x1 - q.x1) <= tol
                            for p, q in zip(prev, row))
                    and not _text_between(lines, prbot, rtop,
                                          min(p.x0 for p in prev), max(p.x1 for p in prev))):
                groups[-1].append(row)      # 上下两排图列边界对齐 -> 同一张 2x2 网格
                continue
        groups.append([row])

    out: List[Block] = []
    for group in groups:
        cells = [b for row in group for b in row]
        if len(cells) == 1:
            out.append(cells[0])
            continue
        g = ImageGrid(group, min(b.x0 for b in cells), min(b.top for b in cells),
                      max(b.x1 for b in cells), max(b.bottom for b in cells))
        out.append(Block("imggrid", g.x0, g.top, g.x1, g.bottom, cells[0].page, g))
    return out


def build_page_blocks(raw: PageRaw, cfg: Config, body_size: float, levels: Dict[float, int]) -> List[Block]:
    blocks: List[Block] = []
    if raw.fallback is not None:
        d = raw.fallback
        return [Block("image", 0, 0, raw.W, raw.H, raw.index, d)]

    lines = raw.lines
    gutter = detect_gutter(lines) if cfg.detect_columns else None
    img_blocks = group_image_blocks(raw.images, lines, raw.W)
    items = []  # (top, is_full, obj)
    for l in lines:
        if gutter is not None:
            lc = [c for c in l.chars if c["x1"] <= gutter]
            rc = [c for c in l.chars if c["x0"] >= gutter]
            if lc and rc and len(lc) + len(rc) == len(l.chars):
                for part, kind in ((lc, "left"), (rc, "right")):
                    pl = make_line(part, cfg)
                    if pl:
                        pl.kind = kind
                        items.append((pl.top, False, pl))
                continue
            if len(lc) == len(l.chars):
                l.kind = "left"
            elif len(rc) == len(l.chars):
                l.kind = "right"
        items.append((l.top, l.kind == "full", l))
    for b in raw.tables + img_blocks:
        full = gutter is None or (b.x0 < gutter - 2 and b.x1 > gutter + 2)
        if not full:
            b.col = 1 if (b.x0 + b.x1) / 2 < gutter else 2
        items.append((b.top, full, b))
    items.sort(key=lambda t: t[0])
    band, prev_full = 0, None
    for top, full, obj in items:
        if prev_full is not None and full != prev_full:
            band += 1
        prev_full = full
        obj.band = band
        if gutter is not None and not full and isinstance(obj, Line):
            obj.col = 1 if obj.kind == "left" else 2
        elif isinstance(obj, Block) and not full:
            pass
        else:
            obj.col = 0

    groups: Dict[Tuple[int, int], List[Line]] = defaultdict(list)
    for _, _, obj in items:
        if isinstance(obj, Line):
            groups[(obj.band, obj.col)].append(obj)
    for (bd, cl), ls in groups.items():
        for p in group_paragraphs(ls):
            if cfg.detect_headings and len(p.lines) <= 3 and p.lines[0].size in levels \
                    and sum(len(l.text) for l in p.lines) < 150:
                p.heading = levels[p.lines[0].size]
            blocks.append(Block(
                "para", min(l.x0 for l in p.lines), min(l.y0 for l in p.lines),
                max(l.x1 for l in p.lines), max(l.y1 for l in p.lines), raw.index, p, band=bd, col=cl))
            for l in p.lines:
                l.chars = []
    blocks.extend(raw.tables + img_blocks)
    blocks.sort(key=lambda b: (b.band, b.col, b.top))
    for i, b in enumerate(blocks):
        if i == 0:
            b.page_start = True
        else:
            pb = blocks[i - 1]
            if (pb.band, pb.col) == (b.band, b.col) or b.col == 0 or pb.col == 0:
                b.gap_before = min(max(b.top - pb.bottom, 0.0), 60.0)
    return blocks


# ============================================================================
# 第 4 阶段：跨页合并
# ============================================================================


def _row_sig(cells: List[CellData], r: int):
    return tuple(sorted((c.c0, c.c1, c.text()) for c in cells if c.r0 == r))


def merge_tables(a: TableData, b: TableData, tol: float) -> bool:
    """尝试把 b (下一页的表格) 并入 a。列边界需对齐；并去掉重复的表头行。"""
    if abs(a.xs[0] - b.xs[0]) > tol or abs(a.xs[-1] - b.xs[-1]) > tol:
        return False
    mp = []
    for x in b.xs:
        j = nearest_idx(a.xs, x)
        if abs(a.xs[j] - x) > tol:
            return False
        mp.append(j)
    if any(m2 <= m1 for m1, m2 in zip(mp, mp[1:])):
        return False
    bm = [replace(c, c0=mp[c.c0], c1=mp[c.c1]) for c in b.cells]
    skip = 0
    for i in range(min(3, b.n_rows - 1, a.n_rows)):
        sa, sb = _row_sig(a.cells, i), _row_sig(bm, i)
        if sa == sb and any(t[2] for t in sa):
            skip = i + 1
        else:
            break
    if any(c.r0 < skip < c.r1 for c in bm):
        skip = 0
    off = a.n_rows - skip
    for c in bm:
        if c.r0 >= skip:
            a.cells.append(replace(c, r0=c.r0 + off, r1=c.r1 + off))
    a.row_h.extend(b.row_h[skip:])
    if skip:
        a.header_rows = max(a.header_rows, skip)
    return True


def merge_across(blocks: List[Block], pages: Dict[int, PageRaw], ref: Tuple[float, float]) -> List[Block]:
    """合并跨页/跨栏被切开的表格与段落。ref=(正文区域顶部, 正文区域底部)"""
    out: List[Block] = []
    for b in blocks:
        if out:
            a = out[-1]
            next_page = b.page_start and a.last_page == b.page - 1
            next_col = a.page == b.page and a.band == b.band and a.col == 1 and b.col == 2
            if (next_page or next_col) and a.kind == b.kind:
                H = pages[a.last_page].H
                try:
                    if a.kind == "table":
                        near_bottom = a.bottom >= ref[1] - 0.06 * H
                        near_top = b.top <= ref[0] + 0.06 * H
                        if (next_col or (near_bottom and near_top)) and \
                                merge_tables(a.data, b.data, tol=max(3.0, 0.008 * pages[b.page].W)):
                            a.last_page, a.bottom = b.last_page, b.bottom
                            a.x1 = max(a.x1, b.x1)
                            log.debug("merged cross-page table: p%d+p%d", a.page + 1, b.page + 1)
                            continue
                    elif a.kind == "para" and can_merge_paras(a, b):
                        a.data.lines.extend(b.data.lines)
                        a.last_page, a.bottom = b.last_page, b.bottom
                        log.debug("merged cross-page paragraph: p%d+p%d", a.page + 1, b.page + 1)
                        continue
                except Exception as e:  # noqa
                    log.debug("merge failed: %s", e)
        out.append(b)
    return out


# ============================================================================
# 第 5 阶段：写 docx
# ============================================================================

_ALIGN = {"left": WD_ALIGN_PARAGRAPH.LEFT, "center": WD_ALIGN_PARAGRAPH.CENTER,
          "right": WD_ALIGN_PARAGRAPH.RIGHT, "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}

# doc.styles[name] 每次都要遍历样式表；整份文档里段落/表格上万，缓存一下能省不少
_STYLE_CACHE: Dict[Tuple[int, str], Any] = {}


def get_style(doc, name: str):
    """带缓存的样式查找；找不到返回 None（调用方照旧跳过即可）。

    调用方只需在新建 Document 时清一次缓存（见 write_docx），保证不会拿到别的文档的样式。
    """
    key = (id(doc), name)
    if key in _STYLE_CACHE:
        return _STYLE_CACHE[key]
    try:
        st = doc.styles[name]
    except KeyError:
        st = None
    _STYLE_CACHE[key] = st
    return st


_RFONTS_TAG = qn("w:rFonts")
_RFONT_ATTRS = tuple(qn(a) for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"))


def set_font(run, name: Optional[str]):
    if not name:
        return
    # 直接写 rFonts：run.font.name 只设 ascii/hAnsi，这里还要 eastAsia/cs，
    # 合成一次操作可以少一遍 rPr 查找（整份文档里 run 数量在十万级）。
    rpr = run._r.get_or_add_rPr()
    rfonts = rpr.find(_RFONTS_TAG)
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    for attr in _RFONT_ATTRS:
        rfonts.set(attr, name)


def write_runs(p, runs: List[Run], heading: bool = False):
    for r in runs:
        run = p.add_run(sanitize(r.text))
        run.font.size = Pt(r.size)
        run.bold = True if r.bold else (False if heading else None)
        run.italic = True if r.italic else None
        if r.color:
            run.font.color.rgb = RGBColor(*r.color)
        elif heading:
            run.font.color.rgb = RGBColor(0, 0, 0)
        set_font(run, r.font)


def write_para(doc, p, pd: ParaData, origin: float, space_before: float, space_after: float = 0.0):
    runs = lines_to_runs(pd.lines)
    if not runs:
        return
    if pd.heading:
        st = get_style(doc, "Heading %d" % pd.heading)
        if st is not None:
            p.style = st
    write_runs(p, runs, bool(pd.heading))
    pf = p.paragraph_format
    align = decide_align(pd)
    p.alignment = _ALIGN[align]
    pf.space_before = Pt(max(space_before, 0))
    pf.space_after = Pt(max(space_after, 0))
    size = pd.lines[0].size
    pf.line_spacing = Pt(min(max(pd.pitch, size), 3 * size))
    pf.line_spacing_rule = WD_LINE_SPACING.AT_LEAST
    if align in ("left", "justify"):
        rest = min(l.x0 for l in pd.lines[1:]) if len(pd.lines) > 1 else pd.lines[0].x0
        left = rest - origin
        first = pd.lines[0].x0 - rest
        if left > 1:
            pf.left_indent = Pt(left)
        if abs(first) > 1:
            pf.first_line_indent = Pt(first)
    pf.widow_control = True


def write_picture(p, img: ImageData, max_w: float, max_h: float):
    w, h = img.w, img.h
    scale = min(1.0, max_w / w if w > 0 else 1.0, max_h / h if h > 0 else 1.0)
    w, h = max(w * scale, 1.0), max(h * scale, 1.0)
    p.add_run().add_picture(io.BytesIO(img.data), width=Pt(w), height=Pt(h))


def _spacer(doc, h: float):
    """高度为 h 磅的空段；用于图片网格之间的垂直间隙（Word 要求相邻表格间有段落）。"""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = Pt(min(max(h, 1.0), 200.0))
    pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    r = p.add_run("")
    r.font.size = Pt(1)
    return p


def _row_columns(row: List[Block]) -> List[Tuple[float, float]]:
    """一行图片 -> 每张图各占一列的区间。

    相邻两张图以两者边界的**中线**为界，这样即使图片在 PDF 里轻微重叠或间隔很小，
    也不会出现两个单元格跨到同一列、合并互相打架的情况（旧实现正是这里把两张图
    塞进同一个合并单元格，渲染出来就糊成一团）。
    """
    n = len(row)
    cols: List[Tuple[float, float]] = []
    left = row[0].x0
    for k, b in enumerate(row):
        right = b.x1 if k == n - 1 else (b.x1 + row[k + 1].x0) / 2.0
        if right <= left:
            right = left + 1.0
        cols.append((left, right))
        left = right
    return cols


def write_image_grid(doc, grid: ImageGrid, left_margin: float, text_w: float, max_h: float):
    """用无边框表格把并排图片按原版面排回去（一图一列，行内无合并）。"""
    row0 = grid.rows[0]
    if not row0:
        return
    rects = _row_columns(row0)
    n_cols = len(rects)
    total = rects[-1][1] - rects[0][0]
    scale = min(1.0, text_w / total) if total > 1 else 1.0
    widths = [max((r[1] - r[0]) * scale, 1.0) for r in rects]

    table = doc.add_table(rows=len(grid.rows), cols=n_cols)
    table.autofit = False
    tblPr = table._tbl.tblPr
    # 表格缩进对齐原始左边界；单元格边距归零，避免默认边距吃掉列宽
    ind = (grid.x0 - left_margin) * scale
    if ind > 1:
        el = OxmlElement("w:tblInd")
        el.set(qn("w:w"), str(int(ind * 20)))
        el.set(qn("w:type"), "dxa")
        tblPr.insert_element_before(el, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
                                    "w:tblLook", "w:tblCaption", "w:tblDescription")
    mar = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement("w:%s" % side)
        e.set(qn("w:w"), "0")
        e.set(qn("w:type"), "dxa")
        mar.append(e)
    tblPr.insert_element_before(mar, "w:tblLook", "w:tblCaption", "w:tblDescription")
    for c in range(n_cols):
        try:
            table.columns[c].width = Pt(widths[c])
        except Exception:
            pass

    for ri, row in enumerate(grid.rows):
        for ci, b in enumerate(row[:n_cols]):
            try:
                cell = table.cell(ri, ci)
            except Exception:
                continue
            try:
                cell.width = Pt(widths[ci])
            except Exception:
                pass
            p = cell.paragraphs[0]
            pf = p.paragraph_format
            pf.space_before = Pt(0)
            pf.space_after = Pt(0)
            pf.line_spacing = 1.0
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            try:
                write_picture(p, b.data, max(widths[ci] - 1.5, 4.0), max_h)
            except Exception as e:  # noqa
                log.warning("image grid picture failed: %s", e)

    # 空单元格：压到 1pt，避免把行高撑到默认字号
    fill = {(ri, ci) for ri, row in enumerate(grid.rows) for ci in range(len(row[:n_cols]))}
    for ri in range(len(grid.rows)):
        for ci in range(n_cols):
            if (ri, ci) in fill:
                continue
            try:
                p = table.cell(ri, ci).paragraphs[0]
            except Exception:
                continue
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(1)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            if not p.runs:
                p.add_run("").font.size = Pt(1)

    # 行内不许跨页拆开；最后一行与紧跟的图注保持同页
    for ri, tr in enumerate(table.rows):
        try:
            tr._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        except Exception:
            pass
        if ri == len(table.rows) - 1:
            for cell in tr.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True


def add_field(p, instr: str, size: float):
    r = p.add_run("1")
    r.font.size = Pt(size)
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), instr)
    r._r.addprevious(fld)
    fld.append(r._r)


def fill_header_footer(container, rows: List[List[HFPart]], W: float, ml: float, mr: float, style=None):
    text_w = max(W - ml - mr, 36.0)
    first = True
    for row in rows:
        p = container.paragraphs[0] if first and container.paragraphs else container.add_paragraph()
        first = False
        if style is not None:
            p.style = style  # 默认 Header 样式自带居中/右对齐制表位，会打乱我们的布局
        for r in list(p.runs):
            r._r.getparent().remove(r._r)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)
        row = sorted(row, key=lambda x: x.line.x0)

        def slot(part: HFPart) -> str:
            c = (part.line.x0 + part.line.x1) / 2
            return "left" if c < 0.38 * W else ("right" if c > 0.62 * W else "center")

        slots = [slot(x) for x in row]
        if len(row) == 1:
            p.alignment = _ALIGN[slots[0]]
        stops = [s for s in ("center", "right") if s in slots[1:] or (s in slots and len(row) > 1)]
        if stops:
            ts = p.paragraph_format.tab_stops
            for s in stops:
                ts.add_tab_stop(Pt(text_w / 2 if s == "center" else text_w),
                                WD_TAB_ALIGNMENT.CENTER if s == "center" else WD_TAB_ALIGNMENT.RIGHT)
        cur = -1
        for part, s in zip(row, slots):
            if len(row) > 1 and s != "left":
                tgt = stops.index(s)
                if tgt > cur:
                    p.add_run("\t" * (tgt - cur))
                    cur = tgt
            ln = part.line
            if part.is_field:
                nums = 0
                for seg in re.split(r"(\d+)", ln.text):
                    if not seg:
                        continue
                    if seg.isdigit():
                        add_field(p, "PAGE" if nums == 0 else "NUMPAGES", ln.size)
                        nums += 1
                    else:
                        write_runs(p, [replace(ln.runs[0], text=seg)])
            else:
                write_runs(p, ln.runs)


def _set_cell_shade(cell, hex_fill: str):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _tiny_paragraph(container):
    p = container.add_paragraph()
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = Pt(1)
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    r = p.add_run("")
    r.font.size = Pt(1)
    return p


def _table_grid(table) -> List[Any]:
    """一次性取出整张表的网格单元格。

    python-docx 的 Table._cells 是 property，每访问一次就遍历全部 tc 重建一串 _Cell；
    Table.cell(r, c) 内部也是访问 _cells。原实现在「行 x 列」双重循环里反复调用，
    复杂度变成 O(单元格数^2)，一张几百单元格的表就要几十秒——这是写 docx 阶段的主要瓶颈。
    """
    try:
        return list(table._cells)
    except Exception:                      # 私有 API 变动时的兼容回退
        n_c = len(table.columns)
        return [table.cell(r, c) for r in range(len(table.rows)) for c in range(n_c)]


def write_table(doc, td: TableData, left_margin: float, tbl_x0: float, text_w: float, max_h: float):
    n_r, n_c = td.n_rows, len(td.xs) - 1
    table = doc.add_table(rows=n_r, cols=n_c)
    st = get_style(doc, "Table Grid")
    if st is not None:
        table.style = st
    table.autofit = False
    widths = [max(td.xs[i + 1] - td.xs[i], 2.0) for i in range(n_c)]
    ind = tbl_x0 - left_margin
    if ind > 1:
        tblPr = table._tbl.tblPr
        el = OxmlElement("w:tblInd")
        el.set(qn("w:w"), str(int(ind * 20)))
        el.set(qn("w:type"), "dxa")
        tblPr.insert_element_before(el, "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar",
                                    "w:tblLook", "w:tblCaption", "w:tblDescription")
    # 行高 / 列宽：网格只取一次
    rows = table.rows
    grid = _table_grid(table)
    for r in range(n_r):
        row = rows[r]
        row.height = Pt(td.row_h[r])
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        if r < td.header_rows:
            row._tr.get_or_add_trPr().append(OxmlElement("w:tblHeader"))
        base = r * n_c
        for c in range(n_c):
            grid[base + c].width = Pt(widths[c])
    for c in range(n_c):
        table.columns[c].width = Pt(widths[c])
    # 先合并（合并会改动 tc 结构，之后重新取一次网格）
    merges = [cd for cd in td.cells if cd.r1 - cd.r0 > 1 or cd.c1 - cd.c0 > 1]
    for cd in merges:
        try:
            grid[cd.r0 * n_c + cd.c0].merge(grid[(cd.r1 - 1) * n_c + (cd.c1 - 1)])
        except Exception as e:  # noqa
            log.debug("cell merge failed: %s", e)
    if merges:
        grid = _table_grid(table)
    # 再写内容
    for cd in td.cells:
        try:
            cell = grid[cd.r0 * n_c + cd.c0]
        except Exception:
            continue
        if cd.shade:
            _set_cell_shade(cell, cd.shade)
        origin = min((l.x0 for p in cd.paras for l in p.lines), default=cd.box[0])
        first = True
        elements = [("p", pr.lines[0].top, pr) for pr in cd.paras] + [("i", t, (im, x0)) for t, im, x0 in cd.images]
        elements.sort(key=lambda e: e[1])
        for kind, _, obj in elements:
            p = cell.paragraphs[0] if first else cell.add_paragraph()
            first = False
            try:
                if kind == "p":
                    write_para(doc, p, obj, origin, obj.space_before)
                else:
                    im, x0 = obj
                    write_picture(p, im, max(cd.box[2] - cd.box[0] - 6, 4), max_h)
                    p.paragraph_format.space_after = Pt(0)
            except Exception as e:  # noqa
                log.warning("cell content failed: %s", e)
        if first:  # 空单元格：避免默认 11pt 空段把行撑高
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = Pt(1)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    return table


def set_section(sec, W: float, H: float, ml: float, mr: float, mt: float, mb: float):
    W, H = min(max(W, 72), 1584), min(max(H, 72), 1584)
    sec.orientation = WD_ORIENT.LANDSCAPE if W > H else WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Pt(W), Pt(H)
    sec.left_margin, sec.right_margin = Pt(ml), Pt(mr)
    sec.top_margin, sec.bottom_margin = Pt(mt), Pt(mb)
    sec.header_distance = Pt(min(max(mt * 0.4, 6), 36))
    sec.footer_distance = Pt(min(max(mb * 0.4, 6), 36))


def write_docx(blocks: List[Block], pages: Dict[int, PageRaw], hf: Dict[str, List[List[HFPart]]],
               cfg: Config, out_path: str, meta: dict, extents: List[Tuple[int, float, float, float, float]],
               also_save_without_hf: Optional[str] = None):
    _STYLE_CACHE.clear()
    doc = Document()
    normal = get_style(doc, "Normal")
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.space_before = Pt(0)
    for lvl in (1, 2, 3):
        hs = get_style(doc, "Heading %d" % lvl)
        if hs is not None:
            hs.paragraph_format.keep_with_next = False
            hs.font.color.rgb = RGBColor(0, 0, 0)
    try:
        cp = doc.core_properties
        cp.title = sanitize(str(meta.get("Title") or ""))[:255]
        cp.author = sanitize(str(meta.get("Author") or ""))[:255]
    except Exception:
        pass

    # 按页面尺寸分组算页边距
    def size_key(pi: int):
        pr = pages[pi]
        return (round(pr.W), round(pr.H))

    runs_of_pages: Dict[Tuple[int, int], list] = defaultdict(list)
    for e in extents:  # (page, x0, top, x1, bottom)：合并跨页块之前的每页真实范围
        if e[0] in pages:
            runs_of_pages[size_key(e[0])].append(e)

    def margins(key):
        bs = runs_of_pages.get(key) or [(0, 36.0, 36.0, key[0] - 36.0, key[1] - 36.0)]
        W, H = key
        ml = max(min(e[1] for e in bs), 12.0)
        mr = max(W - max(e[3] for e in bs), 12.0)
        mt = max(min(e[2] for e in bs), 12.0)
        mb = max(H - max(e[4] for e in bs), 12.0)
        if ml + mr > W - 72:
            ml = mr = 36.0
        if mt + mb > H - 72:
            mt = mb = 36.0
        return ml, mr, mt, mb

    cur_key, cur_sec, cur_m = None, None, None
    section_fills: List[Tuple[Any, Tuple[int, int], Tuple[float, float, float, float]]] = []
    prev: Optional[Block] = None
    prev_para = None
    ncount = 0
    total_blocks = len(blocks)
    step = max(1, total_blocks // 10)
    t_write = time.time()
    for i, b in enumerate(blocks):
        if (i + 1) % step == 0:
            log.info("写入 docx: %d/%d 块 (%.0f%%), 已用 %.1fs",
                     i + 1, total_blocks, 100.0 * (i + 1) / max(total_blocks, 1), time.time() - t_write)
        key = size_key(b.page)
        if key != cur_key:
            cur_sec = doc.sections[0] if cur_key is None else doc.add_section(WD_SECTION.NEW_PAGE)
            cur_m = margins(key)
            set_section(cur_sec, key[0], key[1], *cur_m)
            section_fills.append((cur_sec, key, cur_m))
            cur_key = key
            prev = None
        ml, mr, mt, mb = cur_m
        text_w = key[0] - ml - mr
        max_h = key[1] - mt - mb - 4
        try:
            page_break = cfg.keep_page_breaks and b.page_start and prev is not None and cur_key == size_key(prev.page)
            if b.kind == "para":
                pd: ParaData = b.data
                p = doc.add_paragraph()
                if page_break:
                    p.paragraph_format.page_break_before = True
                write_para(doc, p, pd, ml, b.gap_before if prev is not None else 0.0)
                prev_para = p
            elif b.kind == "image":
                p = doc.add_paragraph()
                if page_break:
                    p.paragraph_format.page_break_before = True
                im: ImageData = b.data
                p.paragraph_format.space_before = Pt(b.gap_before if prev is not None else 0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1
                cx = (b.x0 + b.x1) / 2
                if abs(cx - (ml + (key[0] - mr)) / 2) < 0.03 * text_w:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif b.x1 >= key[0] - mr - 2 and b.x0 - ml > 0.3 * text_w:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                elif b.x0 - ml > 1:
                    p.paragraph_format.left_indent = Pt(min(b.x0 - ml, text_w * 0.9))
                write_picture(p, im, text_w - 1, max_h)
                prev_para = p
            elif b.kind == "imggrid":
                sp = _spacer(doc, b.gap_before if prev is not None else 0)
                if page_break:
                    sp.paragraph_format.page_break_before = True
                write_image_grid(doc, b.data, ml, text_w, max_h)
                prev_para = None
            else:
                td: TableData = b.data
                if prev is not None and prev.kind == "table":
                    _tiny_paragraph(doc)
                elif page_break:
                    tp = _tiny_paragraph(doc)
                    tp.paragraph_format.page_break_before = True
                if prev is not None and prev_para is not None and prev.kind != "table":
                    prev_para.paragraph_format.space_after = Pt(b.gap_before)
                write_table(doc, td, ml, b.x0, text_w, max_h)
                prev_para = None
            ncount += 1
        except Exception as e:  # noqa
            log.warning("page %d: failed to write %s block: %s", b.page + 1, b.kind, e)
        prev = b

    if ncount == 0:
        doc.add_paragraph("")
    log.info("写入 docx 完成: %d 块, 耗时 %.1fs", ncount, time.time() - t_write)
    # --remove-page-head-foot 同时要两份时，先落不带页眉页脚的那份，
    # 再把页眉页脚填回同一个 Document 后保存完整版，避免把整份文档重写一遍。
    if also_save_without_hf:
        t_save0 = time.time()
        try:
            doc.save(also_save_without_hf)
            log.info("已保存去除页眉页脚的版本: %s (%.1fs)", also_save_without_hf, time.time() - t_save0)
        except Exception as e:  # noqa
            log.error("去除页眉页脚的版本保存失败: %s", e)
    if cfg.strip_header_footer:
        for sec, key, m in section_fills:
            for zone, container in (("head", sec.header), ("foot", sec.footer)):
                if hf.get(zone):
                    try:
                        # add_section 会把上一节的 sectPr（含 headerReference）克隆过来，
                        # 导致后续节被当成“链接到前一节”。先断开再新建，保证每节独立页眉页脚。
                        if container.is_linked_to_previous is False:
                            container.is_linked_to_previous = True
                        container.is_linked_to_previous = False
                        fill_header_footer(container, hf[zone], key[0], m[0], m[1], normal)
                    except Exception as e:  # noqa
                        log.warning("header/footer failed: %s", e)
    log.info("正在序列化保存 ...")
    t_save = time.time()
    doc.save(out_path)
    log.info("docx 已落盘: %.1fs", time.time() - t_save)


# ============================================================================
# 主流程
# ============================================================================


def parse_page_range(spec: str, total: int) -> Set[int]:
    res: Set[int] = set()
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            s = int(a) if a else 1
            e = int(b) if b else total
        else:
            s = e = int(part)
        res.update(i - 1 for i in range(max(s, 1), min(e, total) + 1))
    return res


def _open_pdf(path: str, password: Optional[str]):
    try:
        return pdfplumber.open(path, password=password)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            raise RuntimeError("PDF 已加密或密码错误，请使用 --password 提供密码") from e
        log.warning("直接打开失败 (%s)，尝试用 pypdf 修复后重试 ...", e)
        try:
            from pypdf import PdfReader, PdfWriter

            reader = PdfReader(path, strict=False, password=password)
            w = PdfWriter()
            for pg in reader.pages:
                w.add_page(pg)
            buf = io.BytesIO()
            w.write(buf)
            buf.seek(0)
            return pdfplumber.open(buf)
        except Exception as e2:
            raise RuntimeError(f"无法打开 PDF (文件可能已损坏): {e2}") from e2


def hf_stripped_path(docx_path: str) -> str:
    """--remove-page-head-foot 产出的「无页眉页脚」docx 路径。"""
    root, ext = os.path.splitext(docx_path)
    return root + "-no-head-foot" + (ext or ".docx")


def hf_report_path(docx_path: str) -> str:
    """逐页页眉/页脚/页码检测报告的 JSON 路径。"""
    root, _ext = os.path.splitext(docx_path)
    return root + "-page-head-foot.json"


def write_hf_report(report: dict, path: str, source: str, output: str) -> Optional[str]:
    data = {
        "source": os.path.abspath(source),
        "output": os.path.abspath(output),
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        **report,
    }
    try:
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return path
    except Exception as e:  # noqa
        log.warning("页眉页脚报告写入失败: %s", e)
        return None


def log_hf_report(report: dict) -> None:
    """把逐页检测结果打到日志，同时拼一份可读摘要。"""
    if report.get("detected_pages", 0) == 0:
        log.info("页眉页脚检测: 未发现可去除的页眉/页脚/页码")
        return
    n_items = sum(len(e["items"]) for e in report["pages"])
    log.info("页眉页脚检测: %d 页共 %d 处 (阈值 %s 页)",
             report["detected_pages"], n_items, report.get("threshold"))
    for entry in report["pages"]:
        if not entry["items"]:
            continue
        desc = " | ".join("%s: %s" % (_KIND_CN.get(it["type"], it["type"]), it["text"])
                          for it in entry["items"])
        log.info("  第 %d 页  %s", entry["page"], desc)


def _pdf_info(pdf_path: str, cfg: Config) -> Tuple[int, dict]:
    """取页数与元数据（先用主进程打开一次）。"""
    pdf = _open_pdf(pdf_path, cfg.password)
    try:
        return len(pdf.pages), dict(pdf.metadata or {})
    finally:
        pdf.close()


def _parse_chunk(task):
    """子进程入口：打开同一份 PDF，只解析分配到的页。

    每页的解析互不依赖，适合多进程并行；返回 (PageRaw 列表, 页缓存 manifest 条目)。
    """
    pdf_path, password, idxs, cfg, cache_dir = task
    pdf = _open_pdf(pdf_path, password)
    out: List[PageRaw] = []
    entries: List[dict] = []
    local_manifest: Dict[str, Any] = {"pages": []}
    try:
        for i in idxs:
            page = pdf.pages[i]
            t0 = time.time()
            try:
                raw = parse_page(page, i, cfg)
            except Exception as e:  # noqa
                log.warning("第 %d 页解析失败，整页转为图片兜底: %s", i + 1, e)
                raw = PageRaw(i, float(page.width), float(page.height))
                try:
                    raw.fallback = ImageData(render_region(page, page.bbox, cfg.image_dpi), raw.W, raw.H)
                except Exception as e2:  # noqa
                    log.error("第 %d 页兜底渲染也失败，已跳过: %s", i + 1, e2)
            out.append(raw)
            if cache_dir:
                entry = save_page_cache(raw, cache_dir, local_manifest)
                if entry is not None:
                    entry["seconds"] = round(time.time() - t0, 3)
            try:
                page.flush_cache()
            except Exception:
                pass
    finally:
        pdf.close()
    entries = list(local_manifest["pages"])
    return out, entries


def _parse_pages_parallel(pdf_path: str, cfg: Config, wanted: List[int],
                          cache_dir: Optional[str]) -> List[PageRaw]:
    """多进程解析：按页交错分片，保持结果顺序与单进程一致。"""
    from concurrent.futures import ProcessPoolExecutor, as_completed

    jobs = max(1, min(cfg.jobs, len(wanted)))
    chunks: List[List[int]] = [wanted[k::jobs] for k in range(jobs)]
    tasks = [(pdf_path, cfg.password, ch, cfg, cache_dir) for ch in chunks if ch]
    results: Dict[int, PageRaw] = {}
    entries: List[dict] = []
    done = 0
    log.info("多进程解析: %d 进程, 共 %d 页", len(tasks), len(wanted))
    with ProcessPoolExecutor(max_workers=len(tasks)) as ex:
        futs = {ex.submit(_parse_chunk, t): t for t in tasks}
        for fut in as_completed(futs):
            raws, ents = fut.result()
            for raw in raws:
                results[raw.index] = raw
            entries.extend(ents)
            done += len(raws)
            log.info("解析进度: %d/%d 页 (%.0f%%)", done, len(wanted), 100.0 * done / max(len(wanted), 1))
    if cache_dir:
        manifest = {
            "source": os.path.abspath(pdf_path),
            "output": "",
            "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "pages": sorted(entries, key=lambda e: e.get("page", 0)),
        }
        mp = write_cache_manifest(cache_dir, manifest)
        if mp:
            log.info("逐页缓存已保存 %d 页: %s", len(entries), mp)
    return [results[i] for i in sorted(results)]


def pdf_to_docx(pdf_path: str, docx_path: Optional[str] = None, cfg: Optional[Config] = None) -> str:
    cfg = cfg or Config()
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(pdf_path)
    if docx_path is None:
        docx_path = os.path.splitext(pdf_path)[0] + ".docx"
    out_dir = os.path.dirname(os.path.abspath(docx_path))
    os.makedirs(out_dir, exist_ok=True)

    t_all = time.time()
    total, meta = _pdf_info(pdf_path, cfg)
    if total == 0:
        raise RuntimeError("PDF 没有页面")
    wanted = sorted(cfg.pages) if cfg.pages is not None else list(range(total))
    wanted = [i for i in wanted if 0 <= i < total]
    cache_dir = cache_dir_of(cfg, docx_path) if cfg.cache_pages else None
    manifest: Dict[str, Any] = {
        "source": os.path.abspath(pdf_path),
        "output": os.path.abspath(docx_path),
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "pages": [],
    }
    if cache_dir:
        os.makedirs(cache_dir, exist_ok=True)
        log.info("逐页缓存已启用: %s", cache_dir)
    if cfg.jobs > 1 and len(wanted) > 8:
        raws = _parse_pages_parallel(pdf_path, cfg, wanted, cache_dir)
    else:
        raws = []
        pdf = _open_pdf(pdf_path, cfg.password)
        try:
            for n, i in enumerate(wanted, 1):
                page = pdf.pages[i]
                log.info("解析第 %d/%d 页 (PDF 第 %d 页)", n, len(wanted), i + 1)
                t0 = time.time()
                try:
                    raw = parse_page(page, i, cfg)
                except Exception as e:  # noqa
                    log.warning("第 %d 页解析失败，整页转为图片兜底: %s", i + 1, e)
                    raw = PageRaw(i, float(page.width), float(page.height))
                    try:
                        raw.fallback = ImageData(render_region(page, page.bbox, cfg.image_dpi), raw.W, raw.H)
                    except Exception as e2:  # noqa
                        log.error("第 %d 页兜底渲染也失败，已跳过: %s", i + 1, e2)
                raws.append(raw)
                if cache_dir:
                    entry = save_page_cache(raw, cache_dir, manifest)
                    if entry is not None:
                        entry["seconds"] = round(time.time() - t0, 3)
                try:
                    page.flush_cache()
                except Exception:
                    pass
            if cache_dir:
                mp = write_cache_manifest(cache_dir, manifest)
                if mp:
                    log.info("逐页缓存已保存 %d 页: %s", len(manifest["pages"]), mp)
        finally:
            pdf.close()
    log.info("解析完成: %d 页, 耗时 %.1fs", len(raws), time.time() - t_all)

    t_hf = time.time()
    hf: Dict[str, List[List[HFPart]]] = {"head": [], "foot": []}
    hf_report: Optional[dict] = None
    if cfg.strip_header_footer or cfg.remove_page_head_foot:
        hf, hf_report = strip_headers_footers(raws, cfg)
        if not cfg.strip_header_footer:      # 只要删掉，不放进 Word 页眉页脚
            hf = {"head": [], "foot": []}
        log.info("页眉页脚分析完成: %.1fs", time.time() - t_hf)

    # 正文字号 -> 标题层级
    sz_counter: Counter = Counter()
    for r in raws:
        for l in r.lines:
            sz_counter[l.size] += sum(len(x.text) for x in l.runs)
    body = sz_counter.most_common(1)[0][0] if sz_counter else 10.0
    big = sorted({s for s in sz_counter if s >= body * 1.15}, reverse=True)
    levels = {s: min(i + 1, 3) for i, s in enumerate(big)}

    page_map = {r.index: r for r in raws}
    all_blocks: List[Block] = []
    t_layout = time.time()
    for r in raws:
        try:
            all_blocks.extend(build_page_blocks(r, cfg, body, levels))
        except Exception as e:  # noqa
            log.warning("第 %d 页版面分析失败: %s", r.index + 1, e)
    log.info("版面分析完成: %d 块, 耗时 %.1fs", len(all_blocks), time.time() - t_layout)
    extents: List[Tuple[int, float, float, float, float]] = []
    if not all_blocks:
        log.warning("未提取到任何内容 (可能是扫描件，需要先 OCR)")
    else:
        top_ref = min(b.top for b in all_blocks)
        bot_ref = max(b.bottom for b in all_blocks)
        extents = [(b.page, b.x0, b.top, b.x1, b.bottom) for b in all_blocks]
        all_blocks = merge_across(all_blocks, page_map, (top_ref, bot_ref))
        log.info("跨页合并完成: %d 块", len(all_blocks))

    if _STATS["cid"] > 20:
        log.warning("检测到 %d 个无法映射到 Unicode 的字形 (字体缺少 ToUnicode)，这些字符已被丢弃/替换；"
                    "如文字大量缺失，请先对 PDF 做 OCR 再转换", _STATS["cid"])
    t_write = time.time()
    stripped = hf_stripped_path(docx_path) if cfg.remove_page_head_foot else None
    write_docx(all_blocks, page_map, hf, cfg, docx_path, meta, extents, also_save_without_hf=stripped)
    log.info("写出 docx 完成: %.1fs (总耗时 %.1fs)", time.time() - t_write, time.time() - t_all)
    log.info("完成: %s", docx_path)

    # --remove-page-head-foot：先输出逐页检测信息；去页眉页脚的 docx 已在 write_docx 里一并落盘
    if cfg.remove_page_head_foot:
        if hf_report is not None:
            log_hf_report(hf_report)
            rp = write_hf_report(hf_report, hf_report_path(docx_path), pdf_path, docx_path)
            if rp:
                log.info("页眉页脚检测报告: %s", rp)
        if stripped:
            log.info("已额外保存去除页眉页脚的版本: %s", stripped)
    return docx_path


def main(argv=None):
    ap = argparse.ArgumentParser(description="PDF -> Word，尽量保留段落/表格/图片的格式与位置")
    ap.add_argument("pdf")
    ap.add_argument("-o", "--output")
    ap.add_argument("--pages", help="页码范围，如 1-3,5,8-")
    ap.add_argument("--password")
    ap.add_argument("--keep-page-breaks", action="store_true", help="在原 PDF 分页处强制分页")
    ap.add_argument("--no-columns", action="store_true", help="关闭双栏识别")
    ap.add_argument("--no-figures", action="store_true", help="不把矢量图整体渲染为图片")
    ap.add_argument("--no-headings", action="store_true", help="不套用 Word 标题样式")
    ap.add_argument("--no-header-footer", action="store_true", help="不单独识别页眉页脚")
    ap.add_argument("--remove-page-head-foot", action="store_true",
                    help="额外保存一份去除页眉/页脚/页码的 docx（完整版仍照常输出），"
                         "并输出每一页检测到的页眉/页脚/页码信息")
    ap.add_argument("--no-fonts", action="store_true", help="不保留字体名")
    ap.add_argument("--dpi", type=int, default=200)
    ap.add_argument("--cache-page", action="store_true",
                    help="保存每一页的解析结果 (json + txt + 图片) 到缓存目录；默认不缓存")
    ap.add_argument("--cache-dir",
                    help="逐页解析中间结果的保存目录，不存在会自动创建"
                         " (默认: 输出 docx 同目录下的 _pdf2word_pages)")
    ap.add_argument("--jobs", type=int, default=1,
                    help="解析进程数；>1 时多进程按页并行（页面解析互不依赖，长文档能快好几倍）")
    ap.add_argument("-v", "--verbose", action="store_true")
    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.DEBUG if a.verbose else logging.INFO, format="%(levelname)s %(message)s")
    cfg = Config(
        password=a.password, keep_page_breaks=a.keep_page_breaks, detect_columns=not a.no_columns,
        render_figures=not a.no_figures, detect_headings=not a.no_headings,
        strip_header_footer=not a.no_header_footer, keep_fonts=not a.no_fonts, image_dpi=a.dpi,
        cache_pages=a.cache_page, cache_dir=a.cache_dir,
        remove_page_head_foot=a.remove_page_head_foot, jobs=max(1, a.jobs),
    )
    try:
        if a.pages:
            with _open_pdf(a.pdf, a.password) as p:
                cfg.pages = parse_page_range(a.pages, len(p.pages))
        out = pdf_to_docx(a.pdf, a.output, cfg)
        if cfg.cache_pages:
            print("cache_dir: %s" % cache_dir_of(cfg, out))
    except Exception as e:  # noqa
        log.error("转换失败: %s", e)
        if a.verbose:
            raise
        return 1
    print(out)
    if cfg.remove_page_head_foot:
        print(hf_stripped_path(out))
        print(hf_report_path(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())