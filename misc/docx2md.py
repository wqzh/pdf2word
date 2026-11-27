# -*- coding: utf-8 -*-
"""docx → Markdown 转换核心（保留标题层级）。

设计要点
--------
- 直接解析 OOXML（lxml），不走 python-docx 的对象模型：一次读取、一次遍历，速度快、内存可控。
- 标题层级来源（优先级从高到低）：
  1. 段落自身的 ``w:pPr/w:outlineLvl``；
  2. 段落样式的 ``w:pPr/w:outlineLvl``（沿 ``w:basedOn`` 链继承）；
  3. 样式名 ``Heading N`` / ``标题 N``。
  ``outlineLvl`` 取值 0~8 对应 Markdown 的 ``#``~``#########``（超过 6 级夹到 6），9 表示正文。
- 段落文本按文档顺序递归拼接 ``w:t`` / ``w:tab`` / ``w:br``，跳过域代码指令、修订删除、图形等子树。
- 表格转标准 GFM（首行为表头），合并单元格按 ``gridSpan`` 补空列、``vMerge`` 续行沿用起始格文本。
- 块间以空行分隔，段落与表格交错保序。
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from lxml import etree

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W_NS = "{" + W + "}"

MAX_HEADING_LEVEL = 6

_P_TAG = W_NS + "p"
_TBL_TAG = W_NS + "tbl"
_TR_TAG = W_NS + "tr"
_TC_TAG = W_NS + "tc"
_SDT_TAG = W_NS + "sdt"
_SDT_CONTENT_TAG = W_NS + "sdtContent"
_STYLE_TAG = W_NS + "style"
_OUTLINE_LVL_TAG = W_NS + "outlineLvl"
_STYLE_ID_ATTR = W_NS + "styleId"
_VAL_ATTR = W_NS + "val"
_V_MERGE_TAG = W_NS + "vMerge"

# 不参与文本提取的子树：段落/字符属性（含 w:tabs/w:tab，混进来会伪造制表符）、
# 修订删除、域代码指令、图形对象（文本框内的段落由 drawing 自带，重复提取会串行）
_SKIP_SUBTREE = {
    W_NS + tag
    for tag in (
        "pPr",
        "rPr",
        "del",
        "moveFrom",
        "instrText",
        "delInstrText",
        "delText",
        "drawing",
        "pict",
        "object",
    )
}

_HEADING_NAME_RE = re.compile(r"^(?:heading|标题)\s*(\d+)$", re.IGNORECASE)

# 段落级的文本片段标签
_TEXT_TAG = W_NS + "t"
_TAB_TAG = W_NS + "tab"
_BR_TAGS = {W_NS + "br", W_NS + "cr"}
_NO_BREAK_HYPHEN_TAG = W_NS + "noBreakHyphen"


class DocxConversionError(ValueError):
    """docx 结构不可用（非 zip / 缺少主文档部件等）。"""


def _paragraph_text(p_el) -> str:
    """按文档顺序取段落全文。

    ``Paragraph.text`` 只拼接 ``w:p`` 的直接 ``w:r`` 子节点，会漏掉
    ``w:fldSimple``（如目录里的 PAGEREF 页码）、``w:hyperlink``、``w:smartTag``、
    ``w:sdt``、``w:ins`` 等容器内的 run。这里递归遍历所有后代，
    保留 ``w:t`` 文本与 ``w:tab`` / ``w:br`` 等排版字符；
    域代码 ``w:instrText`` 与修订删除内容跳过，域结果（separate 之后的 run）正常保留。
    """
    parts: list[str] = []

    def walk(el) -> None:
        for child in el:
            tag = child.tag
            if tag in _SKIP_SUBTREE:
                continue
            if tag == _TEXT_TAG:
                if child.text:
                    parts.append(child.text)
            elif tag == _TAB_TAG:
                parts.append("\t")
            elif tag in _BR_TAGS:
                parts.append("\n")
            elif tag == _NO_BREAK_HYPHEN_TAG:
                parts.append("-")
            else:
                walk(child)

    walk(p_el)
    return "".join(parts)


def _cell_text(tc) -> str:
    """单元格全文：按文档顺序取单元格内段落，段间以换行分隔；嵌套表格不展开。"""
    parts: list[str] = []

    def walk(el) -> None:
        for child in el:
            tag = child.tag
            if tag == _P_TAG:
                parts.append(_paragraph_text(child))
            elif tag != _TBL_TAG:
                walk(child)  # 内容控件等容器继续向下展开

    walk(tc)
    return "\n".join(parts)


def _int_attr(el) -> int:
    """取可选的整数型 ``w:val``，缺失/非法时返回 0。"""
    if el is None:
        return 0
    value = (el.get(_VAL_ATTR) or "").strip()
    return int(value) if value.lstrip("-").isdigit() else 0


def _normalize_cell(text: str) -> str:
    """单元格文本规范化：去首尾空白、换行折成空格、竖线转义（避免破坏 GFM 列结构）。"""
    return (text or "").strip().replace("\n", " ").replace("|", "\\|")


def _table_to_gfm(tbl) -> str:
    """docx 表格 → GFM（首行为表头，空表返回空串）。

    合并单元格：横向 ``gridSpan`` 补空列；纵向 ``vMerge`` 续行沿用合并起始格文本
    （与 Word 显示一致，每行自成一体，便于下游按行阅读）。
    """
    rows: list[list[str]] = []
    merge_text: dict[int, str] = {}  # 网格列号 → 当前纵向合并区域的文本
    for tr in tbl.iterchildren(_TR_TAG):
        col = 0
        cells: list[str] = []
        for tc in tr.iterchildren(_TC_TAG):
            tc_pr = tc.find(W_NS + "tcPr")
            span = max(1, _int_attr(tc_pr.find(W_NS + "gridSpan") if tc_pr is not None else None))
            merge_el = tc_pr.find(_V_MERGE_TAG) if tc_pr is not None else None
            is_continuation = merge_el is not None and (merge_el.get(_VAL_ATTR) or "continue") == "continue"
            if is_continuation and col in merge_text:
                text = merge_text[col]
            else:
                text = _normalize_cell(_cell_text(tc))
                for offset in range(span):
                    merge_text[col + offset] = text
            cells.append(text)
            if span > 1:
                cells.extend([""] * (span - 1))
            col += span
        rows.append(cells)

    rows = [row for row in rows if any(row)]
    if not rows:
        return ""

    width = max(len(row) for row in rows)
    rows = [row + [""] * (width - len(row)) for row in rows]
    header = [cell or f"列{index + 1}" for index, cell in enumerate(rows[0])]
    lines = [
        "| " + " | ".join(header) + " |",
        "| " + " | ".join(["---"] * width) + " |",
    ]
    lines.extend("| " + " | ".join(row) + " |" for row in rows[1:])
    return "\n".join(lines)


def _style_outline_levels(styles_xml: bytes | None) -> dict[str, int]:
    """styleId → 标题级别（1 起；0 表示不是标题）。

    沿 ``w:basedOn`` 链解析：先看 ``outlineLvl``，再看样式名 ``Heading N`` / ``标题 N``。
    显式的 ``outlineLvl=9`` 表示正文，直接终止继承（不再回溯父样式）。
    """
    if not styles_xml:
        return {}

    raw: dict[str, tuple[int | None, str | None, str | None]] = {}
    try:
        root = etree.fromstring(styles_xml)
    except etree.XMLSyntaxError:
        return {}
    for style in root.iter(_STYLE_TAG):
        style_id = style.get(_STYLE_ID_ATTR)
        if not style_id:
            continue
        name_el = style.find(W_NS + "name")
        based_el = style.find(W_NS + "basedOn")
        outline_el = style.find(f"{W_NS}pPr/{_OUTLINE_LVL_TAG}")
        outline_level = None
        if outline_el is not None:
            value = (outline_el.get(_VAL_ATTR) or "").strip()
            if value.lstrip("-").isdigit():
                outline_level = int(value)
        raw[style_id] = (
            outline_level,
            name_el.get(_VAL_ATTR) if name_el is not None else None,
            based_el.get(_VAL_ATTR) if based_el is not None else None,
        )

    levels: dict[str, int] = {}
    for style_id in raw:
        level = 0
        seen: set[str] = set()
        current: str | None = style_id
        while current and current not in seen:
            seen.add(current)
            outline_level, name, based = raw.get(current, (None, None, None))
            if outline_level is not None:
                # 0~8 → 1~9 级标题；9 及以上是正文
                level = outline_level + 1 if 0 <= outline_level < 9 else 0
                break
            match = _HEADING_NAME_RE.match(name or "")
            if match:
                level = int(match.group(1))
                break
            current = based
        levels[style_id] = level
    return levels


def _paragraph_heading_level(p_el, style_levels: dict[str, int]) -> int:
    """段落标题级别（1 起；0 表示正文）。"""
    ppr = p_el.find(W_NS + "pPr")
    if ppr is None:
        return 0
    outline_el = ppr.find(_OUTLINE_LVL_TAG)
    if outline_el is not None:
        value = (outline_el.get(_VAL_ATTR) or "").strip()
        if value.lstrip("-").isdigit():
            level = int(value)
            return level + 1 if 0 <= level < 9 else 0
    style_el = ppr.find(W_NS + "pStyle")
    if style_el is None:
        return 0
    return style_levels.get(style_el.get(_VAL_ATTR) or "", 0)


def _render_paragraph(p_el, style_levels: dict[str, int]) -> str:
    text = _paragraph_text(p_el).strip()
    if not text:
        return ""
    level = _paragraph_heading_level(p_el, style_levels)
    if level <= 0:
        return text
    return "#" * min(level, MAX_HEADING_LEVEL) + " " + text


def _render_block(el, style_levels: dict[str, int], blocks: list[str]) -> None:
    """把一个 body 级节点渲染成 0..n 个块（段落 / 表格 / 内容控件）。"""
    tag = el.tag
    if tag == _P_TAG:
        block = _render_paragraph(el, style_levels)
        if block:
            blocks.append(block)
    elif tag == _TBL_TAG:
        block = _table_to_gfm(el)
        if block:
            blocks.append(block)
    elif tag in (_SDT_TAG, _SDT_CONTENT_TAG):
        for child in el.iterchildren():
            _render_block(child, style_levels, blocks)


def _read_document_parts(content: bytes) -> tuple[bytes, bytes | None]:
    try:
        with zipfile.ZipFile(io.BytesIO(content)) as package:
            try:
                document_xml = package.read("word/document.xml")
            except KeyError as exc:  # 不是 Word 文档
                raise DocxConversionError("压缩包内缺少 word/document.xml，可能不是 .docx 文件") from exc
            try:
                styles_xml: bytes | None = package.read("word/styles.xml")
            except KeyError:
                styles_xml = None
    except zipfile.BadZipFile as exc:
        raise DocxConversionError("文件不是有效的 .docx（zip）格式") from exc
    return document_xml, styles_xml


def docx_bytes_to_markdown(content: bytes) -> str:
    """把 .docx 二进制内容转成 Markdown 文本（保留标题层级）。"""
    document_xml, styles_xml = _read_document_parts(content)
    style_levels = _style_outline_levels(styles_xml)

    try:
        document = etree.fromstring(document_xml)
    except etree.XMLSyntaxError as exc:
        raise DocxConversionError(f"word/document.xml 解析失败: {exc}") from exc

    body = document.find(W_NS + "body")
    if body is None:
        return ""

    blocks: list[str] = []
    for child in body.iterchildren():
        _render_block(child, style_levels, blocks)
    return "\n\n".join(blocks)


def docx_to_markdown(source: str | Path | bytes) -> str:
    """把 .docx 路径或二进制内容转成 Markdown 文本（保留标题层级）。"""
    if isinstance(source, (bytes, bytearray, memoryview)):
        return docx_bytes_to_markdown(bytes(source))
    return docx_bytes_to_markdown(Path(source).read_bytes())
