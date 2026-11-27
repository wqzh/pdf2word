[中文](https://github.com/wqzh/pdf2word/blob/main/README.md) | [English](https://github.com/wqzh/pdf2word/blob/main/README_EN.md)

# pdf2word-pro

把 PDF 转成 Word（`.docx`）的命令行工具，目标是**尽量保留原 PDF 的版面**：段落、字体字号、表格（含合并单元格与底色）、图片（含并排插图与图注）、页眉页脚、分页与页面尺寸。

纯 Python 自研实现，只依赖 `pdfplumber` 做 PDF 底层解析、`python-docx` 做 Word 组装。安装后命令行入口为 `pdf2word-pro`，同时支持 `python -m pdf2word_pro`；仓库内仍保留 `python pdf2docx.py` 旧入口。

---

## 1. 特性概览

| 类别 | 说明 |
| --- | --- |
| 文字 | 按「字符 → 行 → 段落」重建，保留字体、字号、粗体/斜体、颜色；中文不加空格、英文去连字符断词 |
| 段落 | 用「上一行右侧剩余空间能否放下下一行首个词」判断软换行还是真换行；保留对齐方式、首行缩进、行距/段间距（按原 PDF 的纵向间距换算成 Word 的段前/段后） |
| 标题 | 字号明显大于正文字号的短段落自动套用 Word 的「标题 1/2/3」样式（可在转换后统一改样式） |
| 双栏 | 自动检测双栏排版的中缝（gutter），按栏输出 |
| 表格 | 由 PDF 的框线还原网格、合并单元格、底色；单元格内文字按原字号缩放；跨页续表自动拼接并去掉重复表头 |
| 图片 | 优先取原始图像流（JPEG 原样、无损 RGB/灰度），取不到再按区域渲染；保留原始大小、左右缩进/居中对齐 |
| 并排插图 | 同一水平带的若干张图（如 a) b) c) d) 四宫格、两图并排）按原列位还原成无边框表格，不拆散、不重排 |
| 嵌入图 | 压在大图上的图例/局部放大图并入外层大图一起渲染，避免被当成「第二栏」错位 |
| 矢量图 | 曲线/色块聚类识别出的图表整体渲染成图片，并剔除其内部文字，避免文字重复 |
| 页眉页脚 | 跨页重复的行识别为页眉/页脚，页码写成 `PAGE`/`NUMPAGES` 域；可选择额外导出一份「无页眉页脚」版本 |
| 页面 | 按原页面尺寸和内容边界计算页边距，混合横竖版面时自动分节 |
| 鲁棒性 | 加密/损坏 PDF 用 `pypdf` 重写后再试；单页解析失败整页转图片兜底；过滤非法 XML 字符 |

---

## 2. 环境与依赖

- Python **3.9+**（开发验证环境：Python 3.11.9 / Windows 10）
- 依赖（见 `requirements.txt`）：

| 包 | 用途 | 验证版本 |
| --- | --- | --- |
| `pdfplumber` | 读取 PDF 页面、字符、表格、线条、图片 | 0.11.4 |
| `python-docx` | 生成 `.docx` | 1.2.0 |
| `pillow` | 图片解码/编码（PNG 输出） | 13.0.0 |
| `pypdf` | 修复损坏/加密 PDF 的兜底重写 | 7.1.0 |

`concurrent.futures`、`argparse`、`json`、`logging` 等均为标准库，无需额外安装。

### 安装

```shell
pip install pdf2word-pro
```

从源码安装（开发用）：

```shell
pip install -r requirements.txt
pip install -e .
```

---

## 3. 快速开始

```shell
pdf2word-pro input.pdf
```

不指定 `-o` 时，输出文件为「与 PDF 同名、同目录的 `.docx`」，并把最终路径打印到标准输出。

```shell
pdf2word-pro input.pdf -o output.docx
```

---

## 4. 命令行参数

```text
usage: pdf2word-pro [-h] [-o OUTPUT] [--pages PAGES] [--password PASSWORD]
                   [--keep-page-breaks] [--no-columns] [--no-figures]
                   [--no-headings] [--no-header-footer]
                   [--remove-page-head-foot] [--no-fonts] [--dpi DPI]
                   [--cache-page] [--cache-dir CACHE_DIR] [--jobs JOBS] [-v]
                   pdf
```

| 参数 | 默认值 | 说明 |
| --- | --- | --- |
| `pdf` | 必填 | 输入 PDF 路径（位置参数） |
| `-o`, `--output` | 与输入同名 `.docx` | 输出 `.docx` 路径；目录不存在会自动创建 |
| `--pages` | 全部页 | 只转换指定页，1 开始计数。支持 `1-3,5,8-` 形式（`8-` 表示第 8 页到末页），自动裁剪越界值 |
| `--password` | 无 | 加密 PDF 的打开密码 |
| `--keep-page-breaks` | 关 | 在原 PDF 的分页处强制插入分页符（默认按内容自然流动重排） |
| `--no-columns` | 关 | 关闭双栏识别，按单栏整行读取 |
| `--no-figures` | 关 | 不把矢量图整体渲染成图片（只输出文字和位图） |
| `--no-headings` | 关 | 不套用 Word 标题样式，全部按正文段落输出 |
| `--no-header-footer` | 关 | 不单独识别页眉页脚（页眉页脚文字留在正文里） |
| `--remove-page-head-foot` | 关 | 除完整版外，**额外**保存一份去掉页眉/页脚/页码的 `<名>-no-head-foot.docx`，并把逐页检测结果写入控制台和 `<名>-page-head-foot.json` |
| `--no-fonts` | 关 | 不保留原字体名（统一用文档默认字体） |
| `--dpi` | `200` | 区域渲染图片（含兜底整页渲染）时使用的 DPI，越大越清晰、文件越大 |
| `--cache-page` | 关 | 把每一页的解析结果（JSON + 纯文本 + 图片）落盘，便于排查版面问题 |
| `--cache-dir` | 输出目录下 `_pdf2word_pages` | 逐页缓存目录，不存在会自动创建 |
| `--jobs` | `1` | 解析进程数；`>1` 且页数大于 8 时按页多进程并行（页面解析互不依赖，长文档可快数倍） |
| `-v`, `--verbose` | 关 | 输出 DEBUG 级日志，并在转换失败时抛出完整异常栈 |

退出码：`0` 成功；`1` 转换失败（日志里会打印原因，`-v` 可看完整栈）。

---

## 5. 输出文件

以 `-o output.docx` 为例：

| 文件 | 何时产生 | 说明 |
| --- | --- | --- |
| `output.docx` | 总是 | 完整版，含页眉页脚 |
| `output-no-head-foot.docx` | 加 `--remove-page-head-foot` 时 | 去掉页眉/页脚/页码的版本 |
| `output-page-head-foot.json` | 加 `--remove-page-head-foot` 时 | 逐页检测到的页眉/页脚/页码内容 |
| `_pdf2word_pages\page_0001.json` 等 | 加 `--cache-page` 时 | 该页结构化解析结果（行/字符/表格/图片框，UTF-8 可读） |
| `_pdf2word_pages\page_0001.txt` | 加 `--cache-page` 时 | 该页纯文本视图（逐行文本 + 表格单元格文本） |
| `_pdf2word_pages\page_0001_*.jpg/png` | 加 `--cache-page` 时 | 该页提取出的图片/矢量图渲染结果 |
| `_pdf2word_pages\manifest.json` | 加 `--cache-page` 时 | 全部页的摘要索引（页号、块数、文件、耗时） |

---

## 6. 运行示例

### 6.1 整篇转换

```shell
pdf2word-pro input.pdf -o output.docx
```

### 6.2 只检查某几页（排版回归时最常用）

```shell
pdf2word-pro input.pdf -o sample.docx --pages 12-14,20
```

### 6.3 长文档多进程加速

```shell
pdf2word-pro input.pdf -o output-job8.docx --jobs 8
```

### 6.4 去掉页眉页脚（同时保留完整版）

```shell
pdf2word-pro input.pdf -o output.docx --remove-page-head-foot
# 产出：output.docx、output-no-head-foot.docx、output-page-head-foot.json
```

### 6.5 加密 PDF

```shell
pdf2word-pro encrypted.pdf -o decrypted.docx --password "你的密码"
```

### 6.6 保留原始分页 + 关闭双栏识别（用于核对版面）

```shell
pdf2word-pro input.pdf -o output-layout.docx `
  --keep-page-breaks --no-columns --cache-page --cache-dir _dbg -v
```

### 6.7 作为库调用

```python
from pdf2word_pro import Config, pdf_to_docx

cfg = Config(pages={45, 46, 47}, image_dpi=300, jobs=4)   # pages 为 0 基页码
out = pdf_to_docx("input.pdf", "output.docx", cfg)
print(out)
```

`Config` 的字段与命令行参数一一对应（见 `Config` 定义处的注释），例如 `keep_page_breaks`、`detect_columns`、`render_figures`、`image_dpi`、`jobs`、`cache_pages`、`cache_dir`、`remove_page_head_foot`。

---

## 7. 实现流程

1. **逐页解析**（`parse_page`，pdfplumber）
   - 表格：`find_tables()` → 网格 + 合并单元格 + 底色 + 单元格内文字/图片
   - 图片：优先原始图像流，失败则按区域渲染；区域内嵌图（图例/局部图）并入外层图
   - 矢量图：曲线/色块聚类后整体渲染，并剔除其内部文字
   - 文字：字符 → 行 → 段落（字体/字号/粗斜体/颜色）；自动识别双栏
2. **全局处理**
   - 页眉页脚：跨页重复行识别，写入 Word 页眉/页脚，页码用域代码
   - 段落折行合并：英文去连字符、中文不加空格
   - 跨页合并：表格（列边界对齐 + 去重复表头）与段落（末行写满且下页无缩进）
3. **写 docx**（`write_docx`）
   - 页面尺寸/页边距、对齐、缩进、行距/段间距、表格列宽行高、图片大小/对齐
   - 并排图片写成无边框表格，一行图片不跨页、末行与图注同页
4. **鲁棒性**：加密/损坏 PDF 兜底、单页失败整页转图、非法 XML 字符过滤

---

## 8. 并排图片的排版策略

这是本工具最容易出问题的部分，实现要点：

- **成行判定**：同一水平带内的图片才归为一行；列数不一致的相邻行不合并，避免张冠李戴。
- **一图一列**：网格里每张图各占一列，相邻图以两张图边界的**中线**为界；不做单元格横向合并，因此即使原 PDF 里两张图轻微重叠，也不会出现「两张图挤进同一个合并单元格」的糊图。
- **行高与分页**：表格行设为 `cantSplit`（行内不跨页拆开），末行段落设 `keep_with_next`，使图片行与紧随其后的图注（`a) … b) …` 或 `图 x.x-x`）尽量同页。
- **内嵌图**：用包含关系（面积比 < 90% 且被完全包住）识别「图例/局部放大图」，把它并进外层大图的区域渲染结果里，而不是当成并排的第二张图。
- **对齐**：整行居中时按原表格左边界设置表格缩进，缩放比例按可用宽度等比计算，不拉伸变形。

---

## 9. 已知限制与排错

- **扫描版 PDF**：没有文字层，只能整页转图片（日志提示「未提取到任何内容 (可能是扫描件，需要先 OCR)」）。请先 OCR 再转换。
- **字体缺少 ToUnicode**：会出现「检测到 N 个无法映射到 Unicode 的字形」告警，这些字符只能丢弃/替换。
- **表格单元格内多张图并排**：目前按上下顺序输出，尚未做单元格内的并排还原。
- **矢量图形与位图叠加的图**：若底层是位图、彩色区域是矢量绘制，取原始图像流时可能只剩底图；此时可试 `--no-figures` 对比，或用 `--cache-page` 查看该页解析结果。
- **转换很慢**：加 `--jobs 8`；若怀疑某页卡住，用 `--pages` 单页定位并加 `-v` 看阶段耗时。
- **图片过大/过小**：调 `--dpi`（默认 200）。DPI 越高越清晰，`.docx` 也越大。

---

## 10. 目录结构

```text
pyproject.toml            打包配置（PEP 621 / setuptools）
LICENSE                   MIT 许可证
requirements.txt          依赖清单
README.md                 本文件
README_EN.md              英文说明
pdf2docx.py               兼容入口（旧命令）
src/pdf2word_pro/         包源码
  core.py                 转换实现（原 pdf2docx.py）
  cli.py                  命令行入口
  __main__.py             python -m 入口
  __init__.py             对外接口
```

发布到 PyPI / TestPyPI 的完整步骤见 [docs/PUBLISHING.md](https://github.com/wqzh/pdf2word/blob/main/docs/PUBLISHING.md)。
