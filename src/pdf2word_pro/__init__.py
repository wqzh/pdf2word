"""pdf2word-pro —— 把 PDF 转成 Word(.docx)，尽量保留原版面。

对外只暴露三样东西：

    from pdf2word_pro import Config, pdf_to_docx, __version__

底层实现位于 :mod:`pdf2word_pro.core`（由原单文件 `pdf2docx.py` 迁移而来）。
"""

from ._version import __version__
from .core import Config, pdf_to_docx

__all__ = ["Config", "pdf_to_docx", "__version__"]