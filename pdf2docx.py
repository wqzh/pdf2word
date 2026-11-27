"""兼容入口：保留 `python pdf2docx.py ...` 的老用法。

等价于安装后的 `pdf2word-pro ...` 或 `python -m pdf2word_pro ...`。
"""

import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))

from pdf2word_pro.core import main  # noqa: E402

if __name__ == "__main__":
    sys.exit(main())