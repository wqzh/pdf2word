"""命令行入口：pdf2word-pro <pdf> [-o out.docx] [选项]"""

import sys

from .core import main as _convert_main


def main(argv=None) -> int:
    """console_scripts 入口，返回进程退出码。"""
    return _convert_main(argv)


if __name__ == "__main__":
    sys.exit(main())