"""支持 python -m pdf2word_pro。"""

import sys

from .cli import main

if __name__ == "__main__":
    sys.exit(main())