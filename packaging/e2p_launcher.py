"""PyInstaller 入口：打包后双击或命令行运行。"""

from __future__ import annotations

import sys

from excel2ppt_studio.cli import main

if __name__ == "__main__":
    sys.exit(main())
