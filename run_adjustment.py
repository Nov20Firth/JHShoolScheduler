# -*- coding: utf-8 -*-
"""以 ASCII 檔名啟動中文檔名的國中微調程式。"""

import runpy
from pathlib import Path


runpy.run_path(str(Path(__file__).with_name("國中微調.py")), run_name="__main__")
