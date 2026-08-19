# -*- coding: utf-8 -*-
"""以 ASCII 檔名啟動中文檔名的國中排課程式。"""

import runpy
from pathlib import Path


runpy.run_path(
    str(Path(__file__).with_name("\u570b\u4e2d\u6392\u8ab2.py")),
    run_name="__main__",
)
