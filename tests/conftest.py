# -*- coding: utf-8 -*-
"""pump_calc_app.py를 tests/ 밖(저장소 루트)에서 import할 수 있게 경로를 잡아준다."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
