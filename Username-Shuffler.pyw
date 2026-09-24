"""
Username Shuffler - Windows Windowed Launcher (No Console)
v3.0.0
"""
from __future__ import annotations

import os
import sys

# Ensure current directory is at the top of sys.path
base_dir = os.path.dirname(os.path.abspath(__file__))
if base_dir not in sys.path:
    sys.path.insert(0, base_dir)

from app import main

if __name__ == "__main__":
    main()
