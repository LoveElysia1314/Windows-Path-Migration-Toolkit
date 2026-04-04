#!/usr/bin/env python3
"""
应用启动脚本

将 src 目录加入 Python 路径，然后运行 GUI 应用
"""

import sys
from pathlib import Path

# 将 src 目录添加到 Python 路径
project_root = Path(__file__).parent
src_path = project_root / "src"
sys.path.insert(0, str(src_path))

# 导入并运行主应用
from app_path_migration_gui import main

if __name__ == "__main__":
    main()
