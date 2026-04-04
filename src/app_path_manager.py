"""
统一的路径和配置管理模块

集中管理项目中的所有路径常量，便于维护和配置修改。
"""

import os
from pathlib import Path


class PathConfig:
    """全局路径和文件配置"""

    # ==================== 应用根目录 ====================
    # 从 src/ 上升到项目根目录
    APP_ROOT = Path(__file__).parent.parent

    # ==================== 数据目录 ====================
    DATA_DIR = os.path.join(APP_ROOT, "data")

    # ==================== 数据文件（JSON清单） ====================
    # 应用迁移清单
    MANIFEST_FILE = os.path.join(DATA_DIR, "app_migration_manifest.json")
    # 待清理项目文件
    PENDING_CLEAN_FILE = os.path.join(DATA_DIR, "app_migration_pending_cleanup.json")
    # 盘符修复清单
    DRIVE_FIX_MANIFEST_FILE = os.path.join(DATA_DIR, "drive_fix_manifest.json")
    # UI 状态文件
    UI_STATE_FILE = os.path.join(DATA_DIR, "app_gui_state.json")

    # ==================== 备份根目录 ====================
    BACKUPS_DIR = os.path.join(APP_ROOT, "app_backups")
    # 应用迁移备份
    BACKUP_ROOT = os.path.join(BACKUPS_DIR, "migrations")
    # 盘符修复备份
    DRIVE_FIX_BACKUP_ROOT = os.path.join(BACKUPS_DIR, "drive_fixes")

    # ==================== 日志配置 ====================
    LOG_DIR = os.path.join(APP_ROOT, "app_logs")

    # ==================== 临时目录 ====================
    TEMP_DIR = os.path.join(APP_ROOT, "app_temp")
    PENDING_CLEANUP_DIR = os.path.join(TEMP_DIR, "cleanup_jobs")

    # ==================== 默认迁移路径 ====================
    DEFAULT_TARGET_ROOT = r"D:\Program Files"

    @classmethod
    def ensure_directories(cls):
        """确保所有必要的目录存在"""
        directories = [
            cls.DATA_DIR,
            cls.BACKUP_ROOT,
            cls.DRIVE_FIX_BACKUP_ROOT,
            cls.LOG_DIR,
            cls.PENDING_CLEANUP_DIR,
        ]
        for path in directories:
            os.makedirs(path, exist_ok=True)
