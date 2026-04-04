"""
应用常量定义模块

集中管理所有应用常量，便于维护和修改。
包括：注册表位置、环境变量路径、文件夹名称、排除关键词等。
"""

import re

try:
    import winreg
except ImportError:
    # 非 Windows 环境下，winreg 不可用，使用 None 占位
    winreg = None

# ==================== 应用版本信息 ====================
VERSION = "0.4.0"

# ==================== 注册表卸载项位置（UNINSTALL） ====================
# 扫描已安装应用的注册表位置（四个位置）
UNINSTALL_REGISTRY_LOCATIONS = [
    # 本机 64 位应用
    (
        winreg.HKEY_LOCAL_MACHINE if winreg else None,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "64bit",
    ),
    # 本机 32 位应用（在 64 位系统中）
    (
        winreg.HKEY_LOCAL_MACHINE if winreg else None,
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "32bit",
    ),
    # 当前用户 64 位应用
    (
        winreg.HKEY_CURRENT_USER if winreg else None,
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        "64bit",
    ),
    # 当前用户 32 位应用（在 64 位系统中）
    (
        winreg.HKEY_CURRENT_USER if winreg else None,
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
        "32bit",
    ),
]

# 环境变量注册表位置
ENVIRONMENT_REGISTRY_LOCATIONS = [
    ("HKCU", r"Environment"),
    ("HKLM", r"SYSTEM\CurrentControlSet\Control\Session Manager\Environment"),
]

# ==================== Windows 系统文件夹（不参与迁移） ====================
# 扁平复制时要排除的顶层文件夹
FLAT_COPY_EXCLUDED_FOLDERS = {
    "windows",
    "users",
    "programdata",
    "program files",
    "program files (x86)",
}

# ==================== 排除的软件发行类型 ====================
# 从控制面板中排除扫描的某些软件分类（补丁/更新等系统组件）
CONTROL_PANEL_EXCLUDED_RELEASE_TYPES = {
    "security update",
    "update rollup",
    "hotfix",
}

# ==================== 排除的厂商关键词 ====================
# 迁移时默认排除的厂商（微软、NVIDIA 驱动等系统级组件）
DEFAULT_EXCLUDED_VENDOR_KEYWORDS = [
    "microsoft",
    "windows",
    "nvidia",
    "drivers",
]

# ==================== 标准安装路径匹配 ====================
# 检测应用是否在标准安装目录中（Program Files / Program Files (x86)）
STANDARD_INSTALL_PATH_PATTERN = re.compile(
    r"^[A-Za-z]:\\Program Files(?: \(x86\))?(?:\\|$)",
    re.IGNORECASE,
)

# ==================== 盘符修复搜索根目录 ====================
# 盘符修复时需要扫描的注册表位置（查找包含路径的注册表项）
DRIVE_FIX_SEARCH_REGISTRY_ROOTS = [
    # 应用卸载项（可能包含安装路径）
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    # 文件关联
    r"SOFTWARE\Classes",
    r"SOFTWARE\WOW6432Node\Classes",
    # 腾讯应用（特殊处理）
    r"SOFTWARE\Tencent",
    r"SOFTWARE\WOW6432Node\Tencent",
    # 资源管理器和浏览器历史
    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer",
    # 计划任务
    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks",
    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree",
    # 系统服务
    r"SYSTEM\CurrentControlSet\Services",
]

# ==================== 快捷方式特殊处理 ====================
# 快捷方式中常见的路径标记
# noqa: F841 - 预留常量，用于将来的快捷方式处理功能
COMMON_SHORTCUT_PATH_MARKERS = [
    "TargetPath",  # LNK 快捷方式的目标路径
    "IconLocation",  # 图标位置
    "WorkingDirectory",  # 工作目录
]

# ==================== 环境变量模式 ====================
# 需要修复的环境变量（包含盘符引用的变量）
# noqa: F841 - 预留常量，用于将来的环境变量修复功能
PATH_CONTAINING_ENV_VARS = [
    "PATH",
    "CLASSPATH",
    "PYTHON_PATH",
    "PYTHONPATH",
]

# ==================== 备份保留策略 ====================
# 最多保留的备份版本数（超过此数量的旧备份会被清理）
# noqa: F841 - 预留常量，用于将来的旧备份清理功能
MAX_BACKUP_VERSIONS = 50

# 备份保留天数（超过此天数的备份标记为可清理）
# noqa: F841 - 预留常量，用于将来的旧备份清理功能
BACKUP_RETENTION_DAYS = 30
