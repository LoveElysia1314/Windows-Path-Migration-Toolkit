"""
Windows Path Migration Toolkit

一个面向 Windows 的路径迁移工具集，主要用于：
- 批量迁移已安装应用目录
- 修复盘符变更导致的注册表/环境变量/快捷方式路径失效
- 记录每次操作并支持回滚恢复

公共 API：
    PathConfig - 全局路径和配置管理
    setup_logger - 日志记录器工厂函数
"""

__version__ = "0.4.0"
__author__ = "Windows Path Migration Team"
__description__ = "Windows path migration and drive letter fix toolkit"

# 暴露公共 API (可选导入)
try:
    from .app_path_manager import PathConfig
    from .app_logger import setup_logger

    __all__ = [
        "PathConfig",
        "setup_logger",
        "__version__",
        "__author__",
    ]
except ImportError as e:
    # 如果导入失败，记录但不中断启动
    # 这在某些开发场景下可能发生
    import warnings

    warnings.warn(f"部分模块导入失败: {e}", ImportWarning)
    __all__ = ["__version__", "__author__"]
