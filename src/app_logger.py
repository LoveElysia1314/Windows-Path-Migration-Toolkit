"""
结构化日志管理模块

统一的日志配置，支持同时输出到控制台和文件，便于调试和审计。

使用单例模式确保全局只有一个日志管理器实例。
"""

import logging
import os
from datetime import datetime
from app_path_manager import PathConfig
from app_singleton import Singleton


class LoggerManager(metaclass=Singleton):
    """
    日志管理器 - 单例模式

    确保全局只有一个日志管理器实例，统一管理所有日志记录器。

    特点：
        - 单例模式：全局只有一个实例
        - 线程安全：多线程环境下安全使用
        - 统一配置：集中管理所有日志记录器
        - 避免重复：同名日志记录器只配置一次
    """

    def __init__(self):
        """初始化日志管理器"""
        self._loggers = {}
        self._lock = __import__("threading").Lock()

    def get_logger(self, name: str, log_file: str = None, level=logging.DEBUG):
        """
        获取或创建日志记录器

        Args:
            name: 日志记录器名称（通常为 __name__）
            log_file: 日志文件路径（可选）
            level: 日志级别

        Returns:
            logging.Logger: 配置好的日志记录器

        示例:
            >>> manager = LoggerManager()
            >>> logger = manager.get_logger(__name__)
            >>> logger.info("应用启动")
        """
        # 检查缓存，避免重复配置
        if name in self._loggers:
            return self._loggers[name]

        # 获取或创建 Python 原生日志记录器
        logger = logging.getLogger(name)
        logger.setLevel(level)

        # 如果已有处理器，直接返回（避免重复添加）
        if logger.handlers:
            with self._lock:
                self._loggers[name] = logger
            return logger

        # ==================== 配置控制台处理器 ====================
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_formatter = logging.Formatter(
            "%(asctime)s | %(levelname)-8s | %(message)s", datefmt="%H:%M:%S"
        )
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)

        # ==================== 配置文件处理器 ====================
        if log_file is None:
            # 使用默认日志文件路径
            PathConfig.ensure_directories()
            log_file = os.path.join(
                PathConfig.LOG_DIR, f"app_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
            )

        try:
            os.makedirs(os.path.dirname(log_file), exist_ok=True)
            file_handler = logging.FileHandler(log_file, encoding="utf-8")
            file_handler.setLevel(logging.DEBUG)
            file_formatter = logging.Formatter(
                "%(asctime)s | %(name)-20s | %(levelname)-8s | %(message)s",
                datefmt="%Y-%m-%d %H:%M:%S",
            )
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)
        except Exception as e:
            print(f"[WARN] 无法创建日志文件 {log_file}: {e}")

        # 缓存日志记录器
        with self._lock:
            self._loggers[name] = logger

        return logger

    def get_all_loggers(self) -> dict:
        """
        获取所有已配置的日志记录器

        Returns:
            dict: 日志记录器名称到对象的映射
        """
        return dict(self._loggers)

    def reset(self):
        """
        重置日志管理器（主要用于测试）

        警告：
            仅在测试时使用。生产环境不应调用此方法。
        """
        with self._lock:
            self._loggers.clear()


# ==================== 全局单例实例 ====================
_logger_manager = LoggerManager()


# ==================== 向后兼容的函数接口 ====================
def setup_logger(name, log_file=None, level=logging.DEBUG):
    """
    配置并返回一个结构化日志记录器

    这是对 LoggerManager.get_logger() 的包装，用于向后兼容。

    Args:
        name: 日志记录器名称（通常为 __name__）
        log_file: 日志文件路径（可选）
        level: 日志级别

    Returns:
        logging.Logger: 配置好的日志记录器

    示例:
        >>> logger = setup_logger(__name__)
        >>> logger.info("应用启动")
        >>> logger.debug("调试信息")
        >>> logger.error("错误信息")
    """
    return _logger_manager.get_logger(name, log_file, level)


def get_logger(name):
    """
    获取已配置的日志记录器

    这是对 LoggerManager.get_logger() 的包装，用于向后兼容。

    如果记录器不存在，则进行配置

    Args:
        name: 日志记录器名称

    Returns:
        logging.Logger: 日志记录器
    """
    return _logger_manager.get_logger(name)
