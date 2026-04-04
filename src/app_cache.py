"""
JSON 文件缓存管理

基于文件修改时间的简单缓存，避免频繁重读同一文件。

使用单例模式确保全局只有一个缓存实例。
"""

import os
import json
import threading
from app_logger import get_logger
from app_singleton import Singleton

logger = get_logger(__name__)


class JSONCache(metaclass=Singleton):
    """
    基于文件修改时间的 JSON 文件缓存 - 单例模式

    特点：
        - 单例模式：全局只有一个缓存实例
        - 线程安全：使用锁保护缓存访问
        - 智能淘汰：缓存满时移除最旧的条目 (LRU 策略)
        - 自动失效：检测文件修改时间，自动更新过期数据

    示例:
        >>> cache = JSONCache()
        >>> data = cache.load("config.json", default={})
        >>> cache.invalidate("config.json")

        # 或使用全局函数（向后兼容）
        >>> from app_cache import cache_load
        >>> data = cache_load("config.json")
    """

    def __init__(self, max_size: int = 128):
        """
        初始化缓存

        Args:
            max_size: 最大缓存条目数（默认 128）
        """
        self._cache = {}
        self._timestamps = {}
        self._max_size = max_size
        self._lock = threading.Lock()

    def load(self, path: str, default=None):
        """
        带缓存的 JSON 加载

        对于同一文件，若文件未修改则返回缓存数据。
        否则重新加载并缓存。

        Args:
            path: 文件路径
            default: 文件不存在或加载失败时的默认值

        Returns:
            dict: 加载的数据或默认值
        """
        if default is None:
            default = {}

        if not os.path.exists(path):
            return default

        try:
            mtime = os.path.getmtime(path)
        except OSError:
            return default

        # 线程安全地检查缓存
        with self._lock:
            # 检查缓存是否仍然有效
            if path in self._cache and self._timestamps.get(path) == mtime:
                logger.debug(f"缓存命中: {path}")
                return self._cache[path]

            # 重新加载文件
            try:
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)

                # 检查缓存大小，LRU 淘汰
                if len(self._cache) >= self._max_size:
                    # 移除最旧的条目
                    oldest = min(self._cache.keys(), key=lambda x: self._timestamps.get(x, 0))
                    del self._cache[oldest]
                    del self._timestamps[oldest]
                    logger.debug(f"缓存满，移除旧条目: {oldest}")

                self._cache[path] = data
                self._timestamps[path] = mtime
                logger.debug(f"缓存加载: {path}")
                return data

            except json.JSONDecodeError as e:
                logger.error(f"JSON 解析错误 {path}: {e}")
                return default
            except IOError as e:
                logger.error(f"IO 错误读取 {path}: {e}")
                return default

    def invalidate(self, path: str) -> None:
        """
        使特定文件的缓存失效

        Args:
            path: 要使其缓存失效的文件路径
        """
        with self._lock:
            if path in self._cache:
                del self._cache[path]
            if path in self._timestamps:
                del self._timestamps[path]
            logger.debug(f"缓存失效: {path}")

    def clear(self) -> None:
        """清空所有缓存"""
        with self._lock:
            self._cache.clear()
            self._timestamps.clear()
        logger.debug("缓存已清空")

    @property
    def size(self) -> int:
        """返回当前缓存大小"""
        return len(self._cache)

    def get_stats(self) -> dict:
        """
        获取缓存统计信息

        Returns:
            dict: 包含缓存统计的字典
        """
        return {
            "current_size": len(self._cache),
            "max_size": self._max_size,
            "cached_files": list(self._cache.keys()),
        }


# ==================== 向后兼容的全局函数接口 ====================
# 使用单例的全局实例来实现这些函数
_cache_singleton = JSONCache()


def cache_load(path: str, default=None):
    """
    全局缓存加载函数（向后兼容）

    Args:
        path: 文件路径
        default: 默认值

    Returns:
        dict: 加载的数据或默认值

    示例:
        >>> data = cache_load("config.json")
    """
    return _cache_singleton.load(path, default)


def cache_invalidate(path: str) -> None:
    """
    全局使缓存失效函数（向后兼容）

    Args:
        path: 要失效的文件路径

    示例:
        >>> cache_invalidate("config.json")
    """
    _cache_singleton.invalidate(path)


def cache_clear() -> None:
    """
    全局清空缓存函数

    示例:
        >>> cache_clear()
    """
    _cache_singleton.clear()


def get_cache() -> JSONCache:
    """
    获取全局缓存实例

    Returns:
        JSONCache: 全局单例缓存实例

    示例:
        >>> cache = get_cache()
        >>> stats = cache.get_stats()
    """
    return _cache_singleton
