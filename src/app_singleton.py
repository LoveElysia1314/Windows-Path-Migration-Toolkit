"""
单例模式工具模块

提供线程安全的单例元类，用于确保全局只有一个实例。
"""

import threading
from typing import Any, Dict, Type, TypeVar

T = TypeVar("T")


class Singleton(type):
    """
    线程安全的单例元类

    使用方法：
        class MyClass(metaclass=Singleton):
            def __init__(self):
                self.data = {}

        # 多次调用返回同一实例
        obj1 = MyClass()
        obj2 = MyClass()
        assert obj1 is obj2  # True

    特点：
        - 线程安全：使用双重检查锁定 (Double-Checked Locking) 避免竞态条件
        - 高效：只在第一次创建时加锁
        - 简洁：无需手动管理实例存储
    """

    _instances: Dict[Type, Any] = {}
    _lock: threading.Lock = threading.Lock()

    def __call__(cls, *args: Any, **kwargs: Any) -> Any:
        """
        创建或返回单例实例

        Args:
            *args: 传递给 __init__ 的位置参数
            **kwargs: 传递给 __init__ 的关键字参数

        Returns:
            cls: 该类的单例实例
        """
        # 第一次检查：避免非必要的加锁
        if cls not in Singleton._instances:
            # 加锁以确保线程安全
            with Singleton._lock:
                # 第二次检查：防止多线程竞态条件
                if cls not in Singleton._instances:
                    # 调用父类的 __call__ 创建实例
                    instance = super(Singleton, cls).__call__(*args, **kwargs)
                    Singleton._instances[cls] = instance

        return Singleton._instances[cls]

    @classmethod
    def clear_instances(mcs) -> None:
        """
        清清除所有单例实例（主要用于测试）

        警告：
            仅在测试时使用。生产环境不应调用此方法。
        """
        with Singleton._lock:
            Singleton._instances.clear()

    @classmethod
    def reset_instance(mcs, cls: Type[T]) -> None:
        """
        清除特定类的单例实例（主要用于测试）

        Args:
            cls: 要重置的类

        警告：
            仅在测试时使用。生产环境不应调用此方法。
        """
        with Singleton._lock:
            if cls in Singleton._instances:
                del Singleton._instances[cls]
