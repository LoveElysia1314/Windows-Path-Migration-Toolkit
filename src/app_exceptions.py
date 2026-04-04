"""
自定义异常定义模块

定义应用中使用的所有自定义异常，便于统一错误处理和类型提示。

说明：
  - 部分异常类（标注了 noqa）暂时未在当前代码中使用，但作为完整 API 的一部分保留
  - 这些异常预留用于将来的功能扩展（如快捷方式处理、环境变量修复等）
  - noqa: F841 - 表示 unused 变量/类的 Flake8 警告代码

# noqa: F841
"""


class PathMigrationError(Exception):
    """
    路径迁移操作基础异常

    所有其他异常的父类，用于捕获所有相关错误。
    """

    pass


class AdminPrivilegeError(PathMigrationError):
    """
    缺少管理员权限异常

    当操作需要管理员权限但当前进程无管理员权限时抛出。
    """

    pass


class RegistryError(PathMigrationError):
    """
    注册表操作异常

    包括：注册表读取失败、写入失败、项不存在等。
    """

    pass


class RegistryKeyNotFoundError(RegistryError):  # noqa: F841
    """
    注册表键未找到

    指定的注册表键不存在。
    """

    pass


class RegistryValueError(RegistryError):  # noqa: F841
    """
    注册表值操作失败

    读取或修改注册表值时出错。
    """

    pass


class DriveLookupError(PathMigrationError):  # noqa: F841
    """
    盘符修复操作异常

    包括：盘符不存在、修复失败等。
    """

    pass


class MigrationNotFoundError(PathMigrationError):  # noqa: F841
    """
    迁移记录未找到

    指定的应用迁移记录不存在或已被删除。
    """

    pass


class MigrationAlreadyExistsError(PathMigrationError):  # noqa: F841
    """
    迁移记录已存在

    尝试创建重复的迁移记录时抛出。
    """

    pass


class PathValidationError(PathMigrationError):
    """
    路径验证失败

    路径不存在、无法访问或格式不正确。
    """

    pass


class InvalidPathError(PathValidationError):  # noqa: F841
    """
    无效的路径格式

    路径格式不符合 Windows 规范。
    """

    pass


class PathNotAccessibleError(PathValidationError):  # noqa: F841
    """
    路径无法访问

    路径存在但由于权限问题或其他原因无法访问。
    """

    pass


class BackupError(PathMigrationError):
    """
    备份操作异常

    包括：备份创建失败、备份文件损坏等。
    """

    pass


class BackupNotFoundError(BackupError):  # noqa: F841
    """
    备份文件未找到

    指定的备份不存在。
    """

    pass


class RestoreError(PathMigrationError):  # noqa: F841
    """
    恢复操作异常

    恢复迁移或盘符修复时出错。
    """

    pass


class ConfigurationError(PathMigrationError):  # noqa: F841
    """
    配置错误

    应用配置不正确或不完整。
    """

    pass


class DependencyError(PathMigrationError):  # noqa: F841
    """
    依赖项缺失或不可用

    包括：缺少 pywin32、PySide6 或其他必需模块。
    """

    pass


class ApplicationError(PathMigrationError):  # noqa: F841
    """
    通用应用错误

    不属于上述类别的其他错误。
    """

    pass
