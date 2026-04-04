import ctypes
import json
import os
import re
import shutil
import struct
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

try:
    import winreg
except ImportError:
    print("[错误] 当前环境不支持 winreg，仅可在 Windows 上运行。")
    sys.exit(1)

try:
    import win32com.client
except ImportError:
    print("[错误] 缺少依赖: pywin32")
    print("请先执行: pip install pywin32")
    sys.exit(1)

# 导入统一的配置和日志
from app_path_manager import PathConfig
from app_logger import setup_logger
from app_cache import cache_load, cache_invalidate
from app_constants import (
    UNINSTALL_REGISTRY_LOCATIONS,
    ENVIRONMENT_REGISTRY_LOCATIONS,
    FLAT_COPY_EXCLUDED_FOLDERS,
    CONTROL_PANEL_EXCLUDED_RELEASE_TYPES,
    DEFAULT_EXCLUDED_VENDOR_KEYWORDS,
    STANDARD_INSTALL_PATH_PATTERN,
    DRIVE_FIX_SEARCH_REGISTRY_ROOTS,
)
from app_exceptions import (
    RegistryError,
    PathValidationError,
)

logger = setup_logger(__name__)

# ==================== 向后兼容的常量别名 ====================
# 这些常量现已从 app_constants.py 导入，保留别名以兼容现有代码
UNINSTALL_ROOTS = UNINSTALL_REGISTRY_LOCATIONS
ENVIRONMENT_ROOTS = ENVIRONMENT_REGISTRY_LOCATIONS
CONTROL_PANEL_EXCLUDE_RELEASE_TYPES = CONTROL_PANEL_EXCLUDED_RELEASE_TYPES
STANDARD_INSTALL_PATH_PATTERN = STANDARD_INSTALL_PATH_PATTERN
DEFAULT_EXCLUDE_VENDOR_KEYWORDS = DEFAULT_EXCLUDED_VENDOR_KEYWORDS
FLAT_COPY_TOP_FOLDERS = FLAT_COPY_EXCLUDED_FOLDERS
DRIVE_FIX_SEARCH_ROOTS = DRIVE_FIX_SEARCH_REGISTRY_ROOTS


def is_admin() -> bool:
    """
    检查当前进程是否具有管理员权限

    Returns:
        bool: True 为管理员权限，False 为非管理员
    """
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def relaunch_as_admin() -> bool:
    """
    以管理员权限重启当前程序

    Returns:
        bool: 重启成功返回 True，失败返回 False
    """
    script = os.path.abspath(sys.argv[0])
    params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    ret = ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, f'"{script}" {params}', None, 1
    )
    return ret > 32


def ensure_admin_or_exit() -> None:
    """
    检查管理员权限，不满足则直接退出程序

    Note:
        这是向后兼容版本。新代码推荐直接捕获异常。
        参考: USAGE_GUIDE_AFTER_OPTIMIZATION.md 中的"异常处理"部分
    """
    if is_admin():
        return
    print("检测到当前非管理员权限，正在自动提权...")
    if relaunch_as_admin():
        sys.exit(0)
    print("提权失败，请手动以管理员身份运行。")
    sys.exit(1)


def ensure_dir(path: str) -> None:
    """
    确保目录存在，不存在则创建

    Args:
        path: 目录路径
    """
    os.makedirs(path, exist_ok=True)


def load_json(path: str, default: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """
    加载 JSON 文件（使用缓存）

    Args:
        path: 文件路径
        default: 默认值（文件不存在或加载失败时）

    Returns:
        dict: 加载的数据或默认值
    """
    if default is None:
        default = {}
    return cache_load(path, default)


def save_json(path: str, data: Dict[str, Any]) -> None:
    """
    原子性保存 JSON 文件

    使用临时文件实现写入的原子性，防止数据损坏。

    Args:
        path: 文件路径
        data: 要保存的数据

    Raises:
        IOError: 如果写入失败
    """
    ensure_dir(os.path.dirname(path))

    temp_path = path + ".tmp"
    backup_path = path + ".bak"

    try:
        # 写入临时文件
        with open(temp_path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

        # 备份现有文件
        if os.path.exists(path):
            if os.path.exists(backup_path):
                os.remove(backup_path)
            os.rename(path, backup_path)

        # 原子性替换
        os.rename(temp_path, path)
        logger.debug(f"已保存 JSON: {path}")

        # 失效缓存
        cache_invalidate(path)

    except Exception as e:
        logger.error(f"保存 JSON 失败 {path}: {e}")
        # 尝试恢复备份
        if os.path.exists(backup_path) and not os.path.exists(path):
            os.rename(backup_path, path)
            logger.warning(f"已从备份恢复: {backup_path}")
        raise

    finally:
        # 清理临时文件
        if os.path.exists(temp_path):
            try:
                os.remove(temp_path)
            except Exception:
                pass


def now_batch_id():
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def path_norm(p):
    return os.path.normcase(os.path.normpath(p))


def sanitize_name(name):
    return re.sub(r"[\\/:*?\"<>|]", "_", name).strip("._ ") or "UnnamedApp"


def parse_possible_path(raw):
    if not raw or not isinstance(raw, str):
        return ""

    text = raw.strip().strip("\x00")

    m = re.search(r'"([A-Za-z]:\\[^\"]+)"', text)
    if m:
        return m.group(1)

    m = re.search(r"([A-Za-z]:\\[^\s]+\.(?:exe|msi))", text, re.IGNORECASE)
    if m:
        return m.group(1)

    if re.match(r"^[A-Za-z]:\\", text):
        return text

    return ""


def extract_install_dir(install_location, display_icon, uninstall_str):
    if isinstance(install_location, str):
        p = install_location.strip().strip('"')
        if p and os.path.isdir(p):
            return os.path.abspath(p)

    icon_path = parse_possible_path(display_icon)
    if icon_path:
        icon_path = icon_path.split(",")[0]
        icon_path = icon_path.strip('"')
        if os.path.isfile(icon_path):
            return os.path.dirname(os.path.abspath(icon_path))
        if os.path.isdir(icon_path):
            return os.path.abspath(icon_path)

    uninstall_path = parse_possible_path(uninstall_str)
    if uninstall_path:
        uninstall_path = uninstall_path.strip('"')
        if os.path.isfile(uninstall_path):
            return os.path.dirname(os.path.abspath(uninstall_path))
        if os.path.isdir(uninstall_path):
            return os.path.abspath(uninstall_path)

    return ""


def is_in_standard_install_path(install_dir):
    if not isinstance(install_dir, str) or not install_dir:
        return False
    return bool(STANDARD_INSTALL_PATH_PATTERN.match(os.path.abspath(install_dir)))


def should_exclude_by_keywords(keywords, display_name, publisher, install_dir, subkey_name):
    if not keywords:
        return False

    text = " ".join(
        [
            str(display_name or ""),
            str(publisher or ""),
            str(install_dir or ""),
            str(subkey_name or ""),
        ]
    ).lower()
    return any(k and k.lower() in text for k in keywords)


def get_program_files_roots():
    root_x64 = os.path.abspath(
        os.environ.get("ProgramW6432") or os.environ.get("ProgramFiles") or r"D:\Program Files"
    )
    root_x86 = os.path.abspath(os.environ.get("ProgramFiles(x86)") or r"D:\Program Files (x86)")
    return root_x64, root_x86


def detect_executable_architecture(exe_path):
    try:
        with open(exe_path, "rb") as f:
            if f.read(2) != b"MZ":
                return ""
            f.seek(0x3C)
            pe_offset_data = f.read(4)
            if len(pe_offset_data) != 4:
                return ""
            pe_offset = struct.unpack("<I", pe_offset_data)[0]
            f.seek(pe_offset)
            if f.read(4) != b"PE\x00\x00":
                return ""
            machine_data = f.read(2)
            if len(machine_data) != 2:
                return ""
            machine = struct.unpack("<H", machine_data)[0]
            if machine == 0x8664:
                return "x64"
            if machine == 0x14C:
                return "x86"
    except Exception:
        return ""
    return ""


def detect_app_architecture(app):
    install_dir = str(app.get("install_dir", "") or "")
    if not install_dir or not os.path.isdir(install_dir):
        return "x86" if str(app.get("view", "")) == "32" else "x64"

    lower_dir = path_norm(install_dir)
    if "\\program files (x86)\\" in lower_dir or lower_dir.endswith("\\program files (x86)"):
        return "x86"

    exe_candidates = []
    try:
        for name in os.listdir(install_dir):
            if not name.lower().endswith(".exe"):
                continue
            full = os.path.join(install_dir, name)
            if not os.path.isfile(full):
                continue
            try:
                size = os.path.getsize(full)
            except OSError:
                size = 0
            exe_candidates.append((size, full))
    except OSError:
        exe_candidates = []

    exe_candidates.sort(key=lambda x: x[0], reverse=True)
    exe_candidates = exe_candidates[:20]

    has_x64 = False
    has_x86 = False
    for _, exe in exe_candidates:
        arch = detect_executable_architecture(exe)
        if arch == "x64":
            has_x64 = True
        elif arch == "x86":
            has_x86 = True
        if has_x64 and has_x86:
            break

    if has_x64 and not has_x86:
        return "x64"
    if has_x86 and not has_x64:
        return "x86"

    return "x86" if str(app.get("view", "")) == "32" else "x64"


def choose_target_root_for_app(
    app, manual_root, auto_arch, target_root_x64=None, target_root_x86=None
):
    if not auto_arch:
        return os.path.abspath(manual_root), "manual"

    sys_x64, sys_x86 = get_program_files_roots()
    root_x64 = os.path.abspath(target_root_x64 or sys_x64)
    root_x86 = os.path.abspath(target_root_x86 or sys_x86)
    arch = detect_app_architecture(app)
    return (root_x86 if arch == "x86" else root_x64), arch


def choose_destination_subpath(src, preserve_relative_layout=True):
    src_abs = os.path.abspath(src)
    if not preserve_relative_layout:
        return os.path.basename(src_abs)

    if is_in_standard_install_path(src_abs):
        return os.path.basename(src_abs)

    drive, tail = os.path.splitdrive(src_abs)
    rel = tail.lstrip("\\/")
    if not rel:
        return os.path.basename(src_abs)

    first = rel.split("\\", 1)[0].lower()
    if first in FLAT_COPY_TOP_FOLDERS:
        return os.path.basename(src_abs)

    return rel


def build_migration_preview(
    selected_apps,
    target_root,
    auto_arch=False,
    target_root_x64=None,
    target_root_x86=None,
    preserve_relative_layout=True,
):
    manual_target_root = os.path.abspath(target_root)
    resolved_x64, resolved_x86 = get_program_files_roots()
    if target_root_x64:
        resolved_x64 = os.path.abspath(target_root_x64)
    if target_root_x86:
        resolved_x86 = os.path.abspath(target_root_x86)

    preview = []
    for app in selected_apps:
        src = os.path.abspath(app["install_dir"])
        app_target_root, detected_arch = choose_target_root_for_app(
            app,
            manual_target_root,
            auto_arch,
            target_root_x64=resolved_x64,
            target_root_x86=resolved_x86,
        )
        dst_subpath = choose_destination_subpath(
            src, preserve_relative_layout=preserve_relative_layout
        )
        dst = os.path.abspath(os.path.join(app_target_root, dst_subpath))
        preview.append(
            {
                "name": app.get("display_name", ""),
                "src": src,
                "detected_arch": detected_arch,
                "target_root_used": app_target_root,
                "dst_subpath": dst_subpath,
                "dst": dst,
            }
        )

    return preview


def read_reg_value(key, name):
    try:
        val, _ = winreg.QueryValueEx(key, name)
        return val
    except OSError:
        return ""


def read_reg_dword(key, name, default=0):
    try:
        val, _ = winreg.QueryValueEx(key, name)
    except OSError:
        return default

    if isinstance(val, int):
        return val
    if isinstance(val, str) and val.isdigit():
        return int(val)
    return default


def normalize_install_date(raw):
    if raw is None:
        return ""

    text = str(raw).strip()
    if not text:
        return ""

    digits = re.sub(r"\D", "", text)
    if len(digits) != 8:
        return ""

    y = digits[0:4]
    m = digits[4:6]
    d = digits[6:8]
    try:
        yi = int(y)
        mi = int(m)
        di = int(d)
    except ValueError:
        return ""

    if yi < 1990 or yi > 2100:
        return ""
    if mi < 1 or mi > 12:
        return ""
    if di < 1 or di > 31:
        return ""

    return f"{y}-{m}-{d}"


def is_store_or_builtin_entry(
    subkey_name, display_name, install_location, display_icon, uninstall_str
):
    text = " ".join(
        [
            str(subkey_name or ""),
            str(display_name or ""),
            str(install_location or ""),
            str(display_icon or ""),
            str(uninstall_str or ""),
        ]
    ).lower()

    markers = [
        "windowsapps",
        "shell:appsfolder",
        "ms-resource:",
        "appx",
        "microsoft.windows",
        "explorer.exe shell:appsfolder",
    ]
    return any(m in text for m in markers)


def is_control_panel_visible_entry(
    app_key, subkey_name, display_name, install_location, display_icon, uninstall_str
):
    if not display_name:
        return False

    display_name = str(display_name).strip()
    if not display_name:
        return False

    if display_name.startswith("@{"):
        return False

    if read_reg_dword(app_key, "SystemComponent", 0) == 1:
        return False

    if read_reg_dword(app_key, "NoDisplay", 0) == 1:
        return False

    if read_reg_value(app_key, "ParentKeyName"):
        return False

    if read_reg_value(app_key, "ParentDisplayName"):
        return False

    release_type = str(read_reg_value(app_key, "ReleaseType") or "").strip().lower()
    if release_type in CONTROL_PANEL_EXCLUDE_RELEASE_TYPES:
        return False

    if is_store_or_builtin_entry(
        subkey_name, display_name, install_location, display_icon, uninstall_str
    ):
        return False

    return True


def enum_installed_apps(exclude_standard_paths=False, exclude_keywords=None):
    if exclude_keywords is None:
        exclude_keywords = list(DEFAULT_EXCLUDE_VENDOR_KEYWORDS)

    apps = []

    for hive, root_path, view in UNINSTALL_ROOTS:
        try:
            root = winreg.OpenKey(hive, root_path, 0, winreg.KEY_READ)
        except OSError:
            continue

        idx = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(root, idx)
                idx += 1
            except OSError:
                break

            full_sub = root_path + "\\" + subkey_name
            try:
                app_key = winreg.OpenKey(hive, full_sub, 0, winreg.KEY_READ)
            except OSError:
                continue

            try:
                display_name = read_reg_value(app_key, "DisplayName")
                if not display_name:
                    continue

                install_location = read_reg_value(app_key, "InstallLocation")
                display_icon = read_reg_value(app_key, "DisplayIcon")
                uninstall_str = read_reg_value(app_key, "UninstallString")
                publisher = read_reg_value(app_key, "Publisher")
                install_date_raw = read_reg_value(app_key, "InstallDate")
                install_date = normalize_install_date(install_date_raw)
                icon_path = parse_possible_path(display_icon)
                if icon_path:
                    icon_path = os.path.expandvars(icon_path.split(",")[0].strip('"'))

                if not is_control_panel_visible_entry(
                    app_key,
                    subkey_name,
                    display_name,
                    install_location,
                    display_icon,
                    uninstall_str,
                ):
                    continue

                install_dir = extract_install_dir(install_location, display_icon, uninstall_str)
                if not install_dir:
                    continue
                if not os.path.isdir(install_dir):
                    continue
                if exclude_standard_paths and is_in_standard_install_path(install_dir):
                    continue
                if should_exclude_by_keywords(
                    exclude_keywords,
                    display_name,
                    publisher,
                    install_dir,
                    subkey_name,
                ):
                    continue

                arch = detect_app_architecture({"install_dir": install_dir, "view": view})

                apps.append(
                    {
                        "display_name": str(display_name),
                        "install_dir": install_dir,
                        "publisher": str(publisher or ""),
                        "arch": arch,
                        "install_date": install_date,
                        "display_icon": icon_path if icon_path else "",
                        "hive": "HKLM" if hive == winreg.HKEY_LOCAL_MACHINE else "HKCU",
                        "reg_subkey": full_sub,
                        "view": view,
                    }
                )
            finally:
                winreg.CloseKey(app_key)

        winreg.CloseKey(root)

    dedup = {}
    for app in apps:
        key = (path_norm(app["install_dir"]), app["display_name"].lower())
        if key not in dedup:
            dedup[key] = app

    return sorted(
        dedup.values(),
        key=lambda x: (x["install_dir"].lower(), x["display_name"].lower()),
    )


def export_reg_key(hive_name, subkey, out_file):
    full = ("HKEY_LOCAL_MACHINE" if hive_name == "HKLM" else "HKEY_CURRENT_USER") + "\\" + subkey
    ensure_dir(os.path.dirname(out_file))
    try:
        subprocess.run(
            ["reg", "export", full, out_file, "/y"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except subprocess.CalledProcessError:
        return False


def replace_in_registry_values(hive_name, subkey, view, old_path, new_path):
    root = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER
    access_read = winreg.KEY_READ
    access_write = winreg.KEY_SET_VALUE

    view_flag = winreg.KEY_WOW64_64KEY if view == "64" else winreg.KEY_WOW64_32KEY

    changed = 0
    failed = 0
    details = []

    try:
        key_r = winreg.OpenKey(root, subkey, 0, access_read | view_flag)
        key_w = winreg.OpenKey(root, subkey, 0, access_write | view_flag)
    except OSError:
        return 0, 1, details

    old_norm = path_norm(old_path)
    pattern = re.compile(re.escape(old_path), re.IGNORECASE)

    i = 0
    while True:
        try:
            name, data, typ = winreg.EnumValue(key_r, i)
            i += 1
        except OSError:
            break

        if typ not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ):
            continue
        if not isinstance(data, str):
            continue
        if old_norm not in path_norm(data):
            continue

        new_data = pattern.sub(lambda _m: new_path, data)
        if new_data == data:
            continue

        try:
            winreg.SetValueEx(key_w, name, 0, typ, new_data)
            changed += 1
            details.append({"name": name, "old": data, "new": new_data})
        except OSError:
            failed += 1

    winreg.CloseKey(key_r)
    winreg.CloseKey(key_w)

    return changed, failed, details


def replace_in_environment_values(old_path, new_path):
    old_norm = path_norm(old_path)
    pattern = re.compile(re.escape(old_path), re.IGNORECASE)

    changed = 0
    failed = 0
    details = []

    for hive_name, subkey in ENVIRONMENT_ROOTS:
        root = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER

        try:
            key_r = winreg.OpenKey(root, subkey, 0, winreg.KEY_READ)
            key_w = winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE)
        except OSError:
            continue

        i = 0
        while True:
            try:
                name, data, typ = winreg.EnumValue(key_r, i)
                i += 1
            except OSError:
                break

            if typ not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ):
                continue
            if not isinstance(data, str):
                continue
            if old_norm not in path_norm(data):
                continue

            new_data = pattern.sub(lambda _m: new_path, data)
            if new_data == data:
                continue

            try:
                winreg.SetValueEx(key_w, name, 0, typ, new_data)
                changed += 1
                details.append(
                    {
                        "hive": hive_name,
                        "subkey": subkey,
                        "name": name,
                        "old": data,
                        "new": new_data,
                    }
                )
            except OSError:
                failed += 1

        winreg.CloseKey(key_r)
        winreg.CloseKey(key_w)

    return changed, failed, details


def backup_environment_keys(backup_base, app_name):
    out_files = []
    env_backup_dir = os.path.join(backup_base, "environment")
    ensure_dir(env_backup_dir)

    safe = sanitize_name(app_name)
    for hive_name, subkey in ENVIRONMENT_ROOTS:
        out_file = os.path.join(env_backup_dir, f"{safe}_{hive_name}_{sanitize_name(subkey)}.reg")
        if export_reg_key(hive_name, subkey, out_file):
            out_files.append(out_file)

    return out_files


def candidate_shortcut_roots(scan_all_users=False):
    """
    获取快捷方式根目录

    优先仅扫描当前用户，可选扫描其他用户。

    Args:
        scan_all_users: 是否扫描所有用户（默认False）

    Returns:
        list: 存在的快捷方式目录列表
    """
    paths = []

    # 1. 当前用户的快捷方式（最常被修改）
    user_profile = os.environ.get("USERPROFILE", "")
    if user_profile:
        desktop = os.path.join(user_profile, "Desktop")
        if os.path.isdir(desktop):
            paths.append(desktop)

        programs = os.path.join(
            user_profile, r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs"
        )
        if os.path.isdir(programs):
            paths.append(programs)

    # 2. 公共快捷方式
    public = os.environ.get("PUBLIC", "")
    if public:
        public_desktop = os.path.join(public, "Desktop")
        if os.path.isdir(public_desktop):
            paths.append(public_desktop)

        public_programs = os.path.join(public, r"Microsoft\Windows\Start Menu\Programs")
        if os.path.isdir(public_programs):
            paths.append(public_programs)

    # 3. 其他用户的快捷方式（可选，避免不必要的扫描）
    if scan_all_users:
        system_drive = os.environ.get("SystemDrive", "C:")
        users_root = os.path.join(system_drive, "Users")

        # 排除系统用户
        SYSTEM_USERS = {"Default", "All Users", "Public", "defaultuser0"}

        if os.path.isdir(users_root):
            try:
                for user_name in os.listdir(users_root):
                    if user_name in SYSTEM_USERS:
                        continue
                    if user_name.startswith("."):  # 隐藏项
                        continue

                    user_programs = os.path.join(
                        users_root,
                        user_name,
                        r"AppData\Roaming\Microsoft\Windows\Start Menu\Programs",
                    )
                    if os.path.isdir(user_programs):
                        paths.append(user_programs)
            except OSError:
                logger.warning("无法枚举 Users 目录")

    # 去重并返回
    seen = set()
    unique_paths = []
    for p in paths:
        pn = path_norm(p)
        if pn not in seen:
            seen.add(pn)
            unique_paths.append(p)

    logger.debug(f"快捷方式扫描根目录数: {len(unique_paths)}")
    return unique_paths


def backup_shortcut_file(src, backup_root):
    drive, tail = os.path.splitdrive(src)
    rel = tail.lstrip("\\/")
    out = os.path.join(backup_root, "shortcuts", drive.replace(":", ""), rel + ".bak")
    ensure_dir(os.path.dirname(out))
    shutil.copy2(src, out)
    return out


def update_shortcuts_for_path(old_path, new_path, backup_root, max_files=5000):
    """
    更新快捷方式中的路径

    Args:
        old_path: 旧路径
        new_path: 新路径
        backup_root: 备份目录
        max_files: 最大扫描文件数（防止卡顿）

    Returns:
        dict: 包含 scanned, changed, failed, changes 的统计字典
    """
    shell = win32com.client.Dispatch("WScript.Shell")
    old_norm = path_norm(old_path)
    changed = 0
    failed = 0
    scanned = 0
    records = []

    pattern = re.compile(re.escape(old_path), re.IGNORECASE)

    logger.info(f"开始快捷方式更新: {old_path} -> {new_path}")

    for root in candidate_shortcut_roots(scan_all_users=False):
        if scanned >= max_files:
            logger.warning(f"快捷方式扫描数达到限制 {max_files}")
            break

        logger.debug(f"扫描快捷方式目录: {root}")

        for dirpath, _, filenames in os.walk(
            root, onerror=lambda e: logger.debug(f"目录遍历出错 {dirpath}: {e}")
        ):
            if scanned >= max_files:
                break

            for fn in filenames:
                if not fn.lower().endswith(".lnk"):
                    continue

                scanned += 1
                full = os.path.join(dirpath, fn)
                try:
                    sc = shell.CreateShortCut(full)
                except Exception as e:
                    logger.debug(f"无法打开快捷方式 {full}: {e}")
                    failed += 1
                    continue

                modified = False
                dif = {}
                for attr in ("TargetPath", "WorkingDirectory", "Arguments"):
                    val = getattr(sc, attr, "")
                    if not isinstance(val, str):
                        continue
                    if old_norm not in path_norm(val):
                        continue

                    new_val = pattern.sub(lambda _m: new_path, val)
                    if new_val != val:
                        setattr(sc, attr, new_val)
                        modified = True
                        dif[attr] = {"old": val, "new": new_val}

                if not modified:
                    continue

                try:
                    bak = backup_shortcut_file(full, backup_root)
                    sc.Save()
                    changed += 1
                    records.append({"path": full, "backup": bak, "diff": dif})
                    logger.debug(f"已更新快捷方式: {full}")
                except Exception as e:
                    logger.warning(f"保存快捷方式失败 {full}: {e}")
                    failed += 1

    del shell

    logger.info(f"快捷方式更新完成: 扫描={scanned}, 修改={changed}, 失败={failed}")
    return {
        "scanned": scanned,
        "changed": changed,
        "failed": failed,
        "changes": records,
    }


def copy_app_dir(src, dst):
    ensure_dir(os.path.dirname(dst))
    shutil.copytree(src, dst, dirs_exist_ok=True)


def try_remove_old_dir(path):
    try:
        shutil.rmtree(path)
        return True, ""
    except Exception as e:
        return False, str(e)


def load_manifest():
    data = load_json(PathConfig.MANIFEST_FILE, {"batches": []})
    if not isinstance(data, dict) or "batches" not in data:
        return {"batches": []}
    return data


def save_manifest(data):
    save_json(PathConfig.MANIFEST_FILE, data)


def append_batch(batch):
    m = load_manifest()
    m["batches"].append(batch)
    save_manifest(m)


def update_batch(updated_batch):
    m = load_manifest()
    batches = m.get("batches", [])
    for i, b in enumerate(batches):
        if b.get("id") == updated_batch.get("id"):
            batches[i] = updated_batch
            save_manifest(m)
            return True
    return False


def list_batches(applied_only=False):
    batches = load_manifest().get("batches", [])
    if applied_only:
        batches = [b for b in batches if b.get("status") == "applied"]
    return batches


def delete_migration_record(record_id, delete_backup=False):
    record_id = str(record_id or "").strip()
    if not record_id:
        return False, "invalid_id"

    m = load_manifest()
    batches = m.get("batches", [])
    idx = -1
    target = None
    for i, b in enumerate(batches):
        if b.get("id") == record_id:
            idx = i
            target = b
            break

    if idx < 0 or target is None:
        return False, "not_found"

    del batches[idx]
    save_manifest(m)

    pending = load_pending_cleanup()
    items = pending.get("items", [])
    pending["items"] = [it for it in items if str(it.get("batch_id", "")) != record_id]
    save_pending_cleanup(pending)

    if delete_backup:
        backup_base = str(target.get("backup_base", "") or "")
        if backup_base and os.path.exists(backup_base):
            try:
                shutil.rmtree(backup_base)
            except Exception:
                return True, "record_deleted_backup_failed"

    return True, "ok"


def load_pending_cleanup():
    return load_json(PathConfig.PENDING_CLEAN_FILE, {"items": []})


def save_pending_cleanup(data):
    save_json(PathConfig.PENDING_CLEAN_FILE, data)


def add_pending_cleanup(path, reason, batch_id, app_name):
    data = load_pending_cleanup()
    data.setdefault("items", [])
    data["items"].append(
        {
            "path": path,
            "reason": reason,
            "batch_id": batch_id,
            "app_name": app_name,
            "created_at": datetime.now().isoformat(),
        }
    )
    save_pending_cleanup(data)


def perform_cleanup_pending():
    data = load_pending_cleanup()
    items = data.get("items", [])
    if not items:
        print("当前没有待清理路径。")
        return {
            "ok": 0,
            "fail": 0,
            "remaining": 0,
            "failed_items": [],
        }

    remain = []
    ok = 0
    fail = 0
    failed_items = []

    for it in items:
        p = it.get("path", "")
        if not p:
            continue

        if not os.path.exists(p):
            ok += 1
            continue

        success, reason = try_remove_old_dir(p)
        if success:
            ok += 1
        else:
            it["last_error"] = reason
            it["last_try_at"] = datetime.now().isoformat()
            remain.append(it)
            fail += 1
            failed_items.append(
                {
                    "path": p,
                    "reason": reason,
                    "batch_id": it.get("batch_id", ""),
                    "app_name": it.get("app_name", ""),
                }
            )

    save_pending_cleanup({"items": remain})
    print(f"清理完成: 成功 {ok} | 失败 {fail} | 剩余 {len(remain)}")
    if failed_items:
        print("清理失败详情:")
        for i, item in enumerate(failed_items, 1):
            print(f"{i}. 路径: {item['path']}")
            print(f"   应用: {item.get('app_name', '')}")
            print(f"   记录ID: {item.get('batch_id', '')}")
            print(f"   原因: {item['reason']}")

    return {
        "ok": ok,
        "fail": fail,
        "remaining": len(remain),
        "failed_items": failed_items,
    }


def schedule_cleanup_on_reboot(paths):
    candidates = []
    seen = set()
    for p in paths or []:
        pp = os.path.abspath(str(p or "").strip())
        if not pp:
            continue
        key = path_norm(pp)
        if key in seen:
            continue
        seen.add(key)
        candidates.append(pp)

    if not candidates:
        return False, "no_paths", {}

    task_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    task_name = f"AppMigrationCleanup_{task_id}"
    jobs_dir = PathConfig.PENDING_CLEANUP_DIR
    ensure_dir(jobs_dir)
    script_path = os.path.abspath(os.path.join(jobs_dir, f"cleanup_{task_id}.ps1"))

    ps_paths = [p.replace("'", "''") for p in candidates]
    ps_lines = [
        "$ErrorActionPreference = 'SilentlyContinue'",
        "$paths = @(",
    ]
    for p in ps_paths:
        ps_lines.append(f"    '{p}'")
    ps_lines.extend(
        [
            ")",
            "foreach ($p in $paths) {",
            "    if (-not (Test-Path -LiteralPath $p)) { continue }",
            '    cmd /c "takeown /F ""$p"" /A /R /D Y >nul 2>nul" | Out-Null',
            '    cmd /c "icacls ""$p"" /grant Administrators:F /T /C >nul 2>nul" | Out-Null',
            '    cmd /c "attrib -R -S -H ""$p"" /S /D >nul 2>nul" | Out-Null',
            '    cmd /c "rmdir /S /Q ""$p""" | Out-Null',
            "}",
            f'schtasks /Delete /TN "{task_name}" /F | Out-Null',
            "Remove-Item -LiteralPath $MyInvocation.MyCommand.Path -Force -ErrorAction SilentlyContinue",
        ]
    )

    try:
        with open(script_path, "w", encoding="utf-8") as f:
            f.write("\n".join(ps_lines) + "\n")
    except Exception:
        return False, "write_script_failed", {}

    try:
        subprocess.run(
            [
                "schtasks",
                "/Create",
                "/TN",
                task_name,
                "/SC",
                "ONSTART",
                "/RU",
                "SYSTEM",
                "/RL",
                "HIGHEST",
                "/TR",
                f'powershell -NoProfile -ExecutionPolicy Bypass -File "{script_path}"',
                "/F",
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
    except subprocess.CalledProcessError:
        return False, "create_task_failed", {"script_path": script_path}

    return (
        True,
        "ok",
        {"task_name": task_name, "script_path": script_path, "paths": candidates},
    )


def import_reg_file(reg_file):
    try:
        if not reg_file or not os.path.exists(reg_file):
            return False
        subprocess.run(
            ["reg", "import", reg_file],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
        )
        return True
    except subprocess.CalledProcessError:
        return False


def restore_shortcuts_for_app(app_record):
    ok = 0
    fail = 0
    for item in app_record.get("shortcuts", {}).get("changes", []):
        target = item.get("path")
        backup = item.get("backup")
        try:
            if not target or not backup or not os.path.exists(backup):
                fail += 1
                continue
            ensure_dir(os.path.dirname(target))
            shutil.copy2(backup, target)
            ok += 1
        except Exception:
            fail += 1
    return ok, fail


def restore_registry_for_app(app_record):
    bak = app_record.get("registry", {}).get("backup_file", "")
    if not bak:
        return 0, 0
    return (1, 0) if import_reg_file(bak) else (0, 1)


def restore_environment_for_app(app_record):
    files = app_record.get("environment", {}).get("backup_files", [])
    ok = 0
    fail = 0
    for f in files:
        if import_reg_file(f):
            ok += 1
        else:
            fail += 1
    return ok, fail


def restore_services_for_app(app_record):
    files = app_record.get("services", {}).get("backup_files", [])
    ok = 0
    fail = 0
    for f in files:
        if import_reg_file(f):
            ok += 1
        else:
            fail += 1
    return ok, fail


def restore_tasks_for_app(app_record):
    files = app_record.get("tasks", {}).get("backup_files", [])
    ok = 0
    fail = 0
    for f in files:
        if import_reg_file(f):
            ok += 1
        else:
            fail += 1
    return ok, fail


def restore_program_path_for_app(app_record):
    src = app_record.get("src", "")
    dst = app_record.get("dst", "")
    if app_record.get("copy") != "ok":
        return False, "skip_not_copied"
    if not dst or not os.path.exists(dst):
        return False, "dst_missing"

    try:
        ensure_dir(os.path.dirname(src))
        shutil.copytree(dst, src, dirs_exist_ok=True)
    except Exception as e:
        return False, f"copy_back_failed: {e}"

    try:
        shutil.rmtree(dst)
        return True, "ok"
    except Exception:
        return True, "ok_dst_not_removed"


def restore_migration_batch(batch_id=None):
    batches = list_batches(applied_only=True)
    if not batches:
        return None, "no_applied_batch"

    target = None
    if batch_id:
        for b in batches:
            if b.get("id") == batch_id:
                target = b
                break
        if target is None:
            return None, "batch_not_found"
    else:
        target = batches[-1]

    result = {
        "program_path_ok": 0,
        "program_path_fail": 0,
        "registry_ok": 0,
        "registry_fail": 0,
        "environment_ok": 0,
        "environment_fail": 0,
        "services_ok": 0,
        "services_fail": 0,
        "tasks_ok": 0,
        "tasks_fail": 0,
        "shortcuts_ok": 0,
        "shortcuts_fail": 0,
        "details": [],
    }

    for app in target.get("apps", []):
        name = app.get("name", "")
        p_ok, p_reason = restore_program_path_for_app(app)
        if p_ok:
            result["program_path_ok"] += 1
        else:
            result["program_path_fail"] += 1

        r_ok, r_fail = restore_registry_for_app(app)
        result["registry_ok"] += r_ok
        result["registry_fail"] += r_fail

        e_ok, e_fail = restore_environment_for_app(app)
        result["environment_ok"] += e_ok
        result["environment_fail"] += e_fail

        sv_ok, sv_fail = restore_services_for_app(app)
        result["services_ok"] += sv_ok
        result["services_fail"] += sv_fail

        t_ok, t_fail = restore_tasks_for_app(app)
        result["tasks_ok"] += t_ok
        result["tasks_fail"] += t_fail

        s_ok, s_fail = restore_shortcuts_for_app(app)
        result["shortcuts_ok"] += s_ok
        result["shortcuts_fail"] += s_fail

        result["details"].append({"name": name, "program_path": p_reason})

    target["status"] = "restored"
    target["restored_at"] = datetime.now().isoformat()
    target["restore_result"] = result
    update_batch(target)
    return target, "ok"


def migrate_selected_apps(
    selected_apps,
    target_root,
    auto_arch=False,
    target_root_x64=None,
    target_root_x86=None,
    preserve_relative_layout=True,
    destination_overrides=None,
):
    batch_id = now_batch_id()
    backup_base = os.path.abspath(os.path.join(PathConfig.BACKUP_ROOT, batch_id))
    ensure_dir(backup_base)

    manual_target_root = os.path.abspath(target_root)
    resolved_x64, resolved_x86 = get_program_files_roots()
    if target_root_x64:
        resolved_x64 = os.path.abspath(target_root_x64)
    if target_root_x86:
        resolved_x86 = os.path.abspath(target_root_x86)

    override_map = destination_overrides or {}

    batch = {
        "id": batch_id,
        "created_at": datetime.now().isoformat(),
        "target_root": manual_target_root,
        "auto_arch": bool(auto_arch),
        "target_root_x64": resolved_x64,
        "target_root_x86": resolved_x86,
        "status": "applied",
        "backup_base": backup_base,
        "apps": [],
    }

    logger.info(f"========== 开始应用迁移批次 ==========")
    logger.info(f"批次 ID: {batch_id}")
    logger.info(f"目标根目录: {manual_target_root}")
    logger.info(f"应用数量: {len(selected_apps)}")
    logger.info(f"备份目录: {backup_base}")

    for app in selected_apps:
        name = app["display_name"]
        src = os.path.abspath(app["install_dir"])
        app_target_root, detected_arch = choose_target_root_for_app(
            app,
            manual_target_root,
            auto_arch,
            target_root_x64=resolved_x64,
            target_root_x86=resolved_x86,
        )
        ensure_dir(app_target_root)
        dst_subpath = choose_destination_subpath(
            src, preserve_relative_layout=preserve_relative_layout
        )
        dst = os.path.abspath(os.path.join(app_target_root, dst_subpath))
        override_dst = override_map.get(path_norm(src))
        if override_dst:
            dst = os.path.abspath(override_dst)
            try:
                dst_subpath = os.path.relpath(dst, app_target_root)
            except ValueError:
                dst_subpath = dst

        app_record = {
            "name": name,
            "src": src,
            "dst": dst,
            "dst_subpath": dst_subpath,
            "target_root_used": app_target_root,
            "detected_arch": detected_arch,
            "hive": app["hive"],
            "reg_subkey": app["reg_subkey"],
            "view": app["view"],
            "copy": "pending",
            "registry": {},
            "environment": {},
            "services": {},
            "tasks": {},
            "shortcuts": {},
            "delete_old": {},
        }

        logger.info(f"\n---- 迁移应用: {name} ----")
        if auto_arch:
            logger.info(f"检测位数: {detected_arch}")
            logger.info(f"目标根目录: {app_target_root}")
        if preserve_relative_layout:
            logger.info(f"目标子路径策略: 保留相对层级 -> {dst_subpath}")
        else:
            logger.info(f"目标子路径策略: 扁平化 -> {dst_subpath}")
        logger.info(f"源路径: {src}")
        logger.info(f"目标路径: {dst}")

        if path_norm(src) == path_norm(dst):
            app_record["copy"] = "skipped_same_path"
            app_record["error"] = "源路径与目标路径相同"
            batch["apps"].append(app_record)
            logger.warning("跳过: 源路径与目标路径相同")
            continue

        try:
            copy_app_dir(src, dst)
            app_record["copy"] = "ok"
            logger.info(f"应用文件复制成功: {name}")
        except Exception as e:
            app_record["copy"] = "failed"
            app_record["error"] = f"复制失败: {e}"
            batch["apps"].append(app_record)
            logger.error(f"应用文件复制失败 {name}: {e}")
            continue

        reg_bak_dir = os.path.join(backup_base, "registry")
        reg_bak_file = os.path.join(reg_bak_dir, f"{sanitize_name(name)}.reg")
        exported = export_reg_key(app["hive"], app["reg_subkey"], reg_bak_file)

        reg_changed, reg_failed, reg_details = replace_in_registry_values(
            app["hive"], app["reg_subkey"], app["view"], src, dst
        )
        app_record["registry"] = {
            "backup_file": reg_bak_file if exported else "",
            "backup_exported": exported,
            "changed": reg_changed,
            "failed": reg_failed,
            "details": reg_details,
        }

        env_baks = backup_environment_keys(backup_base, name)
        env_changed, env_failed, env_details = replace_in_environment_values(src, dst)
        app_record["environment"] = {
            "backup_files": env_baks,
            "changed": env_changed,
            "failed": env_failed,
            "details": env_details,
        }

        service_matches = scan_service_path_matches(src, dst, max_depth=2)
        service_baks = export_registry_backups_for_matches(service_matches, backup_base)
        svc_changed, svc_failed = apply_registry_drive_matches(service_matches)
        app_record["services"] = {
            "matched": len(service_matches),
            "backup_files": service_baks,
            "changed": svc_changed,
            "failed": svc_failed,
            "changes": service_matches,
        }

        task_matches = scan_taskcache_path_matches(src, dst, max_depth=3)
        task_baks = export_registry_backups_for_matches(task_matches, backup_base)
        task_changed, task_failed = apply_registry_drive_matches(task_matches)
        app_record["tasks"] = {
            "matched": len(task_matches),
            "backup_files": task_baks,
            "changed": task_changed,
            "failed": task_failed,
            "changes": task_matches,
        }

        shortcut_result = update_shortcuts_for_path(src, dst, backup_base)
        app_record["shortcuts"] = shortcut_result

        deleted, reason = try_remove_old_dir(src)
        app_record["delete_old"] = {"success": deleted, "reason": reason}
        if not deleted:
            add_pending_cleanup(src, reason, batch_id, name)
            logger.warning("旧路径删除失败，已写入待清理列表。")

        batch["apps"].append(app_record)
        logger.info(
            f"应用迁移完成: 复制={app_record['copy']} | "
            f"注册表变更={reg_changed} | 环境变量变更={env_changed} | "
            f"服务项变更={svc_changed} | 计划任务变更={task_changed} | "
            f"快捷方式变更={shortcut_result['changed']} | "
            f"删除旧路径={'成功' if deleted else '失败'}"
        )

    append_batch(batch)
    print("\n========================================")
    print("迁移批次完成")
    print(f"批次ID: {batch_id}")
    print(f"备份目录: {backup_base}")
    print(f"清单文件: {PathConfig.MANIFEST_FILE}")
    print(f"待清理文件: {PathConfig.PENDING_CLEAN_FILE}")
    return batch


def normalize_drive(text):
    text = str(text or "").strip().upper()
    if len(text) == 1 and text.isalpha():
        text = f"{text}:"
    if len(text) == 2 and text[1] == ":" and text[0].isalpha():
        return text
    return ""


def build_drive_shortcut_regex(old_drive):
    return re.compile(rf"(?i){re.escape(old_drive)}(?=[\\/])")


def build_drive_registry_regex(old_drive):
    return re.compile(rf"(?i)(?<![\\w:]){re.escape(old_drive)}(?!/)(?=[\\/\"'\\s;]|$)")


def load_drive_fix_manifest():
    data = load_json(PathConfig.DRIVE_FIX_MANIFEST_FILE, {"batches": []})
    if not isinstance(data, dict) or not isinstance(data.get("batches"), list):
        return {"batches": []}
    return data


def save_drive_fix_manifest(data):
    save_json(PathConfig.DRIVE_FIX_MANIFEST_FILE, data)


def append_drive_fix_batch(batch):
    manifest = load_drive_fix_manifest()
    manifest["batches"].append(batch)
    save_drive_fix_manifest(manifest)


def update_drive_fix_batch(updated_batch):
    manifest = load_drive_fix_manifest()
    for i, item in enumerate(manifest["batches"]):
        if item.get("id") == updated_batch.get("id"):
            manifest["batches"][i] = updated_batch
            save_drive_fix_manifest(manifest)
            return True
    return False


def list_drive_fix_batches(applied_only=False):
    batches = load_drive_fix_manifest().get("batches", [])
    if applied_only:
        batches = [b for b in batches if b.get("status") == "applied"]
    return batches


def scan_service_path_matches(old_path, new_path, max_depth=2):
    services_root = r"SYSTEM\CurrentControlSet\Services"
    pattern = re.compile(re.escape(old_path), re.IGNORECASE)
    old_norm = path_norm(old_path)
    matches = []

    def _scan(path, depth):
        try:
            hkey = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                path,
                0,
                winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
            )
        except OSError:
            return

        try:
            i = 0
            while True:
                try:
                    vname, vdata, vtype = winreg.EnumValue(hkey, i)
                    i += 1
                except OSError:
                    break

                if vtype not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) or not isinstance(vdata, str):
                    continue
                if old_norm not in path_norm(vdata):
                    continue

                new_val = pattern.sub(lambda _m: new_path, vdata)
                if new_val == vdata:
                    continue

                matches.append(
                    {
                        "root": "HKLM",
                        "path": path,
                        "view": "64",
                        "value_name": vname,
                        "old_value": vdata,
                        "new_value": new_val,
                        "type": vtype,
                    }
                )

            if depth < max_depth:
                j = 0
                while True:
                    try:
                        sub = winreg.EnumKey(hkey, j)
                        j += 1
                    except OSError:
                        break
                    _scan(f"{path}\\{sub}", depth + 1)
        finally:
            winreg.CloseKey(hkey)

    _scan(services_root, 0)
    return matches


def scan_taskcache_path_matches(old_path, new_path, max_depth=3):
    roots = [
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tasks",
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Schedule\TaskCache\Tree",
    ]
    pattern = re.compile(re.escape(old_path), re.IGNORECASE)
    old_norm = path_norm(old_path)
    matches = []

    def _scan(path, depth):
        try:
            hkey = winreg.OpenKey(
                winreg.HKEY_LOCAL_MACHINE,
                path,
                0,
                winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
            )
        except OSError:
            return

        try:
            i = 0
            while True:
                try:
                    vname, vdata, vtype = winreg.EnumValue(hkey, i)
                    i += 1
                except OSError:
                    break

                if vtype not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) or not isinstance(vdata, str):
                    continue
                if old_norm not in path_norm(vdata):
                    continue

                new_val = pattern.sub(lambda _m: new_path, vdata)
                if new_val == vdata:
                    continue

                matches.append(
                    {
                        "root": "HKLM",
                        "path": path,
                        "view": "64",
                        "value_name": vname,
                        "old_value": vdata,
                        "new_value": new_val,
                        "type": vtype,
                    }
                )

            if depth < max_depth:
                j = 0
                while True:
                    try:
                        sub = winreg.EnumKey(hkey, j)
                        j += 1
                    except OSError:
                        break
                    _scan(f"{path}\\{sub}", depth + 1)
        finally:
            winreg.CloseKey(hkey)

    for root in roots:
        _scan(root, 0)
    return matches


def delete_drive_fix_record(record_id, delete_backup=False):
    record_id = str(record_id or "").strip()
    if not record_id:
        return False, "invalid_id"

    m = load_drive_fix_manifest()
    batches = m.get("batches", [])
    idx = -1
    target = None
    for i, b in enumerate(batches):
        if b.get("id") == record_id:
            idx = i
            target = b
            break

    if idx < 0 or target is None:
        return False, "not_found"

    del batches[idx]
    save_drive_fix_manifest(m)

    if delete_backup:
        backup_base = str(target.get("backup_base", "") or "")
        if backup_base and os.path.exists(backup_base):
            try:
                shutil.rmtree(backup_base)
            except Exception:
                return True, "record_deleted_backup_failed"

    return True, "ok"


def scan_registry_drive_matches(old_drive, new_drive, max_depth=3, search_roots=None):
    regex = build_drive_registry_regex(old_drive)
    roots = [(winreg.HKEY_LOCAL_MACHINE, "HKLM"), (winreg.HKEY_CURRENT_USER, "HKCU")]
    matches = []
    roots_to_scan = search_roots or DRIVE_FIX_SEARCH_ROOTS

    def _scan_subkey(root_key, path, root_name, view_flag, depth):
        try:
            hkey = winreg.OpenKey(root_key, path, 0, winreg.KEY_READ | view_flag)
        except (FileNotFoundError, PermissionError, OSError):
            return

        try:
            i = 0
            while True:
                try:
                    vname, vdata, vtype = winreg.EnumValue(hkey, i)
                    if vtype in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) and isinstance(vdata, str):
                        if regex.search(vdata):
                            new_val = regex.sub(new_drive, vdata)
                            if any(
                                x in new_val.lower()
                                for x in ("ms-resourced://", "httpd://", "filed://")
                            ):
                                i += 1
                                continue
                            matches.append(
                                {
                                    "root": root_name,
                                    "path": path,
                                    "view": ("64" if view_flag == winreg.KEY_WOW64_64KEY else "32"),
                                    "value_name": vname,
                                    "old_value": vdata,
                                    "new_value": new_val,
                                    "type": vtype,
                                }
                            )
                    i += 1
                except OSError:
                    break

            if depth < max_depth:
                j = 0
                while True:
                    try:
                        sub_name = winreg.EnumKey(hkey, j)
                        _scan_subkey(
                            root_key,
                            f"{path}\\{sub_name}",
                            root_name,
                            view_flag,
                            depth + 1,
                        )
                        j += 1
                    except OSError:
                        break
        finally:
            winreg.CloseKey(hkey)

    for root_key, root_name in roots:
        for key in roots_to_scan:
            for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
                _scan_subkey(root_key, key, root_name, view, depth=0)

    return matches


def export_registry_backups_for_matches(matches, backup_base):
    if not matches:
        return []

    reg_dir = os.path.join(backup_base, "registry")
    ensure_dir(reg_dir)

    key_set = set()
    for m in matches:
        hive = "HKEY_LOCAL_MACHINE" if m["root"] == "HKLM" else "HKEY_CURRENT_USER"
        key_set.add(f"{hive}\\{m['path']}")

    backups = []
    for key_path in sorted(key_set):
        out_file = os.path.join(reg_dir, f"{sanitize_name(key_path)}.reg")
        try:
            subprocess.run(
                ["reg", "export", key_path, out_file, "/y"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
            )
            backups.append(out_file)
        except subprocess.CalledProcessError:
            pass
    return backups


def apply_registry_drive_matches(matches):
    success = 0
    failed = 0

    for m in matches:
        try:
            root = winreg.HKEY_LOCAL_MACHINE if m["root"] == "HKLM" else winreg.HKEY_CURRENT_USER
            access = winreg.KEY_SET_VALUE | (
                winreg.KEY_WOW64_64KEY if m["view"] == "64" else winreg.KEY_WOW64_32KEY
            )
            hkey = winreg.OpenKey(root, m["path"], 0, access)
            winreg.SetValueEx(hkey, m["value_name"], 0, m["type"], m["new_value"])
            winreg.CloseKey(hkey)
            success += 1
        except Exception:
            failed += 1
    return success, failed


def replace_drive_in_environment_values(old_drive, new_drive):
    regex = build_drive_shortcut_regex(old_drive)
    changed = 0
    failed = 0
    details = []

    for hive_name, subkey in ENVIRONMENT_ROOTS:
        root = winreg.HKEY_LOCAL_MACHINE if hive_name == "HKLM" else winreg.HKEY_CURRENT_USER

        try:
            key_r = winreg.OpenKey(root, subkey, 0, winreg.KEY_READ)
            key_w = winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE)
        except OSError:
            continue

        i = 0
        while True:
            try:
                name, data, typ = winreg.EnumValue(key_r, i)
                i += 1
            except OSError:
                break

            if typ not in (winreg.REG_SZ, winreg.REG_EXPAND_SZ) or not isinstance(data, str):
                continue
            if not regex.search(data):
                continue

            new_data = regex.sub(new_drive, data)
            if new_data == data:
                continue

            try:
                winreg.SetValueEx(key_w, name, 0, typ, new_data)
                changed += 1
                details.append(
                    {
                        "hive": hive_name,
                        "subkey": subkey,
                        "name": name,
                        "old": data,
                        "new": new_data,
                    }
                )
            except OSError:
                failed += 1

        winreg.CloseKey(key_r)
        winreg.CloseKey(key_w)

    return changed, failed, details


def fix_shortcuts_for_drive(old_drive, new_drive, backup_base, shortcut_roots=None):
    regex = build_drive_shortcut_regex(old_drive)
    shell = win32com.client.Dispatch("WScript.Shell")

    roots = shortcut_roots or candidate_shortcut_roots()
    scanned = 0
    changed = 0
    failed = 0
    changes = []

    for root in roots:
        if not root or not os.path.isdir(root):
            continue
        print(f"[快捷方式] 扫描 {root}")
        for dirpath, _, filenames in os.walk(root, onerror=lambda e: None):
            for name in filenames:
                if not name.lower().endswith(".lnk"):
                    continue

                scanned += 1
                lnk_path = os.path.join(dirpath, name)
                try:
                    shortcut = shell.CreateShortCut(lnk_path)
                except Exception:
                    failed += 1
                    continue

                modified = False
                before_after = {}
                for attr in ("TargetPath", "WorkingDirectory", "Arguments"):
                    val = getattr(shortcut, attr, "")
                    if isinstance(val, str) and regex.search(val):
                        new_val = regex.sub(new_drive, val)
                        before_after[attr] = {"old": val, "new": new_val}
                        setattr(shortcut, attr, new_val)
                        modified = True

                if not modified:
                    continue

                try:
                    backup_path = backup_shortcut_file(lnk_path, backup_base)
                    shortcut.Save()
                    changed += 1
                    changes.append(
                        {
                            "path": lnk_path,
                            "backup": backup_path,
                            "diff": before_after,
                        }
                    )
                except Exception:
                    failed += 1

    del shell
    return {
        "roots": roots,
        "scanned": scanned,
        "changed": changed,
        "failed": failed,
        "changes": changes,
    }


def restore_drive_fix_shortcuts(batch):
    ok = 0
    fail = 0
    for item in batch.get("shortcuts", {}).get("changes", []):
        target = item.get("path")
        backup = item.get("backup")
        try:
            if not target or not backup or not os.path.exists(backup):
                fail += 1
                continue
            ensure_dir(os.path.dirname(target))
            shutil.copy2(backup, target)
            ok += 1
        except Exception:
            fail += 1
    return ok, fail


def restore_drive_fix_registry(batch):
    ok = 0
    fail = 0
    for reg_file in batch.get("registry", {}).get("backup_files", []):
        if import_reg_file(reg_file):
            ok += 1
        else:
            fail += 1
    return ok, fail


def restore_drive_fix_environment(batch):
    ok = 0
    fail = 0
    for reg_file in batch.get("environment", {}).get("backup_files", []):
        if import_reg_file(reg_file):
            ok += 1
        else:
            fail += 1
    return ok, fail


def run_drive_letter_fix(
    old_drive,
    new_drive,
    include_registry=True,
    include_environment=True,
    include_shortcuts=True,
    shortcut_roots=None,
    registry_depth=3,
    search_roots=None,
):
    old_drive = normalize_drive(old_drive)
    new_drive = normalize_drive(new_drive)
    if not old_drive or not new_drive:
        raise ValueError("盘符格式无效")
    if old_drive == new_drive:
        raise ValueError("原盘符和新盘符不能相同")

    batch_id = now_batch_id()
    backup_base = os.path.abspath(os.path.join(PathConfig.DRIVE_FIX_BACKUP_ROOT, batch_id))
    ensure_dir(backup_base)

    batch = {
        "id": batch_id,
        "created_at": datetime.now().isoformat(),
        "old_drive": old_drive,
        "new_drive": new_drive,
        "status": "applied",
        "backup_base": backup_base,
        "options": {
            "include_registry": bool(include_registry),
            "include_environment": bool(include_environment),
            "include_shortcuts": bool(include_shortcuts),
            "registry_depth": int(registry_depth),
        },
        "shortcuts": {},
        "registry": {},
        "environment": {},
        "restored_at": None,
    }

    print("\n=== 开始执行盘符修复 ===")
    print(f"原盘符: {old_drive} -> 新盘符: {new_drive}")

    if include_shortcuts:
        print("\n[1/3] 快捷方式扫描与修复...")
        sc_result = fix_shortcuts_for_drive(
            old_drive, new_drive, backup_base, shortcut_roots=shortcut_roots
        )
        batch["shortcuts"] = sc_result
    else:
        batch["shortcuts"] = {"skipped": True, "changes": []}

    if include_registry:
        print("\n[2/3] 注册表扫描、备份与修复...")
        reg_matches = scan_registry_drive_matches(
            old_drive,
            new_drive,
            max_depth=registry_depth,
            search_roots=search_roots,
        )
        reg_backup_files = export_registry_backups_for_matches(reg_matches, backup_base)
        reg_ok, reg_fail = apply_registry_drive_matches(reg_matches)
        batch["registry"] = {
            "matched": len(reg_matches),
            "applied_success": reg_ok,
            "applied_failed": reg_fail,
            "backup_files": reg_backup_files,
            "changes": reg_matches,
        }
    else:
        batch["registry"] = {"skipped": True, "changes": []}

    if include_environment:
        print("\n[3/3] 环境变量扫描、备份与修复...")
        env_backups = backup_environment_keys(backup_base, f"drive_fix_{old_drive}_to_{new_drive}")
        env_changed, env_failed, env_details = replace_drive_in_environment_values(
            old_drive, new_drive
        )
        batch["environment"] = {
            "backup_files": env_backups,
            "changed": env_changed,
            "failed": env_failed,
            "details": env_details,
        }
    else:
        batch["environment"] = {"skipped": True, "details": []}

    append_drive_fix_batch(batch)

    print("\n=== 盘符修复完成 ===")
    print(f"批次ID: {batch_id}")
    print(f"备份目录: {backup_base}")
    return batch


def restore_drive_fix_batch(batch_id=None):
    batches = list_drive_fix_batches(applied_only=True)
    if not batches:
        return None, "no_applied_batch"

    target = None
    if batch_id:
        for b in batches:
            if b.get("id") == batch_id:
                target = b
                break
        if target is None:
            return None, "batch_not_found"
    else:
        target = batches[-1]

    s_ok, s_fail = restore_drive_fix_shortcuts(target)
    r_ok, r_fail = restore_drive_fix_registry(target)
    e_ok, e_fail = restore_drive_fix_environment(target)

    target["status"] = "restored"
    target["restored_at"] = datetime.now().isoformat()
    target["restore_result"] = {
        "shortcut_success": s_ok,
        "shortcut_failed": s_fail,
        "registry_success": r_ok,
        "registry_failed": r_fail,
        "environment_success": e_ok,
        "environment_failed": e_fail,
    }
    update_drive_fix_batch(target)
    return target, "ok"
