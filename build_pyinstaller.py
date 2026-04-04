#!/usr/bin/env python3
"""
Windows Path Migration Toolkit 的 PyInstaller 编译脚本。

默认策略（贴合当前项目）：
- 启用管理员权限（--uac-admin）
- 目录模式打包（默认不启用 --onefile）
- GUI 模式（--windowed）
- 将 src 目录加入 PyInstaller 分析路径（--paths）
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent
DEFAULT_ENTRY = "main.py"
DEFAULT_NAME = "WindowsPathMigrationToolkit"


def remove_path(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def collect_data_args(project_root: Path) -> list[str]:
    """收集需要随包分发的数据目录与文件。"""
    args: list[str] = []

    data_dir = project_root / "data"
    if data_dir.exists() and data_dir.is_dir():
        args.extend(["--add-data", f"{data_dir};data"])

    test_reg = project_root / "test.reg"
    if test_reg.exists() and test_reg.is_file():
        args.extend(["--add-data", f"{test_reg};."])

    return args


def build_command(args: argparse.Namespace) -> list[str]:
    entry_path = (PROJECT_ROOT / args.entry).resolve()
    src_path = (PROJECT_ROOT / "src").resolve()

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--windowed",
        "--uac-admin",
        "--name",
        args.name,
        "--paths",
        str(src_path),
        "--hidden-import",
        "win32com.client",
        "--hidden-import",
        "pythoncom",
        "--hidden-import",
        "pywintypes",
    ]

    if args.clean:
        cmd.append("--clean")

    if args.onefile:
        cmd.append("--onefile")

    if args.debug:
        cmd.append("--debug=all")

    cmd.extend(collect_data_args(PROJECT_ROOT))
    cmd.append(str(entry_path))
    return cmd


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="PyInstaller build script for Windows Path Migration Toolkit"
    )
    parser.add_argument("--name", default=DEFAULT_NAME, help="输出程序名称")
    parser.add_argument("--entry", default=DEFAULT_ENTRY, help="入口脚本路径")
    parser.add_argument(
        "--onefile",
        action="store_true",
        help="启用单文件打包（默认关闭，使用目录模式）",
    )
    parser.add_argument("--clean", action="store_true", help="启用 PyInstaller --clean")
    parser.add_argument("--debug", action="store_true", help="启用 PyInstaller 调试输出")
    parser.add_argument(
        "--purge-output",
        action="store_true",
        help="打包前删除 build 和 dist 目录",
    )
    return parser.parse_args()


def main() -> int:
    if sys.platform != "win32":
        print(f"[WARN] 当前系统为 {sys.platform}，该脚本主要面向 Windows。")

    args = parse_args()
    version = "0.4.0"
    entry_path = (PROJECT_ROOT / args.entry).resolve()
    if not entry_path.exists():
        print(f"[ERROR] 入口脚本不存在: {entry_path}")
        return 2

    if args.purge_output:
        remove_path(PROJECT_ROOT / "build")
        remove_path(PROJECT_ROOT / "dist")

    cmd = build_command(args)
    printable_cmd = " ".join(f'"{part}"' if " " in part else part for part in cmd)
    print("[INFO] 即将执行:")
    print(printable_cmd)

    try:
        completed = subprocess.run(cmd, cwd=PROJECT_ROOT, check=False)
    except FileNotFoundError as exc:
        print("[ERROR] 无法启动 PyInstaller:", exc)
        return 3

    if completed.returncode != 0:
        print(f"[ERROR] 打包失败，退出码: {completed.returncode}")
        return completed.returncode

    print("[INFO] 打包完成。")
    print(f"[INFO] 输出目录: {PROJECT_ROOT / 'dist'}")
    print(f"[INFO] spec 文件: {PROJECT_ROOT / (args.name + '.spec')}")

    # 创建压缩包
    dist_dir = PROJECT_ROOT / "dist"
    app_dir = dist_dir / args.name
    zip_name = f"{args.name}-v{version}.zip"
    zip_path = dist_dir / zip_name

    if app_dir.exists():
        print(f"[INFO] 创建压缩包: {zip_path}")
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for file_path in app_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(app_dir)
                    zipf.write(file_path, arcname)
        print(f"[INFO] 压缩包创建完成: {zip_path}")
    else:
        print(f"[WARN] 应用目录不存在: {app_dir}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
