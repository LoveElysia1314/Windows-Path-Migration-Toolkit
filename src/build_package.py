import argparse
import shutil
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent
ENTRY_DEFAULT = "app_path_migration_gui.py"
APP_NAME_DEFAULT = "WindowsPathMigrationToolkit"

DATA_FILES = [
    "app_gui_state.json",
    "app_migration_manifest.json",
    "app_migration_pending_cleanup.json",
    "drive_fix_manifest.json",
]


def remove_path(path: Path) -> None:
    if not path.exists():
        return
    if path.is_dir():
        shutil.rmtree(path)
    else:
        path.unlink()


def collect_data_args() -> list[str]:
    args: list[str] = []
    for name in DATA_FILES:
        src = ROOT / name
        if src.exists():
            args.extend(["--add-data", f"{src};."])
    return args


def build_command(name: str, entry: str, profile: str, onefile: bool, clean: bool) -> list[str]:
    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--noconfirm",
        "--windowed",
        "--uac-admin",
        "--name",
        name,
        "--hidden-import",
        "win32com.client",
    ]

    if clean:
        cmd.append("--clean")

    if onefile:
        cmd.append("--onefile")

    if profile == "full":
        # Full profile is more robust but larger in size.
        cmd.extend(["--collect-all", "PySide6"])
    else:
        # Slim profile only collects commonly used Qt modules.
        cmd.extend(
            [
                "--collect-submodules",
                "PySide6.QtCore",
                "--collect-submodules",
                "PySide6.QtGui",
                "--collect-submodules",
                "PySide6.QtWidgets",
            ]
        )

    cmd.extend(collect_data_args())
    cmd.append(entry)
    return cmd


def main() -> int:
    parser = argparse.ArgumentParser(description="PyInstaller build helper for this project")
    parser.add_argument("--name", default=APP_NAME_DEFAULT, help="Output application name")
    parser.add_argument("--entry", default=ENTRY_DEFAULT, help="Entry script path")
    parser.add_argument(
        "--profile",
        choices=["full", "slim"],
        default="full",
        help="full: larger, safer; slim: smaller, may need extra hidden imports",
    )
    parser.add_argument("--onefile", action="store_true", help="Build a single executable")
    parser.add_argument("--clean", action="store_true", help="Enable PyInstaller --clean")
    parser.add_argument(
        "--purge-output",
        action="store_true",
        help="Delete existing build/dist folders before packaging",
    )

    args = parser.parse_args()

    if sys.platform != "win32":
        print(
            "[WARN] This script targets Windows packaging. Current platform:",
            sys.platform,
        )

    entry_path = (ROOT / args.entry).resolve()
    if not entry_path.exists():
        print(f"[ERROR] Entry script not found: {entry_path}")
        return 2

    if args.purge_output:
        remove_path(ROOT / "build")
        remove_path(ROOT / "dist")

    cmd = build_command(args.name, str(entry_path), args.profile, args.onefile, args.clean)

    print("[INFO] Running command:")
    print(" ".join(f'"{part}"' if " " in part else part for part in cmd))

    try:
        result = subprocess.run(cmd, cwd=ROOT, check=False)
    except FileNotFoundError as exc:
        print("[ERROR] Failed to start PyInstaller:", exc)
        return 3

    if result.returncode != 0:
        print(f"[ERROR] Build failed with exit code: {result.returncode}")
        return result.returncode

    print("[INFO] Build completed.")
    print(f"[INFO] Output folder: {ROOT / 'dist'}")
    print(f"[INFO] Spec file: {ROOT / (args.name + '.spec')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
