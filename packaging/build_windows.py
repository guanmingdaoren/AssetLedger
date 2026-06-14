from pathlib import Path
import shutil
import subprocess
import sys


PROJECT_ROOT = Path(__file__).resolve().parent.parent
DIST_DIR = PROJECT_ROOT / "dist" / "设备资产台账"


def main() -> int:
    try:
        subprocess.run(
            [
                sys.executable,
                "-m",
                "PyInstaller",
                "--noconfirm",
                "--clean",
                "AssetLedger.spec",
            ],
            cwd=PROJECT_ROOT,
            check=True,
        )
    except subprocess.CalledProcessError as exc:
        return exc.returncode

    (DIST_DIR / "data").mkdir(parents=True, exist_ok=True)
    shutil.copy2(PROJECT_ROOT / "README.md", DIST_DIR / "使用说明.md")
    print(f"\nBuild complete: {DIST_DIR / '设备资产台账.exe'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
