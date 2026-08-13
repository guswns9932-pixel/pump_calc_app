# -*- coding: utf-8 -*-
"""
Windows용 실행 파일(.exe)을 만드는 빌드 스크립트.

PyInstaller는 실행하는 OS용 실행 파일만 만들 수 있으므로(크로스 컴파일 불가),
반드시 Windows PC에서 이 스크립트를 실행해야 .exe가 나온다.

사용법 (Windows, cmd 또는 PowerShell):
    pip install pyinstaller
    python build_exe.py

완료되면 dist\\PumpPriceCalculator.exe 가 생성된다. 이 exe와 data 폴더를 같은
폴더에 두고 실행하면 된다 (data 폴더는 exe 안에 포함하지 않는다 — 이력을
새로 등록/수정할 때마다 exe 옆의 data 폴더에 그대로 저장되어야 하기 때문에,
읽기 전용으로 번들에 굳어지는 --add-data 방식을 쓰지 않았다. pump_calc_app.py의
BASE_DIR 계산이 frozen(exe) 상태를 감지해 exe가 있는 폴더를 기준으로 삼는다).
"""
import subprocess
import sys
import shutil
import os

APP_NAME = "PumpPriceCalculator"
ENTRY_SCRIPT = "pump_calc_app.py"


def main():
    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller가 설치되어 있지 않습니다. 먼저 다음을 실행하세요:")
        print("    pip install pyinstaller")
        sys.exit(1)

    here = os.path.dirname(os.path.abspath(__file__))
    entry = os.path.join(here, ENTRY_SCRIPT)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",  # 콘솔 창 없이 실행 (Windows)
        "--name", APP_NAME,
        "--noconfirm",
        entry,
    ]
    print("실행:", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=here)

    dist_dir = os.path.join(here, "dist")
    print()
    print(f"완료: {dist_dir} 폴더에 {APP_NAME} 실행 파일이 생성되었습니다.")
    print("이 실행 파일을 data 폴더와 같은 위치에 두고 실행하세요")
    print("(기존 사용자라면 원래 쓰던 data 폴더를 그대로 옆에 복사해 오면 이력이 유지됩니다).")


if __name__ == "__main__":
    main()
