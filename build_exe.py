# -*- coding: utf-8 -*-
"""
Windows용 실행 파일(.exe)을 만드는 빌드 스크립트.

PyInstaller는 실행하는 OS용 실행 파일만 만들 수 있으므로(크로스 컴파일 불가),
반드시 Windows PC에서 이 스크립트를 실행해야 .exe가 나온다.

사용법 (Windows, cmd 또는 PowerShell):
    pip install pyinstaller

    # 이름을 직접 지정
    python build_exe.py --name 우리팀_판가계산기

    # 이름을 지정하지 않으면 실행 중 물어본다 (엔터만 누르면 기본값 사용)
    python build_exe.py

완료되면 dist\\<지정한 이름>.exe 가 생성된다. 이 exe와 data 폴더를 같은
폴더에 두고 실행하면 된다 (data 폴더는 exe 안에 포함하지 않는다 — 이력을
새로 등록/수정할 때마다 exe 옆의 data 폴더에 그대로 저장되어야 하기 때문에,
읽기 전용으로 번들에 굳어지는 --add-data 방식을 쓰지 않았다. pump_calc_app.py의
BASE_DIR 계산이 frozen(exe) 상태를 감지해 exe가 있는 폴더를 기준으로 삼는다).
"""
import argparse
import os
import subprocess
import sys

DEFAULT_APP_NAME = "PumpPriceCalculator"
ENTRY_SCRIPT = "pump_calc_app.py"

# Windows 파일명에 쓸 수 없는 문자 (실행 파일 이름이자 작업관리자에 뜨는 프로세스 이름이 된다)
_INVALID_NAME_CHARS = '\\/:*?"<>|'


def sanitize_app_name(raw_name):
    """입력받은 이름에서 .exe 확장자와 파일명으로 쓸 수 없는 문자를 정리한다."""
    name = (raw_name or "").strip()
    if name.lower().endswith(".exe"):
        name = name[:-4]
    name = "".join(ch for ch in name if ch not in _INVALID_NAME_CHARS).strip()
    return name or DEFAULT_APP_NAME


def resolve_app_name(cli_name):
    if cli_name:
        return sanitize_app_name(cli_name)
    try:
        entered = input(f"실행 파일(프로세스) 이름을 입력하세요 [기본값: {DEFAULT_APP_NAME}]: ")
    except EOFError:
        entered = ""
    return sanitize_app_name(entered)


def main():
    parser = argparse.ArgumentParser(description="pump_calc_app.py를 Windows용 .exe로 빌드합니다.")
    parser.add_argument(
        "--name", "-n", default=None,
        help=f"생성할 실행 파일(프로세스) 이름. 지정하지 않으면 실행 중 물어봅니다 (기본값: {DEFAULT_APP_NAME}).",
    )
    args = parser.parse_args()

    try:
        import PyInstaller  # noqa: F401
    except ImportError:
        print("PyInstaller가 설치되어 있지 않습니다. 먼저 다음을 실행하세요:")
        print("    pip install pyinstaller")
        sys.exit(1)

    app_name = resolve_app_name(args.name)

    here = os.path.dirname(os.path.abspath(__file__))
    entry = os.path.join(here, ENTRY_SCRIPT)

    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",  # 콘솔 창 없이 실행 (Windows)
        "--name", app_name,
        "--noconfirm",
        entry,
    ]
    print("실행:", " ".join(cmd))
    subprocess.run(cmd, check=True, cwd=here)

    dist_dir = os.path.join(here, "dist")
    print()
    print(f"완료: {dist_dir} 폴더에 {app_name}.exe 실행 파일이 생성되었습니다.")
    print("이 실행 파일을 data 폴더와 같은 위치에 두고 실행하세요")
    print("(기존 사용자라면 원래 쓰던 data 폴더를 그대로 옆에 복사해 오면 이력이 유지됩니다).")


if __name__ == "__main__":
    main()
