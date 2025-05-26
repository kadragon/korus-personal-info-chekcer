"""
데이터 검사 프로세스를 실행하는 메인 스크립트입니다.

이 스크립트는 다양한 검사 모듈의 실행을 조정합니다:
- 다운로드 사유 검사기 (`sayu_checker`)
- 로그인 검사기 (`login_checker`)
- 개인 파일 접근 검사기 (`personal_file_checker`)

디렉토리와 같은 환경 변수 등 필요한 구성을 설정한 후,
각 검사기의 메인 함수를 호출합니다.
이 스크립트는 애플리케이션의 기본 진입점으로 실행되도록 설계되었습니다.
"""

import os
from dotenv import load_dotenv

from utils import get_prev_month_yyyymm, make_save_dir
from checkers.download_reason_checker import sayu_checker
from checkers.login_checker import login_checker
from checkers.personal_file_checker import personal_file_checker

# .env 파일이 있는 경우 환경 변수를 로드합니다.
# DOWNLOAD_DIR 및 SAVE_DIR을 하드코딩하지 않고 구성하는 데 유용합니다.
load_dotenv()
download_dir = os.getenv("DOWNLOAD_DIR")  # 원본 Excel 파일이 있는 디렉토리입니다.
base_save_dir = os.getenv("SAVE_DIR")  # 출력 보고서가 저장될 기본 디렉토리입니다.


def main():
    """
    데이터 검사 프로세스를 실행하는 메인 함수입니다.

    분석 대상 월(지난달)을 결정하고, 필요한 저장 디렉토리 구조를 생성한 후,
    선택된 검사 함수들을 실행합니다.

    현재 `personal_file_checker`가 활성화되어 있으며, `sayu_checker`와
    `login_checker`는 주석 처리되어 있습니다. 실행에 포함하려면 주석을 해제하십시오.
    """
    if not base_save_dir:
        print(
            "Error: SAVE_DIR environment variable is not set. Please configure it in your .env file or environment."
        )
        return
    if not download_dir:
        print(
            "Error: DOWNLOAD_DIR environment variable is not set. Please configure it in your .env file or environment."
        )
        return

    prev_month_str = get_prev_month_yyyymm()
    # 이전 달의 보고서를 위한 특정 하위 디렉토리를 생성합니다.
    # 예: base_save_dir가 /reports이면, save_dir는 /reports/YYYYMM이 됩니다.
    reports_save_dir = make_save_dir(base_save_dir)

    print(f"Starting data checks for month: {prev_month_str}")
    print(f"Source data directory: {download_dir}")
    print(f"Reports will be saved in: {reports_save_dir}")

    # 계속하기 전에 디렉토리가 유효한지 확인합니다.
    if download_dir and reports_save_dir:
        # 개인 정보 다운로드 사유 확인 섹션입니다.
        # 원본 주석: "# 개인정보 다운로드 사유 검사"
        print("\n### Running Download Reason Checker ###")
        try:
            sayu_checker(download_dir, reports_save_dir, prev_month_str)
            print("Download Reason Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Download Reason Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Download Reason Checker: {e}")

        # 로그인 IP 패턴 확인 섹션입니다.
        # 원본 주석: "# 로그인 IP 검사"
        print("\n### Running Login Checker ###")
        try:
            login_checker(download_dir, reports_save_dir, prev_month_str)
            print("Login Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Login Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Login Checker: {e}")

        # 개인 정보 접근 기록 확인 섹션입니다.
        # 원본 주석: "# 개인정보 조회 기록 점검"
        print("\n### Running Personal File Checker ###")
        try:
            personal_file_checker(download_dir, reports_save_dir, prev_month_str)
            print("Personal File Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Personal File Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Personal File Checker: {e}")

        print("\nAll checks finished.")
    else:
        # 이 경우는 초기에 base_save_dir 및 download_dir에 대한 확인에서 처리되어야 합니다.
        print("Error: Download directory or save directory is not properly configured.")


if __name__ == "__main__":
    # 스크립트가 직접 실행될 때만 main()이 호출되도록 합니다,
    # 모듈로 가져올 때는 호출되지 않습니다.
    main()
