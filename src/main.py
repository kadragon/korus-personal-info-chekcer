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

from checkers.download_reason_checker import sayu_checker
from checkers.login_checker import login_checker
from checkers.personal_file_checker import personal_file_checker
from display import (
    print_error,
    print_header,
    print_info,
    print_summary,
    print_zip_header,
)
from utils import get_prev_month_yyyymm, make_save_dir, zip_files_by_prefix

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
    """
    if not base_save_dir:
        print_error(
            "SAVE_DIR 환경 변수가 설정되지 않았습니다. "
            ".env 파일이나 환경에서 설정해주세요."
        )
        return
    if not download_dir:
        print_error(
            "DOWNLOAD_DIR 환경 변수가 설정되지 않았습니다. "
            ".env 파일이나 환경에서 설정해주세요."
        )
        return

    prev_month_str = get_prev_month_yyyymm()
    reports_save_dir = make_save_dir(base_save_dir)

    print_header(f"개인정보 처리 현황 자동 점검 ({prev_month_str}월분)")
    print_info(f"원본 데이터 경로: {download_dir}")
    print_info(f"결과 저장 경로: {reports_save_dir}")

    if download_dir and reports_save_dir:
        total_count = 0

        try:
            count = sayu_checker(download_dir, reports_save_dir, prev_month_str)
            total_count += count
        except Exception as e:
            print_error(f"다운로드 사유 검사 중 예상치 못한 오류 발생: {e}")

        try:
            count = login_checker(download_dir, reports_save_dir, prev_month_str)
            total_count += count
        except Exception as e:
            print_error(f"로그인 검사 중 예상치 못한 오류 발생: {e}")

        try:
            count = personal_file_checker(
                download_dir, reports_save_dir, prev_month_str
            )
            total_count += count
        except Exception as e:
            print_error(f"개인 파일 접근 검사 중 예상치 못한 오류 발생: {e}")

        print_zip_header()
        try:
            zip_files_by_prefix(reports_save_dir, ["[붙임2", "[붙임3", "[붙임4"])
        except Exception as e:
            print_error(f"압축 작업 중 오류 발생: {e}")

        print_summary(reports_save_dir, total_count)

    else:
        print_error("다운로드 경로나 저장 경로가 올바르게 설정되지 않았습니다.")


if __name__ == "__main__":
    main()
