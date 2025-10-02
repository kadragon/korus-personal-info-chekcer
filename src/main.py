"""
데이터 검사 프로세스를 실행하는 메인 스크립트입니다.

이 스크립트는 `checkers` 패키지 내의 모든 검사 모듈을 동적으로 찾아 실행합니다.
각 검사기는 `*_checker.py` 형식의 파일명을 가져야 하며, 내부에 `*_checker`라는
이름의 메인 함수를 포함해야 합니다.

디렉토리와 같은 환경 변수 등 필요한 구성을 설정한 후,
각 검사기의 메인 함수를 호출합니다.
이 스크립트는 애플리케이션의 기본 진입점으로 실행되도록 설계되었습니다.
"""

import importlib
import os
import pkgutil
from types import ModuleType

from dotenv import load_dotenv

import checkers
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


def discover_and_run_checkers(
    download_dir: str, reports_save_dir: str, prev_month_str: str
) -> int:
    """
    `checkers` 패키지 내의 모든 검사기 모듈을 동적으로 찾아 실행합니다.
    """
    total_count = 0
    for module_info in pkgutil.iter_modules(checkers.__path__):
        module_name = module_info.name
        if not module_name.endswith("_checker"):
            continue

        try:
            module: ModuleType = importlib.import_module(f"checkers.{module_name}")
            checker_func_name = module_name
            checker_func = getattr(module, checker_func_name, None)

            if callable(checker_func):
                count = checker_func(download_dir, reports_save_dir, prev_month_str)
                total_count += count
            else:
                print_error(
                    f"'{module_name}' 모듈에서 '{checker_func_name}' 함수를 "
                    f"찾을 수 없거나 실행할 수 없습니다."
                )

        except Exception as e:
            print_error(f"'{module_name}' 검사 중 예상치 못한 오류 발생: {e}")

    return total_count


def main():
    """
    데이터 검사 프로세스를 실행하는 메인 함수입니다.

    분석 대상 월(지난달)을 결정하고, 필요한 저장 디렉토리 구조를 생성한 후,
    `checkers` 패키지 내의 모든 검사기들을 동적으로 찾아 실행합니다.
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
        total_count = discover_and_run_checkers(
            download_dir, reports_save_dir, prev_month_str
        )

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
