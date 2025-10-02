"""
이 모듈은 개인 정보 접근 로그를 확인하는 역할을 합니다.
사용자의 개인 데이터 접근 기록이 포함된 Excel 파일을 처리하며,
여러 파일이 발견되면 병합한 후 다음 여러 검사를 수행합니다:
- '인사마스터' (HR 마스터) 프로그램 접근 (본인 접근 제외).
- 사용자의 대량 데이터 조회.
- 사용자의 대량 데이터 저장.

각 검사에 대한 필터링된 결과는 별도의 Excel 파일에 저장되며, 대량 접근 보고서에는
임계값을 초과하는 사용자별로 시트가 생성될 수 있습니다.
"""

import os
from datetime import datetime

import pandas as pd

import config as cfg
from display import print_checker_header, print_result
from utils import find_and_prepare_excel_file, run_and_save_check


def personal_file_checker(download_dir: str, save_dir: str, prev_month: str) -> int:
    """
    개인 정보 접근 로그를 확인하는 메인 함수입니다.

    `download_dir`에서 모든 관련 Excel 파일을 찾아 단일 DataFrame으로 병합한 후,
    다음과 같은 다양한 필터를 적용합니다:
    1. '인사마스터' (HR 마스터) 프로그램 접근 (사용자가 자신의 기록에
       접근하는 경우 제외).
    2. 비정상적으로 많은 수의 기록을 조회한 사용자 (`VIEW_THRESHOLD` 초과).
    3. 비정상적으로 많은 수의 기록을 저장한 사용자 (`SAVE_THRESHOLD` 초과).

    각 검사 결과는 별도의 Excel 파일에 저장됩니다. 대량 접근에 대한 보고서는
    임계값을 초과한 특정 사용자의 모든 기록을 포함하는 시트가 있는
    다중 시트 Excel 파일입니다.

    매개변수:
        download_dir (str): 원본 개인 정보 접근 로그 Excel 파일이 있는 디렉토리입니다.
        save_dir (str): 생성된 보고서 Excel 파일이 저장될 디렉토리입니다.
        prev_month (str): 'YYYYMM' 형식의 이전 달로, 출력 파일 이름 지정에 사용됩니다.

    반환 값:
        int: 처리된 원본 데이터의 행 수입니다. 파일을 찾을 수 없으면 0입니다.

    예외:
        ValueError: 유효한 Excel 파일을 병합할 수 없거나 필수 열이 누락된 경우.
    """
    print_checker_header(cfg.PERSONAL_INFO_REPORT_BASE)

    file_prefix = (
        f"{cfg.PERSONAL_INFO_ACCESS_LOG_PREFIX}{datetime.today().strftime('%Y%m')}"
    )

    df, _ = find_and_prepare_excel_file(
        download_dir,
        file_prefix,
        save_dir,
        cfg.PERSONAL_INFO_REPORT_BASE,
        prev_month,
    )

    if df is None:
        return 0

    df_to_analyze = df

    # 필터 1: '인사마스터'(HR 마스터) 접근, 본인 접근 제외.
    base_report_name = cfg.MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE.split(".")[
        0
    ].format(prev_month)
    save_path_master = os.path.join(
        save_dir,
        f"{base_report_name}({cfg.PERSONAL_INFO_ACCESS_MASTER_SUFFIX}).xlsx",
    )
    run_and_save_check(
        df=df_to_analyze,
        check_func=_filter_by_job_master_exclude_detail_id,
        save_path=save_path_master,
        result_description="인사마스터 타인 조회",
    )

    # 필터 2: 대량의 기록을 조회하는 사용자.
    save_path_high_views = os.path.join(
        save_dir,
        f"{base_report_name}({cfg.PERSONAL_INFO_ACCESS_HIGH_VOLUME_VIEWS_SUFFIX}).xlsx",
    )
    _extract_and_save_by_job(
        df_to_analyze,
        save_path_high_views,
        job="조회",
        threshold=cfg.VIEW_THRESHOLD,
        job_column_name=cfg.COL_JOB_PERFORMANCE,
    )

    # 필터 3: 대량의 기록을 저장하는 사용자.
    save_path_high_saves = os.path.join(
        save_dir,
        f"{base_report_name}({cfg.PERSONAL_INFO_ACCESS_HIGH_VOLUME_SAVES_SUFFIX}).xlsx",
    )
    _extract_and_save_by_job(
        df_to_analyze,
        save_path_high_saves,
        job="저장",
        threshold=cfg.SAVE_THRESHOLD,
        job_column_name=cfg.COL_JOB_PERFORMANCE,
    )

    return len(df)


def _filter_by_job_master_exclude_detail_id(df: pd.DataFrame) -> pd.DataFrame:
    """
    '인사마스터' (HR 마스터) 프로그램 접근 기록을 필터링하며, 사용자의 ID가
    '상세내용' 필드에 나타나는 경우(즉, 본인 접근)는 제외합니다.

    매개변수:
        df (pd.DataFrame): 개인 정보 접근 로그를 포함하는 DataFrame입니다.
                           예상 열: `COL_PROGRAM_NAME`,
                           `COL_EMPLOYEE_ID`,
                           `COL_ACCESS_TIME`, `COL_DETAIL_CONTENT`.

    반환 값:
        pd.DataFrame: 필터링된 기록을 포함하며 직원 ID와 접근 시간으로 정렬된
                      DataFrame입니다. 해당 기록이 없으면 빈 DataFrame을 반환합니다.

    예외:
        ValueError: 필터링에 필수적인 열이 누락된 경우.
    """
    employee_id_col_to_use = cfg.COL_EMPLOYEE_ID

    required_cols = [
        cfg.COL_PROGRAM_NAME,
        employee_id_col_to_use,
        cfg.COL_ACCESS_TIME,
        cfg.COL_DETAIL_CONTENT,
    ]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(
                f"'{col}' 컬럼을 찾을 수 없어 HR 마스터 필터링을 할 수 없습니다."
            )

    # '인사마스터' 프로그램 접근 기록 중, 본인 조회가 아닌 경우만 필터링합니다.
    # 조건 1: '프로그램명'이 '인사마스터'인 기록만 선택합니다.
    hr_master_df = df[df[cfg.COL_PROGRAM_NAME] == "인사마스터"].copy()

    # 조건 2: '상세내용'에 자신의 '직원ID'가 포함되어 있지 않은 기록만 남깁니다.
    #         (즉, 타인 조회 기록만 필터링)
    # 각 행을 순회하며 '직원ID'와 '상세내용'을 비교해야 하므로 apply 함수를 사용합니다.
    is_not_self_access = [
        str(emp_id) not in str(detail)
        for emp_id, detail in zip(
            hr_master_df[employee_id_col_to_use],
            hr_master_df[cfg.COL_DETAIL_CONTENT],
            strict=True,
        )
    ]
    filtered_df = hr_master_df[is_not_self_access]

    # 결과를 정렬합니다.
    return filtered_df.sort_values([employee_id_col_to_use, cfg.COL_ACCESS_TIME])


def _extract_and_save_by_job(
    df: pd.DataFrame, save_path: str, job: str, threshold: int, job_column_name: str
):
    """
    특정 `job`(예: '조회', '저장')을 `threshold` 횟수 이상 수행한 사용자를 식별합니다.
    이러한 각 사용자에 대해 해당 작업과 일치하는 기록뿐만 아니라 모든 기록을
    지정된 Excel 파일의 별도 시트에 저장합니다.

    매개변수:
        df (pd.DataFrame): 모든 개인 정보 접근 로그를 포함하는 DataFrame입니다.
        save_path (str): 결과가 저장될 Excel 파일의 전체 경로입니다.
        job (str): 계산할 특정 작업 유형입니다 (예: '조회', '저장').
        threshold (int): 사용자를 표시하기 위해 `job`을 수행해야 하는 최소 횟수입니다.
        job_column_name (str): `df`에서 작업 유형을 포함하는 열의 이름입니다
                               (예: `COL_JOB_PERFORMANCE`).

    예외:
        ValueError: 필수 열(`COL_EMPLOYEE_ID`, `COL_EMPLOYEE_NAME`,
                      `job_column_name`)이 누락된 경우.
    """
    employee_id_col_to_use = cfg.COL_EMPLOYEE_ID

    required_cols_check = [
        employee_id_col_to_use,
        cfg.COL_EMPLOYEE_NAME,
        job_column_name,
    ]
    for col in required_cols_check:
        if col not in df.columns:
            raise ValueError(f"'{col}' 컬럼을 찾을 수 없어 작업을 추출할 수 없습니다.")

    job_specific_df = df[df[job_column_name] == job]
    job_counts_per_user = job_specific_df.groupby(employee_id_col_to_use).size()
    target_user_ids = job_counts_per_user[
        job_counts_per_user >= threshold
    ].index.tolist()

    description = f"{job} {threshold}건 이상"

    if not target_user_ids:
        print_result(is_detected=False, description=description)
        return

    with pd.ExcelWriter(save_path) as writer:
        for employee_id in target_user_ids:
            user_all_records_df = df[df[employee_id_col_to_use] == employee_id]
            user_name = ""
            if (
                not user_all_records_df.empty
                and cfg.COL_EMPLOYEE_NAME in user_all_records_df.columns
            ):
                user_name = user_all_records_df[cfg.COL_EMPLOYEE_NAME].iloc[0]

            sheet_name = f"{employee_id}_{user_name}"
            if len(sheet_name) > cfg.SHEET_NAME_MAX_CHARS:
                sheet_name = sheet_name[: cfg.SHEET_NAME_MAX_CHARS]

            user_all_records_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print_result(
        is_detected=True,
        description=f"{description} ({len(target_user_ids)}명)",
        filename=os.path.basename(save_path),
    )
