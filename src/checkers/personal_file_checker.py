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
from utils import save_excel_with_autofit, find_and_prepare_excel_file

# personal_file_checker.py 상수
PERSONAL_INFO_ACCESS_LOG_PREFIX = (
    "개인정보 접속기록 조회_"  # 개인 정보 접근 로그 파일의 접두사
)
PERSONAL_INFO_REPORT_BASE = "[붙임3] 개인정보 접속기록 조회"
MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE = "[붙임3] 개인정보 접속기록 조회_{}.xlsx"
PERSONAL_INFO_ACCESS_MASTER_SUFFIX = (
    "인사마스터"  # 보고서 접미사: '인사마스터'(HR 마스터) 프로그램 접근
)
PERSONAL_INFO_ACCESS_HIGH_VOLUME_VIEWS_SUFFIX = (
    "1000건이상조회"  # 보고서 접미사: VIEW_THRESHOLD보다 많은 기록을 조회한 사용자
)
PERSONAL_INFO_ACCESS_HIGH_VOLUME_SAVES_SUFFIX = (
    "100건이상저장"  # 보고서 접미사: SAVE_THRESHOLD보다 많은 기록을 저장한 사용자
)
COL_ACCESS_TIME = "접속일시"  # 접근 타임스탬프 (예: "YYYY-MM-DD HH:MM:SS")
COL_EMPLOYEE_ID = (
    "교번"  # 직원 ID. 주로 다운로드 사유 및 개인 파일 접근 로그에 사용됩니다.
)
COL_EMPLOYEE_ID_LOGIN = "신분번호"  # 직원 ID, 특히 로그인 기록 파일에서 발견됨.
COL_PROGRAM_NAME = "프로그램명"  # 사용자가 접근한 프로그램/시스템의 이름
COL_DETAIL_CONTENT = "상세내용"  # 접근한 활동 또는 내용에 대한 상세 설명
COL_JOB_PERFORMANCE = "수행업무"  # 사용자가 수행한 작업/업무 (예: '조회', '저장')
COL_EMPLOYEE_NAME = "성명"  # 직원 이름
VIEW_THRESHOLD = (
    1000  # personal_file_checker용: 이 수보다 많은 기록을 조회한 사용자를 표시합니다.
)
SAVE_THRESHOLD = (
    100  # personal_file_checker용: 이 수보다 많은 기록을 저장한 사용자를 표시합니다.
)
SHEET_NAME_MAX_CHARS = 31  # Excel 시트 이름에 허용되는 최대 문자 수입니다.
EXCEL_EXTENSIONS = (
    ".xlsx",
    ".xls",
)  # 입력 파일에 지원되는 Excel 파일 확장자 튜플입니다.


def personal_file_checker(download_dir: str, save_dir: str, prev_month: str):
    """
    개인 정보 접근 로그를 확인하는 메인 함수입니다.

    `download_dir`에서 모든 관련 Excel 파일을 찾아 단일 DataFrame으로 병합한 후,
    다음과 같은 다양한 필터를 적용합니다:
    1. '인사마스터' (HR 마스터) 프로그램 접근 (사용자가 자신의 기록에 접근하는 경우 제외).
    2. 비정상적으로 많은 수의 기록을 조회한 사용자 (`VIEW_THRESHOLD` 초과).
    3. 비정상적으로 많은 수의 기록을 저장한 사용자 (`SAVE_THRESHOLD` 초과).

    각 검사 결과는 별도의 Excel 파일에 저장됩니다. 대량 접근에 대한 보고서는
    임계값을 초과한 특정 사용자의 모든 기록을 포함하는 시트가 있는 다중 시트 Excel 파일입니다.

    매개변수:
        download_dir (str): 원본 개인 정보 접근 로그 Excel 파일이 있는 디렉토리입니다.
        save_dir (str): 생성된 보고서 Excel 파일이 저장될 디렉토리입니다.
        prev_month (str): 'YYYYMM' 형식의 이전 달로, 출력 파일 이름 지정에 사용됩니다.

    예외:
        EnvironmentError: `download_dir`가 환경 변수에 설정되지 않은 경우.
        FileNotFoundError: `download_dir`에서 관련 Excel 파일을 찾을 수 없는 경우.
        ValueError: 유효한 Excel 파일을 병합할 수 없거나 필수 열이 누락된 경우.
    """
    if not download_dir:
        # 이 확인은 원본 스크립트 구조를 기반으로 합니다.
        # download_dir이 항상 사용 가능하다면 직접 인수로 만드는 것을 고려하십시오.
        raise EnvironmentError("DOWNLOAD_DIR environment variable is not set.")

    PERSONAL_INFO_ACCESS_LOG_PREFIX_FILE_PREFIX = f"{PERSONAL_INFO_ACCESS_LOG_PREFIX}{datetime.today().strftime('%Y%m')}"

    df, _ = find_and_prepare_excel_file(
        download_dir,
        PERSONAL_INFO_ACCESS_LOG_PREFIX_FILE_PREFIX,
        save_dir,
        PERSONAL_INFO_REPORT_BASE,
        prev_month,
    )

    if df is None:
        raise FileNotFoundError(
            f"Download reason Excel file starting with '{PERSONAL_INFO_ACCESS_LOG_PREFIX_FILE_PREFIX}' not found in '{download_dir}'."
        )

    df_to_analyze = df

    # 필터 1: '인사마스터'(HR 마스터) 접근, 본인 접근 제외.
    # 원본 주석: "인사마스터에서 조회한 기록 (본인 제외)"
    filtered_master_access = _filter_by_job_master_exclude_detail_id(
        df_to_analyze)
    if not filtered_master_access.empty:
        # MERGED_...TEMPLATE을 기반으로 파일 이름을 구성한 다음 특정 접미사를 추가합니다.
        base_report_name_for_master = (
            MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE.split(".")[0].format(
                prev_month
            )
        )
        save_path_master = os.path.join(
            save_dir,
            f"{base_report_name_for_master}({PERSONAL_INFO_ACCESS_MASTER_SUFFIX}).xlsx",
        )
        save_excel_with_autofit(filtered_master_access, save_path_master)
        print(
            f"HR Master access (excluding self) results saved to: {save_path_master}")
    else:
        print("No records found for HR Master access (excluding self) check.")

    # 필터 2: 대량의 기록을 조회하는 사용자.
    # 원본 주석: "조회 VIEW_THRESHOLD 이상 교번별 전체 기록 시트별 저장"
    base_report_name_for_views = MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE.split(
        "."
    )[0].format(prev_month)
    save_path_high_views = os.path.join(
        save_dir,
        f"{base_report_name_for_views}({PERSONAL_INFO_ACCESS_HIGH_VOLUME_VIEWS_SUFFIX}).xlsx",
    )
    _extract_and_save_by_job(
        df_to_analyze,
        save_path_high_views,
        job="조회",
        threshold=VIEW_THRESHOLD,
        job_column_name=COL_JOB_PERFORMANCE,
    )
    print(
        f"High-volume view check (>{VIEW_THRESHOLD} views) results processing attempted."
    )

    # 필터 3: 대량의 기록을 저장하는 사용자.
    # 원본 주석: "저장 SAVE_THRESHOLD 이상 교번별 전체 기록 시트별 저장"
    base_report_name_for_saves = MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE.split(
        "."
    )[0].format(prev_month)
    save_path_high_saves = os.path.join(
        save_dir,
        f"{base_report_name_for_saves}({PERSONAL_INFO_ACCESS_HIGH_VOLUME_SAVES_SUFFIX}).xlsx",
    )
    _extract_and_save_by_job(
        df_to_analyze,
        save_path_high_saves,
        job="저장",
        threshold=SAVE_THRESHOLD,
        job_column_name=COL_JOB_PERFORMANCE,
    )
    print(
        f"High-volume save check (>{SAVE_THRESHOLD} saves) results processing attempted."
    )


def _filter_by_job_master_exclude_detail_id(df: pd.DataFrame) -> pd.DataFrame:
    """
    '인사마스터' (HR 마스터) 프로그램 접근 기록을 필터링하며, 사용자의 ID가
    '상세내용' 필드에 나타나는 경우(즉, 본인 접근)는 제외합니다.

    매개변수:
        df (pd.DataFrame): 개인 정보 접근 로그를 포함하는 DataFrame입니다.
                           예상 열: `COL_PROGRAM_NAME`, `COL_EMPLOYEE_ID` (또는 `COL_EMPLOYEE_ID_LOGIN`),
                           `COL_ACCESS_TIME`, `COL_DETAIL_CONTENT`.

    반환 값:
        pd.DataFrame: 필터링된 기록을 포함하며 직원 ID와 접근 시간으로 정렬된 DataFrame입니다.
                      해당 기록이 없으면 빈 DataFrame을 반환합니다.

    예외:
        ValueError: 필터링에 필수적인 열이 누락된 경우.
    """
    # 사용할 직원 ID 열('교번' 또는 '신분번호')을 결정합니다.
    employee_id_col_to_use = COL_EMPLOYEE_ID
    if COL_EMPLOYEE_ID not in df.columns and COL_EMPLOYEE_ID_LOGIN in df.columns:
        employee_id_col_to_use = COL_EMPLOYEE_ID_LOGIN
    elif COL_EMPLOYEE_ID not in df.columns:
        raise ValueError(
            f"Required employee ID column ('{COL_EMPLOYEE_ID}' or '{COL_EMPLOYEE_ID_LOGIN}') not found in DataFrame."
        )

    required_cols = [
        COL_PROGRAM_NAME,
        employee_id_col_to_use,
        COL_ACCESS_TIME,
        COL_DETAIL_CONTENT,
    ]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(
                f"Required column '{col}' not found in DataFrame for HR Master filtering."
            )

    # '인사마스터'(HR 마스터) 프로그램 접근을 필터링합니다. '인사마스터'는 특정 프로그램 이름입니다.
    # 원본 주석: "2. '프로그램명' == '인사마스터' 필터"
    filtered_df = df[df[COL_PROGRAM_NAME] == "인사마스터"]

    # 직원의 ID가 상세 내용(본인 접근) 내에서 발견되는 기록을 제외합니다.
    # 원본 주석: "3. 교번이 상세내용에 포함되어 있지 않은 것만 남기기"
    # 이 람다 함수는 사용자 ID의 문자열 표현이 상세 내용 문자열의 일부인지 확인합니다.
    # 강력한 비교를 위해 둘 다 문자열로 처리되도록 합니다.
    filtered_df = filtered_df[
        ~filtered_df.apply(
            lambda row: str(row[employee_id_col_to_use])
            in str(row[COL_DETAIL_CONTENT]),
            axis=1,
        )
    ]

    # 결과를 정렬합니다.
    # 원본 주석: "4. 교번, 접속일시 정렬"
    return filtered_df.sort_values([employee_id_col_to_use, COL_ACCESS_TIME])


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
        job_column_name (str): `df`에서 작업 유형을 포함하는 열의 이름입니다 (예: `COL_JOB_PERFORMANCE`).

    예외:
        ValueError: 필수 열(`COL_EMPLOYEE_ID` 또는 대체 열, `COL_EMPLOYEE_NAME`, `job_column_name`)이 누락된 경우.
    """
    # 사용할 직원 ID 열('교번' 또는 '신분번호')을 결정합니다.
    employee_id_col_to_use = COL_EMPLOYEE_ID
    if COL_EMPLOYEE_ID not in df.columns and COL_EMPLOYEE_ID_LOGIN in df.columns:
        employee_id_col_to_use = COL_EMPLOYEE_ID_LOGIN
    elif COL_EMPLOYEE_ID not in df.columns:
        raise ValueError(
            f"Required employee ID column ('{COL_EMPLOYEE_ID}' or '{COL_EMPLOYEE_ID_LOGIN}') not found."
        )

    required_cols_check = [
        employee_id_col_to_use,
        COL_EMPLOYEE_NAME,
        job_column_name,
    ]
    for col in required_cols_check:
        if col not in df.columns:
            raise ValueError(
                f"Required column '{col}' not found in DataFrame for job extraction."
            )

    # 특정 작업 유형과 일치하는 기록을 필터링합니다.
    job_specific_df = df[df[job_column_name] == job]
    # 직원 ID별로 그룹화하고 작업 발생 횟수를 계산합니다.
    job_counts_per_user = job_specific_df.groupby(
        employee_id_col_to_use).size()
    # 이 작업에 대한 임계값을 충족하거나 초과하는 사용자를 식별합니다.
    target_user_ids = job_counts_per_user[
        job_counts_per_user >= threshold
    ].index.tolist()

    if not target_user_ids:
        print(
            f"No users found exceeding threshold of {threshold} for job '{job}' in '{save_path}'."
        )
        # 원하는 경우 빈 파일 또는 알림이 있는 파일을 만듭니다.
        # 지금은 메시지만 출력합니다. 빈 파일이 필요한 경우:
        # with pd.ExcelWriter(save_path) as writer:
        #     pd.DataFrame().to_excel(writer, sheet_name="NoData", index=False)
        return

    # 여러 시트를 저장하기 위해 ExcelWriter를 만듭니다.
    with pd.ExcelWriter(save_path) as writer:
        for employee_id in target_user_ids:
            # 식별된 사용자의 모든 기록(작업별 기록뿐만 아니라)을 가져옵니다.
            user_all_records_df = df[df[employee_id_col_to_use] == employee_id]

            # 시트 이름에 사용할 사용자 이름을 가져오려고 시도합니다.
            user_name = ""
            if (
                not user_all_records_df.empty
                and COL_EMPLOYEE_NAME in user_all_records_df.columns
            ):
                user_name = user_all_records_df[COL_EMPLOYEE_NAME].iloc[0]

            # Excel의 문자 제한 내에서 시트 이름을 만듭니다.
            sheet_name = f"{employee_id}_{user_name}"
            if len(sheet_name) > SHEET_NAME_MAX_CHARS:
                # 필요한 경우 잘라내고 고유성을 보장하지만 여기서는 단순 잘라내기를 사용합니다.
                sheet_name = sheet_name[:SHEET_NAME_MAX_CHARS]

            user_all_records_df.to_excel(
                writer, sheet_name=sheet_name, index=False)
            # 자동 맞춤 참고: 현재 `save_excel_with_autofit` 유틸리티는 DataFrame을 새 파일에 저장합니다.
            # 그런 다음 자동 맞춤합니다. ExcelWriter 컨텍스트 내의 시트에 직접 적용하려면
            # 유틸리티를 수정하거나 저장된 다중 시트 파일을 후처리해야 합니다.
            # 이는 이전 docstring 업데이트에서도 언급되었습니다.

    print(
        f"Saved {len(target_user_ids)} users' data to separate sheets in: {save_path}"
    )
