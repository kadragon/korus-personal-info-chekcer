"""
이 모듈은 개인 정보 다운로드 사유를 확인하는 역할을 합니다.
개인 데이터 다운로드 로그를 분석하고 다음과 같은 의심스러운 활동을 표시합니다:
- 매우 짧거나 단순한 사유로 다운로드 (예: "asdfg", "12345").
- 사용자가 과도한 총 기록 수를 다운로드하는 경우.
- 사용자가 짧은 시간 내에 비정상적으로 높은 빈도로 데이터를 다운로드하는 경우.
- 표준 업무 시간 외 또는 공휴일/주말에 다운로드하는 경우.

메인 함수 `sayu_checker`("사유 검사기")는 이러한 검사를 수행하고
필터링된 결과를 별도의 Excel 파일에 저장합니다.
"""

import os
from datetime import datetime
from typing import cast

import holidays
import pandas as pd

from utils import find_and_prepare_excel_file, save_excel_with_autofit

# download_reason_checker.py 상수
PERSONAL_INFO_DOWNLOAD_REASON_PREFIX = "개인정보 다운로드 사유 조회_"
DOWNLOAD_REASON_REPORT_BASE = "[붙임4] 개인정보 다운로드 사유"
DOWNLOAD_REASON_INVALID_REASON_SUFFIX = "사유이상"
DOWNLOAD_REASON_HIGH_DOWNLOAD_COUNT_SUFFIX = "100건 초과"
DOWNLOAD_REASON_HIGH_FREQUENCY_SUFFIX = "1시간20건초과"
DOWNLOAD_REASON_OFF_HOURS_SUFFIX = "업무시간외"
COL_ACCESS_TIME = "접속일시"
COL_EMPLOYEE_ID = "교번"
COL_DOWNLOAD_REASON = "다운로드사유"
COL_DOWNLOAD_COUNT = "다운로드데이터수(건)"
DOWNLOAD_COUNT_THRESHOLD = 100
DOWNLOAD_FREQUENCY_THRESHOLD = 20
DOWNLOAD_OFF_HOURS_START = 23
DOWNLOAD_OFF_HOURS_END = 8


def _unique_char_count_below_5(text_input) -> bool:
    """
    주어진 문자열의 고유 문자 수가 5개 이하인지 확인합니다.
    이는 잠재적으로 의심스럽거나 설명이 부족한 다운로드 사유를 식별하는 데 사용됩니다.

    매개변수:
        text_input: 확인할 문자열입니다. 일반적으로 다운로드 사유입니다.

    반환 값:
        bool: 고유 문자 수가 5개 이하이면 True, 그렇지 않으면 False입니다.
              입력이 NaN(Not a Number)이면 False를 반환합니다.
    """
    if pd.isna(text_input):
        return False
    return len(set(str(text_input))) <= 5


def sayu_checker(download_dir: str, save_dir: str, prev_month: str):
    """
    개인 정보 다운로드 사유에서 의심스러운 패턴을 확인하는 메인 함수입니다.

    다운로드 사유 로그 파일을 읽고 여러 필터를 적용합니다:
    1. 유효하지 않거나 짧은 다운로드 사유.
    2. 사용자별 높은 총 다운로드 수.
    3. 한 시간 내 사용자별 높은 다운로드 빈도.
    4. 업무 시간 외 또는 공휴일/주말 다운로드.

    필터링된 각 결과 집합은 별도의 Excel 파일에 저장됩니다.

    매개변수:
        download_dir (str): 원본 다운로드 사유 Excel 파일이 있는 디렉토리입니다.
        save_dir (str): 생성된 보고서 Excel 파일이 저장될 디렉토리입니다.
        prev_month (str): 'YYYYMM' 형식의 이전 달로, 출력 파일 이름 지정에 사용됩니다.

    예외:
        FileNotFoundError: 다운로드 사유 Excel 파일을 찾을 수 없는 경우.
    """
    # 다운로드 사유 Excel 파일을 찾아 복사하고 읽어들입니다.
    file_prefix = (
        f"{PERSONAL_INFO_DOWNLOAD_REASON_PREFIX}{datetime.today().strftime('%Y%m')}"
    )

    df, _ = find_and_prepare_excel_file(
        download_dir,
        file_prefix,
        save_dir,
        DOWNLOAD_REASON_REPORT_BASE,
        prev_month,
    )

    if df is None:
        raise FileNotFoundError(
            f"Download reason Excel file starting with '{file_prefix}' "
            f"not found in '{download_dir}'."
        )

    # 의심스럽거나 짧은 사유의 다운로드를 필터링합니다.
    # 원본 주석: "사유 비정상 작성"
    filtered_invalid_reason = _check_download_sayu(df)
    if not filtered_invalid_reason.empty:
        save_path_invalid_reason = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_INVALID_REASON_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_invalid_reason, save_path_invalid_reason)
        print(
            f"Results for invalid download reasons saved to: {save_path_invalid_reason}"
        )
    else:
        print("No records found for invalid download reason check.")

    # 총 다운로드 기록 수가 많은 사용자를 필터링합니다.
    # 원본 주석: "100건 이상 개인정보 다운로드"
    filtered_high_download = _filter_high_download_users(df)
    if not filtered_high_download.empty:
        save_path_high_download = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_HIGH_DOWNLOAD_COUNT_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_high_download, save_path_high_download)
        print(
            f"Results for high download count (>{DOWNLOAD_COUNT_THRESHOLD}) "
            f"saved to: {save_path_high_download}"
        )
    else:
        print(
            f"No records found for high download count "
            f"(>{DOWNLOAD_COUNT_THRESHOLD}) check."
        )

    # 한 시간 내 다운로드 빈도가 높은 사용자를 필터링합니다.
    # 원본 주석: "1시간 이내 다운로드 횟수 20건 이상"
    filtered_high_freq = _filter_high_freq_download(df)
    if not filtered_high_freq.empty:
        save_path_high_freq = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_HIGH_FREQUENCY_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_high_freq, save_path_high_freq)
        print(
            f"Results for high download frequency (>{DOWNLOAD_FREQUENCY_THRESHOLD}/hr) "
            f"saved to: {save_path_high_freq}"
        )
    else:
        print(
            f"No records found for high download frequency "
            f"(>{DOWNLOAD_FREQUENCY_THRESHOLD}/hr) check."
        )

    # 업무 시간 외 또는 공휴일/주말에 발생한 다운로드를 필터링합니다.
    # 원본 주석: "업무시간 외 다운로드"
    filtered_off_hours_holiday = _filter_off_hour_and_holiday(df)
    if not filtered_off_hours_holiday.empty:
        save_path_off_hours_holiday = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_OFF_HOURS_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_off_hours_holiday, save_path_off_hours_holiday)
        print(
            f"Results for off-hours/holiday downloads saved to: "
            f"{save_path_off_hours_holiday}"
        )
    else:
        print("No records found for off-hours/holiday download check.")


def _check_download_sayu(df: pd.DataFrame) -> pd.DataFrame:
    """
    다운로드 사유가 의심스러운(너무 짧거나 단순한) 기록을 필터링합니다.
    검사를 위해 `_unique_char_count_below_5` 헬퍼 함수를 사용합니다.

    매개변수:
        df (pd.DataFrame): 다운로드 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_DOWNLOAD_REASON`(다운로드 사유),
                           `COL_EMPLOYEE_ID`(직원 ID),
                           `COL_ACCESS_TIME`(접근 타임스탬프).

    반환 값:
        pd.DataFrame: 의심스러운 다운로드 사유가 있는 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 필터링된 DataFrame입니다.

    예외:
        ValueError: 예상되는 다운로드 사유 열(`COL_DOWNLOAD_REASON`)이
                    5번째 위치(인덱스 4)에 없는 경우.
    """
    expected_reason_col_index = 4
    if df.columns[expected_reason_col_index] != COL_DOWNLOAD_REASON:
        raise ValueError(
            f"Expected '{COL_DOWNLOAD_REASON}' column at index "
            f"{expected_reason_col_index}. Found: "
            f"{df.columns[expected_reason_col_index]}"
        )

    # 다운로드 사유의 고유 문자 수에 대한 필터를 적용합니다.
    # 원본 주석: "5. 고유 문자 개수 5개 이하인 row 필터링"
    return df[df[COL_DOWNLOAD_REASON].apply(_unique_char_count_below_5)].sort_values(
        [COL_EMPLOYEE_ID, COL_ACCESS_TIME]
    )


def _filter_high_download_users(df: pd.DataFrame) -> pd.DataFrame:
    """
    총 다운로드 기록 수가 정의된 임계값을 초과하는 사용자를 필터링합니다.

    매개변수:
        df (pd.DataFrame): 다운로드 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_DOWNLOAD_COUNT`(다운로드된 기록 수),
                           `COL_EMPLOYEE_ID`(직원 ID),
                           `COL_ACCESS_TIME`(접근 타임스탬프).

    반환 값:
        pd.DataFrame: 임계값을 초과한 사용자의 모든 다운로드 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 DataFrame입니다.

    예외:
        ValueError: 예상되는 다운로드 수 열(`COL_DOWNLOAD_COUNT`)이
                    6번째 위치(인덱스 5)에 없는 경우.
    """
    expected_count_col_index = 5
    if df.columns[expected_count_col_index] != COL_DOWNLOAD_COUNT:
        raise ValueError(
            f"Expected '{COL_DOWNLOAD_COUNT}' column at index "
            f"{expected_count_col_index}. Found: {df.columns[expected_count_col_index]}"
        )

    # 직원 ID별로 그룹화하고 다운로드 수를 합산합니다.
    download_sum_per_user = df.groupby(COL_EMPLOYEE_ID)[COL_DOWNLOAD_COUNT].sum()
    # 다운로드 수 임계값을 충족하거나 초과하는 사용자를 식별합니다.
    target_users = download_sum_per_user[
        download_sum_per_user >= DOWNLOAD_COUNT_THRESHOLD
    ].index

    # 식별된 사용자의 모든 기록을 반환합니다.
    return df[df[COL_EMPLOYEE_ID].isin(target_users)].sort_values(
        [COL_EMPLOYEE_ID, COL_ACCESS_TIME]
    )


def _filter_high_freq_download(df: pd.DataFrame) -> pd.DataFrame:
    """
    높은 빈도(한 시간 내에 임계값 횟수 이상)로 데이터를 다운로드한
    사용자를 필터링합니다.

    매개변수:
        df (pd.DataFrame): 다운로드 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_ACCESS_TIME`(접근 타임스탬프),
                           `COL_EMPLOYEE_ID`(직원 ID).

    반환 값:
        pd.DataFrame: 높은 빈도의 다운로드 폭주에 해당하는 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 DataFrame입니다.
                      해당 폭주가 없으면 빈 DataFrame을 반환합니다.

    예외:
        ValueError: 입력 DataFrame `df`가 None인 경우.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    flagged_indices = (
        set()
    )  # 높은 빈도의 다운로드 폭주에 속하는 기록의 원본 인덱스를 저장합니다.

    # 직원 ID별로 그룹화하여 각 사용자의 다운로드 패턴을 분석합니다.
    for _, group in df_copy.groupby(COL_EMPLOYEE_ID):
        # 정수 인덱스 i와 함께 .loc를 사용하기 위해 인덱스를 재설정합니다.
        group = group.sort_values(COL_ACCESS_TIME).reset_index()

        for i in range(len(group)):
            current_download_time = cast(pd.Timestamp, group.loc[i, COL_ACCESS_TIME])
            # 현재 다운로드 시간으로부터 1시간 창을 정의합니다.
            window_end_time = current_download_time + pd.Timedelta(hours=1)

            # 이 1시간 창 내의 다운로드를 선택합니다.
            downloads_in_window = group[
                (group[COL_ACCESS_TIME] >= current_download_time)
                & (group[COL_ACCESS_TIME] <= window_end_time)
            ]

            # 이 창 내의 다운로드 수가 빈도 임계값을 충족하면 플래그를 지정합니다.
            if len(downloads_in_window) >= DOWNLOAD_FREQUENCY_THRESHOLD:
                flagged_indices.update(
                    downloads_in_window["index"].tolist()
                )  # reset_index() 후 저장된 원본 인덱스를 사용합니다.

    if flagged_indices:
        result_df = df_copy.loc[
            sorted(flagged_indices)
        ]  # 원본 인덱스를 사용하여 선택합니다.
        return result_df.sort_values([COL_EMPLOYEE_ID, COL_ACCESS_TIME])
    else:
        return pd.DataFrame(columns=df.columns)  # 동일한 열을 가진 빈 DataFrame 반환


def _filter_off_hour_and_holiday(df: pd.DataFrame) -> pd.DataFrame:
    """
    표준 업무 시간 외 또는 대한민국 공휴일/주말에 발생한 다운로드 기록을 필터링합니다.
    업무 시간 외는 `DOWNLOAD_OFF_HOURS_START` 및 `DOWNLOAD_OFF_HOURS_END`로 정의됩니다.

    매개변수:
        df (pd.DataFrame): 다운로드 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_ACCESS_TIME`(접근 타임스탬프),
                           `COL_EMPLOYEE_ID`(직원 ID).

    반환 값:
        pd.DataFrame: 업무 시간 외 또는 공휴일/주말의 다운로드 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 DataFrame입니다.

    예외:
        ValueError: 입력 DataFrame `df`가 None인 경우.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # 해당 연도에 대한 대한민국 공휴일을 초기화합니다.
    years = df_copy[COL_ACCESS_TIME].dt.year.unique()
    kr_holidays = holidays.KR(years=years)  # type: ignore [attr-defined]

    # 검사를 위한 시간적 특징을 추출합니다.
    weekday = df_copy[COL_ACCESS_TIME].dt.weekday
    hour = df_copy[COL_ACCESS_TIME].dt.hour
    date_only = df_copy[COL_ACCESS_TIME].dt.date  # 공휴일 확인용

    # 업무 시간 외, 주말 및 공휴일 조건을 정의합니다.
    is_off_hour = (hour < DOWNLOAD_OFF_HOURS_END) | (hour >= DOWNLOAD_OFF_HOURS_START)
    is_weekend = weekday >= 5  # 월요일은 0이고 일요일은 6입니다; 토요일=5, 일요일=6.
    is_holiday = date_only.isin(kr_holidays)

    # 조건 결합: 업무 시간 외 또는 주말 또는 공휴일인 모든 기록.
    mask = is_off_hour | is_weekend | is_holiday
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID, COL_ACCESS_TIME])
