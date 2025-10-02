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

import pandas as pd

import config as cfg
from display import print_checker_header
from utils import (
    filter_by_time_conditions,
    find_and_prepare_excel_file,
    run_and_save_check,
)


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


def sayu_checker(download_dir: str, save_dir: str, prev_month: str) -> int:
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

    반환 값:
        int: 처리된 원본 데이터의 행 수입니다. 파일을 찾을 수 없으면 0입니다.
    """
    print_checker_header(cfg.DOWNLOAD_REASON_REPORT_BASE)

    file_prefix = (
        f"{cfg.PERSONAL_INFO_DOWNLOAD_REASON_PREFIX}{datetime.today().strftime('%Y%m')}"
    )

    df, _ = find_and_prepare_excel_file(
        download_dir,
        file_prefix,
        save_dir,
        cfg.DOWNLOAD_REASON_REPORT_BASE,
        prev_month,
    )

    if df is None:
        return 0

    checks_to_run = [
        {
            "function": _check_download_sayu,
            "suffix": cfg.DOWNLOAD_REASON_INVALID_REASON_SUFFIX,
            "description": "다운로드 사유 비정상",
        },
        {
            "function": _filter_high_download_users,
            "suffix": cfg.DOWNLOAD_REASON_HIGH_DOWNLOAD_COUNT_SUFFIX,
            "description": f"다운로드 {cfg.DOWNLOAD_COUNT_THRESHOLD}건 초과",
        },
        {
            "function": _filter_high_freq_download,
            "suffix": cfg.DOWNLOAD_REASON_HIGH_FREQUENCY_SUFFIX,
            "description": (
                f"1시간 내 {cfg.DOWNLOAD_FREQUENCY_THRESHOLD}건 이상 다운로드"
            ),
        },
        {
            "function": lambda df: filter_by_time_conditions(
                df,
                time_col=cfg.COL_ACCESS_TIME,
                employee_id_col=cfg.COL_EMPLOYEE_ID,
                check_off_hours=True,
                check_holidays_weekends=True,
                off_hours_start=cfg.DOWNLOAD_OFF_HOURS_START,
                off_hours_end=cfg.DOWNLOAD_OFF_HOURS_END,
            ),
            "suffix": cfg.DOWNLOAD_REASON_OFF_HOURS_SUFFIX,
            "description": "업무 시간 외/휴일 다운로드",
        },
    ]

    for check in checks_to_run:
        save_path = os.path.join(
            save_dir,
            f"{cfg.DOWNLOAD_REASON_REPORT_BASE}({check['suffix']})_{prev_month}.xlsx",
        )
        run_and_save_check(
            df=df,
            check_func=check["function"],
            save_path=save_path,
            result_description=str(check["description"]),
        )

    return len(df)


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
    if df.columns[expected_reason_col_index] != cfg.COL_DOWNLOAD_REASON:
        raise ValueError(
            (
                f"'{cfg.COL_DOWNLOAD_REASON}' 컬럼이 "
                f"{expected_reason_col_index} 위치에 없습니다. "
                f"실제 컬럼: {df.columns[expected_reason_col_index]}"
            )
        )

    # 다운로드 사유의 고유 문자 수에 대한 필터를 적용합니다.
    # 원본 주석: "5. 고유 문자 개수 5개 이하인 row 필터링"
    filtered_df = df[df[cfg.COL_DOWNLOAD_REASON].apply(_unique_char_count_below_5)]
    return filtered_df.sort_values([cfg.COL_EMPLOYEE_ID, cfg.COL_ACCESS_TIME])


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
    if df.columns[expected_count_col_index] != cfg.COL_DOWNLOAD_COUNT:
        raise ValueError(
            (
                f"'{cfg.COL_DOWNLOAD_COUNT}' 컬럼이 "
                f"{expected_count_col_index} 위치에 없습니다. "
                f"실제 컬럼: {df.columns[expected_count_col_index]}"
            )
        )

    # 직원 ID별로 그룹화하고 다운로드 수를 합산합니다.
    download_sum_per_user = df.groupby(cfg.COL_EMPLOYEE_ID)[
        cfg.COL_DOWNLOAD_COUNT
    ].sum()
    # 다운로드 수 임계값을 충족하거나 초과하는 사용자를 식별합니다.
    target_users = download_sum_per_user[
        download_sum_per_user >= cfg.DOWNLOAD_COUNT_THRESHOLD
    ].index

    # 식별된 사용자의 모든 기록을 반환합니다.
    return df[df[cfg.COL_EMPLOYEE_ID].isin(target_users)].sort_values(
        [cfg.COL_EMPLOYEE_ID, cfg.COL_ACCESS_TIME]
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

    flagged_indices = (
        set()
    )  # 높은 빈도의 다운로드 폭주에 속하는 기록의 원본 인덱스를 저장합니다.

    # 직원 ID별로 그룹화하여 각 사용자의 다운로드 패턴을 분석합니다.
    for _, group in df_copy.groupby(cfg.COL_EMPLOYEE_ID):
        # 정수 인덱스 i와 함께 .loc를 사용하기 위해 인덱스를 재설정합니다.
        group = group.sort_values(cfg.COL_ACCESS_TIME).reset_index()

        for i in range(len(group)):
            current_download_time = cast(
                pd.Timestamp, group.loc[i, cfg.COL_ACCESS_TIME]
            )
            # 현재 다운로드 시간으로부터 1시간 창을 정의합니다.
            window_end_time = current_download_time + pd.Timedelta(hours=1)

            # 이 1시간 창 내의 다운로드를 선택합니다.
            downloads_in_window = group[
                (group[cfg.COL_ACCESS_TIME] >= current_download_time)
                & (group[cfg.COL_ACCESS_TIME] <= window_end_time)
            ]

            # 이 창 내의 다운로드 수가 빈도 임계값을 충족하면 플래그를 지정합니다.
            if len(downloads_in_window) >= cfg.DOWNLOAD_FREQUENCY_THRESHOLD:
                flagged_indices.update(
                    downloads_in_window["index"].tolist()
                )  # reset_index() 후 저장된 원본 인덱스를 사용합니다.

    if flagged_indices:
        result_df = df_copy.loc[
            sorted(flagged_indices)
        ]  # 원본 인덱스를 사용하여 선택합니다.
        return result_df.sort_values([cfg.COL_EMPLOYEE_ID, cfg.COL_ACCESS_TIME])
    else:
        return pd.DataFrame(columns=df.columns)  # 동일한 열을 가진 빈 DataFrame 반환
