"""
이 모듈은 사용자 로그인 기록 데이터에 대한 검사를 수행합니다.
다음과 같은 의심스러운 로그인 패턴을 식별합니다:
- 짧은 시간 내 여러 IP 주소에서의 로그인.
- 업무 시간 외 로그인.
- 공휴일 및 주말 로그인.

메인 함수 `login_checker`는 이러한 검사를 조정하고 결과를
별도의 Excel 파일에 저장합니다.
"""

import os
from datetime import datetime

import pandas as pd

from display import print_checker_header
from utils import (
    filter_by_time_conditions,
    find_and_prepare_excel_file,
    run_and_save_check,
)

# Constants for login_checker.py
LOGIN_LOG_FILE_PREFIX = "사용자접속내역_Login내역_"  # 사용자 로그인 기록 파일의 접두사
LOGIN_CHECK_REPORT_BASE = "[붙임2] 코러스 사용자 접근 기록"  # 로그인 점검 보고서
# 보고서 접미사: LOGIN_IP_SWITCH_WINDOW_HOURS 내에 여러 IP에서 로그인하는 사용자
LOGIN_REPORT_IP_SWITCH_SUFFIX = "60분IP"
LOGIN_REPORT_OFF_HOURS_SUFFIX = "업무시간외"  # 보고서 접미사: 표준 근무 시간 외 로그인
LOGIN_REPORT_HOLIDAY_SUFFIX = "휴일"  # 보고서 접미사: 공휴일 또는 주말 로그인
COL_IP = "IP"  # IP 주소
COL_ACCESS_TIME = "접근일시"  # 접근 타임스탬프 (예: "YYYY-MM-DD HH:MM:SS")
COL_EMPLOYEE_ID_LOGIN = "신분번호"  # 직원 ID, 특히 로그인 기록 파일에서 발견됨.
# login_checker용: 동일 사용자에 대해 여러 IP에서의 로그인을
# 감지하기 위한 시간 창(시간 단위).
LOGIN_IP_SWITCH_WINDOW_HOURS = 1
LOGIN_IP_SWITCH_MIN_IPS = (
    3  # login_checker용: IP 변경 알림을 트리거하기 위한 창 내 최소 고유 IP 수.
)
LOGIN_OFF_HOURS_START = (
    23  # 로그인 업무 시간 외 시작 시간(포함, 24시간 형식) (예: 오후 11시)
)
LOGIN_OFF_HOURS_END = (
    7  # 로그인 업무 시간 외 종료 시간(미포함, 24시간 형식) (예: 오전 7시 이전 활동)
)


def login_checker(download_dir: str, save_dir: str, prev_month: str) -> int:
    """
    로그인 기록 데이터에 대한 다양한 검사를 수행하는 메인 함수입니다.

    로그인 기록 Excel 파일을 읽은 후 다음 필터를 적용합니다:
    1. IP 주소 변경: 정의된 시간 내에 여러 IP에서 로그인하는 사용자.
    2. 업무 시간 외 접근: 표준 업무 시간 외에 발생한 로그인.
    3. 공휴일/주말 접근: 공휴일 또는 주말에 발생한 로그인.

    필터링된 각 결과 집합은 별도의 Excel 파일에 저장됩니다.

    매개변수:
        download_dir (str): 원본 로그인 기록 Excel 파일이 있는 디렉토리입니다.
        save_dir (str): 생성된 보고서 Excel 파일이 저장될 디렉토리입니다.
        prev_month (str): 'YYYYMM' 형식의 이전 달로, 출력 파일 이름 지정에 사용됩니다.

    반환 값:
        int: 처리된 원본 데이터의 행 수입니다. 파일을 찾을 수 없으면 0입니다.

    예외:
        ValueError: 예상되는 'IP' 열이 10번째 위치(인덱스 9)에 없는 경우.
    """
    print_checker_header(LOGIN_CHECK_REPORT_BASE)

    file_prefix = f"{LOGIN_LOG_FILE_PREFIX}{datetime.today().strftime('%Y%m')}"

    df, _ = find_and_prepare_excel_file(
        download_dir,
        file_prefix,
        save_dir,
        LOGIN_CHECK_REPORT_BASE,
        prev_month,
    )

    if df is None:
        return 0

    expected_ip_col_index = 9
    if df.columns[expected_ip_col_index] != COL_IP:
        raise ValueError(
            f"'{COL_IP}' 컬럼이 {expected_ip_col_index} 위치에 없습니다. "
            f"실제 컬럼: {df.columns[expected_ip_col_index]}"
        )

    checks_to_run = [
        {
            "function": _filter_ip_switch,
            "suffix": LOGIN_REPORT_IP_SWITCH_SUFFIX,
            "description": (
                f"{LOGIN_IP_SWITCH_WINDOW_HOURS}시간 내 "
                f"{LOGIN_IP_SWITCH_MIN_IPS}개 이상 IP 사용"
            ),
        },
        {
            "function": lambda df: filter_by_time_conditions(
                df,
                time_col=COL_ACCESS_TIME,
                employee_id_col=COL_EMPLOYEE_ID_LOGIN,
                check_off_hours=True,
                check_holidays_weekends=False,
                off_hours_start=LOGIN_OFF_HOURS_START,
                off_hours_end=LOGIN_OFF_HOURS_END,
            ),
            "suffix": LOGIN_REPORT_OFF_HOURS_SUFFIX,
            "description": "업무 시간 외 로그인",
        },
        {
            "function": lambda df: filter_by_time_conditions(
                df,
                time_col=COL_ACCESS_TIME,
                employee_id_col=COL_EMPLOYEE_ID_LOGIN,
                check_off_hours=False,
                check_holidays_weekends=True,
                off_hours_start=0,  # Not used
                off_hours_end=0,  # Not used
            ),
            "suffix": LOGIN_REPORT_HOLIDAY_SUFFIX,
            "description": "휴일/주말 로그인",
        },
    ]

    for check in checks_to_run:
        save_path = os.path.join(
            save_dir,
            f"{LOGIN_CHECK_REPORT_BASE}({check['suffix']})_{prev_month}.xlsx",
        )
        run_and_save_check(
            df=df,
            check_func=check["function"],
            save_path=save_path,
            result_description=str(check["description"]),
        )

    return len(df)


def _filter_ip_switch(df: pd.DataFrame) -> pd.DataFrame:
    """
    정의된 시간 창 내에 여러 개의 고유 IP 주소에서 로그인한 사용자를 필터링합니다.

    매개변수:
        df (pd.DataFrame): 로그인 기록을 포함하는 DataFrame입니다. 예상되는 열에는
                           `COL_ACCESS_TIME`(접근 타임스탬프) 및 `COL_IP`(IP 주소),
                           그리고 `COL_EMPLOYEE_ID_LOGIN`(직원 식별자)이 포함됩니다.

    반환 값:
        pd.DataFrame: IP 변경 알림을 트리거한 사용자 기록을 포함하는 DataFrame으로,
                      직원 ID와 접근 시간으로 정렬됩니다.
                      해당 기록이 없으면 빈 DataFrame을 반환합니다.

    예외:
        ValueError: 입력 DataFrame `df`가 None인 경우.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()

    flagged_indices = (
        set()
    )  # 중복을 피하기 위해 플래그가 지정된 행의 인덱스를 저장하는 데 세트를 사용합니다.

    # 직원 ID별로 그룹화하여 각 사용자의 로그인 패턴을 분석합니다.
    for _, group in df_copy.groupby(COL_EMPLOYEE_ID_LOGIN):
        group = group.sort_values(COL_ACCESS_TIME)

        # 사용자의 각 로그인 이벤트를 반복합니다.
        for i in range(len(group)):
            current_login_time = group.iloc[i][COL_ACCESS_TIME]
            # 후속 로그인을 확인하기 위한 시간 창을 정의합니다.
            window_end_time = current_login_time + pd.Timedelta(
                hours=LOGIN_IP_SWITCH_WINDOW_HOURS
            )

            # 이 창 내의 로그인을 선택합니다.
            logins_in_window = group[
                (group[COL_ACCESS_TIME] >= current_login_time)
                & (group[COL_ACCESS_TIME] <= window_end_time)
            ]

            # 이 창 내의 고유 IP 수가 임계값을 충족하는지 확인합니다.
            if len(set(logins_in_window[COL_IP])) >= LOGIN_IP_SWITCH_MIN_IPS:
                flagged_indices.update(
                    logins_in_window.index
                )  # 이 창의 모든 기록을 추가합니다.

    if flagged_indices:
        result_df = df_copy.loc[sorted(flagged_indices)]
        return result_df.sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])
    else:
        return pd.DataFrame(
            columns=df.columns
        )  # 일치하는 항목이 없으면 동일한 열을 가진 빈 DataFrame을 반환합니다.
