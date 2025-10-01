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

import holidays
import pandas as pd

from utils import find_and_prepare_excel_file, save_excel_with_autofit

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


def login_checker(download_dir: str, save_dir: str, prev_month: str):
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
        prev_month (str): 'YYYYMM' 형식의 이전 달로, 출력 파일 이름 지정 및
                          `find_and_prepare_excel_file`에서 처리하지 않은 경우
                          올바른 입력 파일을 선택하는 데 사용될 수 있습니다.

    예외:
        FileNotFoundError: 지정된 로그인 기록 Excel 파일을 찾을 수 없는 경우.
        ValueError: 예상되는 'IP' 열이 10번째 위치(인덱스 9)에 없는 경우.
    """
    file_prefix = f"{LOGIN_LOG_FILE_PREFIX}{datetime.today().strftime('%Y%m')}"

    # 유틸리티 함수를 사용하여 로그인 기록 Excel 파일을 찾고, 복사하고, 읽습니다.
    # 복사된 파일은 save_dir에 표준화된 이름으로 저장됩니다.
    df, _ = find_and_prepare_excel_file(
        download_dir,
        file_prefix,
        save_dir,
        LOGIN_CHECK_REPORT_BASE,
        prev_month,
    )

    if df is None:
        # find_and_prepare_excel_file은 파일이 없는 경우 이미 경고를 출력합니다.
        # 기본 데이터 소스가 누락된 경우 실행을 중지하기 위해 이 오류가 발생합니다.
        raise FileNotFoundError(
            f"Login history Excel file starting with "
            f"'{file_prefix}' not found in '{download_dir}'."
        )

    # 10번째 열(인덱스 9)이 'IP'인지 확인합니다.
    # 이는 예상 파일 형식에 기반한 온전성 검사입니다.
    # 원본 주석: "5. 컬럼명 확인"
    expected_ip_col_index = 9
    if df.columns[expected_ip_col_index] != COL_IP:
        raise ValueError(
            f"Expected '{COL_IP}' column at index {expected_ip_col_index}. "
            f"Found: {df.columns[expected_ip_col_index]}"
        )

    # 짧은 시간 내에 여러 IP에서 로그인하는 사용자를 필터링합니다.
    # 원본 주석: "6. 60분 이내에 다른 IP 접속"
    filtered_ip_switch = _filter_ip_switch(df)
    if not filtered_ip_switch.empty:
        save_path_ip_switch = os.path.join(
            save_dir,
            f"{LOGIN_CHECK_REPORT_BASE}({LOGIN_REPORT_IP_SWITCH_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_ip_switch, save_path_ip_switch)
        print(f"IP switch check results saved to: {save_path_ip_switch}")
    else:
        print("No records found for IP switch check.")

    # 표준 업무 시간 외의 로그인을 필터링합니다.
    # 원본 주석: "7. 08:00~19:00 이외 접속" - 참고: 상수가 이를 더 정확하게 정의합니다.
    filtered_off_hours = _filter_off_hours(df)
    if not filtered_off_hours.empty:
        save_path_off_hours = os.path.join(
            save_dir,
            f"{LOGIN_CHECK_REPORT_BASE}({LOGIN_REPORT_OFF_HOURS_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_off_hours, save_path_off_hours)
        print(f"Off-hours login results saved to: {save_path_off_hours}")
    else:
        print("No records found for off-hours login check.")

    # 공휴일 또는 주말 로그인을 필터링합니다.
    # 원본 주석: "8. 토, 일, 공휴일 접속"
    filtered_holiday_weekend = _filter_holiday_and_weekend(df)
    if not filtered_holiday_weekend.empty:
        save_path_holiday_weekend = os.path.join(
            save_dir,
            f"{LOGIN_CHECK_REPORT_BASE}({LOGIN_REPORT_HOLIDAY_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_holiday_weekend, save_path_holiday_weekend)
        print(f"Holiday/weekend login results saved to: {save_path_holiday_weekend}")
    else:
        print("No records found for holiday/weekend login check.")


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
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

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


def _filter_off_hours(df: pd.DataFrame) -> pd.DataFrame:
    """
    표준 업무 시간 외에 발생한 로그인 기록을 필터링합니다.
    업무 시간 외는 `LOGIN_OFF_HOURS_START` 및 `LOGIN_OFF_HOURS_END`로 정의됩니다.

    매개변수:
        df (pd.DataFrame): 로그인 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_ACCESS_TIME` 및 `COL_EMPLOYEE_ID_LOGIN`.

    반환 값:
        pd.DataFrame: 업무 시간 외에 발생한 로그인 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 DataFrame입니다.

    예외:
        ValueError: 입력 DataFrame `df`가 None인 경우.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # 접근 시간에서 시간을 추출합니다.
    hours = df_copy[COL_ACCESS_TIME].dt.hour

    # 오전 업무 시간 외 종료 이전의 기록에 대한 마스크를 생성합니다.
    # 또는 저녁 업무 시간 외 시작 이후의 기록에 대한 마스크를 생성합니다.
    mask = (hours < LOGIN_OFF_HOURS_END) | (hours >= LOGIN_OFF_HOURS_START)
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])


def _filter_holiday_and_weekend(df: pd.DataFrame) -> pd.DataFrame:
    """
    대한민국 공휴일 또는 주말(토요일, 일요일)에 발생한 로그인 기록을 필터링합니다.

    매개변수:
        df (pd.DataFrame): 로그인 기록을 포함하는 DataFrame입니다. 예상 열:
                           `COL_ACCESS_TIME` 및 `COL_EMPLOYEE_ID_LOGIN`.

    반환 값:
        pd.DataFrame: 공휴일 또는 주말에 발생한 로그인 기록을 포함하며,
                      직원 ID와 접근 시간으로 정렬된 DataFrame입니다.

    예외:
        ValueError: 입력 DataFrame `df`가 None인 경우.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # 접근 시간에서 고유한 연도를 가져와 holidays 객체를 올바르게 초기화합니다.
    years = df_copy[COL_ACCESS_TIME].dt.year.unique()
    # 해당 연도에 대한 대한민국 공휴일을 초기화합니다.
    # type: ignore # holidays.KR은 유효합니다.
    kr_holidays = holidays.KR(years=years)  # type: ignore [attr-defined]

    # 로그인 날짜가 주말인지 확인합니다 (토요일=5, 일요일=6).
    is_weekend = df_copy[COL_ACCESS_TIME].dt.weekday >= 5
    # 로그인 날짜가 공휴일인지 확인합니다.
    is_holiday = df_copy[COL_ACCESS_TIME].dt.date.isin(kr_holidays)

    mask = is_weekend | is_holiday
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])
