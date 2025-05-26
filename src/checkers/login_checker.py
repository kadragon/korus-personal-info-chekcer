"""
This module performs checks on user login history data.
It identifies suspicious login patterns such as:
- Logins from multiple IP addresses within a short time window.
- Logins during off-hours.
- Logins on holidays and weekends.

The main function `login_checker` orchestrates these checks and saves the results
into separate Excel files.
"""

import os
import pandas as pd
import holidays

from utils import save_excel_with_autofit, find_and_prepare_excel_file

# Constants for login_checker.py
LOGIN_LOG_FILE_PREFIX = (
    "사용자접속내역_Login내역_"  # Prefix for user login history files
)
LOGIN_CHECK_REPORT_BASE = "[붙임2] 코러스 개인정보처리시스템 접속기록 점검 대장"  # Base name for the login check report
LOGIN_REPORT_IP_SWITCH_SUFFIX = "60분IP"  # Suffix for report: users logging in from multiple IPs within LOGIN_IP_SWITCH_WINDOW_HOURS
LOGIN_REPORT_OFF_HOURS_SUFFIX = (
    "업무시간외"  # Suffix for report: logins outside standard working hours
)
LOGIN_REPORT_HOLIDAY_SUFFIX = (
    "휴일"  # Suffix for report: logins on holidays or weekends
)
COL_IP = "IP"  # IP Address
COL_ACCESS_TIME = "접근일시"  # Access Timestamp (e.g., "YYYY-MM-DD HH:MM:SS")
COL_EMPLOYEE_ID_LOGIN = (
    "신분번호"  # Employee ID, specifically found in login history files.
)
LOGIN_IP_SWITCH_WINDOW_HOURS = 1  # For login_checker: time window in hours to detect logins from multiple IPs for the same user.
LOGIN_IP_SWITCH_MIN_IPS = 3  # For login_checker: minimum number of unique IPs within the window to trigger an IP switch alert.
LOGIN_OFF_HOURS_START = (
    23  # Start hour (inclusive, 24-hour format) for login off-hours (e.g., 11 PM)
)
LOGIN_OFF_HOURS_END = 7  # End hour (exclusive, 24-hour format) for login off-hours (e.g., activity before 7 AM)


def login_checker(download_dir: str, save_dir: str, prev_month: str):
    """
    Main function to perform various checks on login history data.

    It reads the login history Excel file, then applies filters for:
    1. IP address switching: Users logging in from multiple IPs within a defined time window.
    2. Off-hours access: Logins occurring outside of standard business hours.
    3. Holiday/weekend access: Logins occurring on official holidays or weekends.

    Each set of filtered results is saved to a separate Excel file.

    Args:
        download_dir (str): The directory where the source login history Excel file is located.
        save_dir (str): The directory where the generated report Excel files will be saved.
        prev_month (str): The previous month in 'YYYYMM' format, used for naming output files
                          and potentially for selecting the correct input file if not handled by `find_and_prepare_excel_file`.

    Raises:
        FileNotFoundError: If the specified login history Excel file cannot be found.
        ValueError: If the expected 'IP' column is not found at the 10th position (index 9).
    """
    # Use the utility function to find, copy, and read the login history Excel file.
    # The copied file is saved with a standardized name in the save_dir.
    df, _ = find_and_prepare_excel_file(
        download_dir,
        LOGIN_LOG_FILE_PREFIX,
        save_dir,
        LOGIN_CHECK_REPORT_BASE,
        prev_month,
    )

    if df is None:
        # find_and_prepare_excel_file already prints a warning if no file is found.
        # This error is raised to stop execution if the primary data source is missing.
        raise FileNotFoundError(
            f"Login history Excel file starting with '{LOGIN_LOG_FILE_PREFIX}' not found in '{download_dir}'."
        )

    # Validate that the 10th column (index 9) is 'IP'. This is a sanity check based on expected file format.
    # Original comment: "5. 컬럼명 확인" (Check column name)
    expected_ip_col_index = 9
    if df.columns[expected_ip_col_index] != COL_IP:
        raise ValueError(
            f"Expected '{COL_IP}' column at index {expected_ip_col_index}. Found: {df.columns[expected_ip_col_index]}"
        )

    # Filter for users logging in from multiple IPs within a short time.
    # Original comment: "6. 60분 이내에 다른 IP 접속" (Logins from different IPs within 60 minutes)
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

    # Filter for logins outside of standard business hours.
    # Original comment: "7. 08:00~19:00 이외 접속" (Logins outside 08:00-19:00) - Note: Constants define this more precisely.
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

    # Filter for logins on holidays or weekends.
    # Original comment: "8. 토, 일, 공휴일 접속" (Logins on Saturdays, Sundays, or public holidays)
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
    Filters for users who logged in from multiple distinct IP addresses within a defined time window.

    Args:
        df (pd.DataFrame): DataFrame containing login records. Expected columns include
                           `COL_ACCESS_TIME` (access timestamp) and `COL_IP` (IP address),
                           and `COL_EMPLOYEE_ID_LOGIN` (employee identifier).

    Returns:
        pd.DataFrame: A DataFrame containing records of users who triggered the IP switch alert,
                      sorted by employee ID and access time. Returns an empty DataFrame if no such records are found.

    Raises:
        ValueError: If the input DataFrame `df` is None.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    flagged_indices = (
        set()
    )  # Using a set to store indices of flagged rows to avoid duplicates.

    # Group by employee ID to analyze each user's login patterns.
    for _, group in df_copy.groupby(COL_EMPLOYEE_ID_LOGIN):
        group = group.sort_values(COL_ACCESS_TIME)

        # Iterate through each login event for the user.
        for i in range(len(group)):
            current_login_time = group.iloc[i][COL_ACCESS_TIME]
            # Define the time window for checking subsequent logins.
            window_end_time = current_login_time + pd.Timedelta(
                hours=LOGIN_IP_SWITCH_WINDOW_HOURS
            )

            # Select logins within this window.
            logins_in_window = group[
                (group[COL_ACCESS_TIME] >= current_login_time)
                & (group[COL_ACCESS_TIME] <= window_end_time)
            ]

            # Check if the number of unique IPs in this window meets the threshold.
            if len(set(logins_in_window[COL_IP])) >= LOGIN_IP_SWITCH_MIN_IPS:
                flagged_indices.update(
                    logins_in_window.index
                )  # Add all records in this window

    if flagged_indices:
        result_df = df_copy.loc[sorted(list(flagged_indices))]
        return result_df.sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])
    else:
        return pd.DataFrame(
            columns=df.columns
        )  # Return empty DataFrame with same columns if no matches


def _filter_off_hours(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters login records that occurred outside of standard business hours.
    Off-hours are defined by `LOGIN_OFF_HOURS_START` and `LOGIN_OFF_HOURS_END`.

    Args:
        df (pd.DataFrame): DataFrame containing login records. Expected columns:
                           `COL_ACCESS_TIME` and `COL_EMPLOYEE_ID_LOGIN`.

    Returns:
        pd.DataFrame: A DataFrame containing login records that occurred during off-hours,
                      sorted by employee ID and access time.

    Raises:
        ValueError: If the input DataFrame `df` is None.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # Extract the hour from the access time.
    hours = df_copy[COL_ACCESS_TIME].dt.hour

    # Create a mask for records that are before the end of off-hours in the morning
    # OR at or after the start of off-hours in the evening.
    mask = (hours < LOGIN_OFF_HOURS_END) | (hours >= LOGIN_OFF_HOURS_START)
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])


def _filter_holiday_and_weekend(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters login records that occurred on South Korean public holidays or weekends (Saturday, Sunday).

    Args:
        df (pd.DataFrame): DataFrame containing login records. Expected columns:
                           `COL_ACCESS_TIME` and `COL_EMPLOYEE_ID_LOGIN`.

    Returns:
        pd.DataFrame: A DataFrame containing login records that occurred on holidays or weekends,
                      sorted by employee ID and access time.

    Raises:
        ValueError: If the input DataFrame `df` is None.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # Get unique years from the access times to initialize the holidays object correctly.
    years = df_copy[COL_ACCESS_TIME].dt.year.unique()
    # Initialize South Korean holidays for the relevant years.
    kr_holidays = holidays.KR(years=years)  # type: ignore # holidays.KR is valid

    # Check if the login date is a weekend (Saturday=5, Sunday=6).
    is_weekend = df_copy[COL_ACCESS_TIME].dt.weekday >= 5
    # Check if the login date is a public holiday.
    is_holiday = df_copy[COL_ACCESS_TIME].dt.date.isin(kr_holidays)

    mask = is_weekend | is_holiday
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID_LOGIN, COL_ACCESS_TIME])
