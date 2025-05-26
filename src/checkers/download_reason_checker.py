"""
This module is responsible for checking personal information download reasons.
It analyzes logs of personal data downloads and flags suspicious activities such as:
- Downloads with overly short or simplistic reasons (e.g., "asdfg", "12345").
- Users downloading an excessive total number of records.
- Users downloading data with unusually high frequency within a short time.
- Downloads occurring outside of standard business hours or on holidays/weekends.

The main function `sayu_checker` (Korean for "reason checker") drives these checks
and saves filtered results into separate Excel files.
"""

import os
import pandas as pd
import holidays

from utils import save_excel_with_autofit, find_and_prepare_excel_file

# Constants for download_reason_checker.py
PERSONAL_INFO_DOWNLOAD_REASON_PREFIX = "개인정보 다운로드 사유 조회_"
DOWNLOAD_REASON_REPORT_BASE = "[붙임4] 개인정보 다운로드 사유"
DOWNLOAD_REASON_INVALID_REASON_SUFFIX = "사유이상"
DOWNLOAD_REASON_HIGH_DOWNLOAD_COUNT_SUFFIX = "100건 초과"
DOWNLOAD_REASON_HIGH_FREQUENCY_SUFFIX = "1시간20건초과"
DOWNLOAD_REASON_OFF_HOURS_SUFFIX = "업무시간외"
COL_ACCESS_TIME = "접근일시"
COL_EMPLOYEE_ID = "교번"
COL_DOWNLOAD_REASON = "다운로드사유"
COL_DOWNLOAD_COUNT = "다운로드데이터수(건)"
DOWNLOAD_COUNT_THRESHOLD = 100
DOWNLOAD_FREQUENCY_THRESHOLD = 20
DOWNLOAD_OFF_HOURS_START = 23
DOWNLOAD_OFF_HOURS_END = 8


def _unique_char_count_below_5(text_input) -> bool:
    """
    Checks if the number of unique characters in a given string is less than or equal to 5.
    This is used to identify potentially suspicious or non-descriptive download reasons.

    Args:
        text_input: The string to check. Typically a download reason.

    Returns:
        bool: True if the number of unique characters is 5 or less, False otherwise.
              Returns False if the input is NaN (Not a Number).
    """
    if pd.isna(text_input):
        return False
    return len(set(str(text_input))) <= 5


def sayu_checker(download_dir: str, save_dir: str, prev_month: str):
    """
    Main function to check personal information download reasons for suspicious patterns.

    Reads the download reason log file and applies several filters:
    1. Invalid/short download reasons.
    2. High total download count by user.
    3. High frequency of downloads by user within an hour.
    4. Downloads during off-hours or holidays/weekends.

    Each set of filtered results is saved to a separate Excel file.

    Args:
        download_dir (str): Directory containing the source download reason Excel file.
        save_dir (str): Directory where the generated report Excel files will be saved.
        prev_month (str): Previous month in 'YYYYMM' format, used for naming output files.

    Raises:
        FileNotFoundError: If the download reason Excel file cannot be found.
    """
    # Find, copy, and read the download reason Excel file.
    df, _ = find_and_prepare_excel_file(
        download_dir,
        PERSONAL_INFO_DOWNLOAD_REASON_PREFIX,
        save_dir,
        DOWNLOAD_REASON_REPORT_BASE,
        prev_month,
    )

    if df is None:
        raise FileNotFoundError(
            f"Download reason Excel file starting with '{PERSONAL_INFO_DOWNLOAD_REASON_PREFIX}' not found in '{download_dir}'."
        )

    # Filter for downloads with suspicious/short reasons.
    # Original comment: "사유 비정상 작성" (Abnormally written reason)
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

    # Filter for users who downloaded a high total number of records.
    # Original comment: "100건 이상 개인정보 다운로드" (Personal info download over 100 records)
    filtered_high_download = _filter_high_download_users(df)
    if not filtered_high_download.empty:
        save_path_high_download = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_HIGH_DOWNLOAD_COUNT_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_high_download, save_path_high_download)
        print(
            f"Results for high download count (>{DOWNLOAD_COUNT_THRESHOLD}) saved to: {save_path_high_download}"
        )
    else:
        print(
            f"No records found for high download count (>{DOWNLOAD_COUNT_THRESHOLD}) check."
        )

    # Filter for users with a high frequency of downloads within an hour.
    # Original comment: "1시간 이내 다운로드 횟수 20건 이상" (Download frequency over 20 times within 1 hour)
    filtered_high_freq = _filter_high_freq_download(df)
    if not filtered_high_freq.empty:
        save_path_high_freq = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_HIGH_FREQUENCY_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_high_freq, save_path_high_freq)
        print(
            f"Results for high download frequency (>{DOWNLOAD_FREQUENCY_THRESHOLD}/hr) saved to: {save_path_high_freq}"
        )
    else:
        print(
            f"No records found for high download frequency (>{DOWNLOAD_FREQUENCY_THRESHOLD}/hr) check."
        )

    # Filter for downloads that occurred during off-hours or on holidays/weekends.
    # Original comment: "업무시간 외 다운로드" (Off-hours download)
    filtered_off_hours_holiday = _filter_off_hour_and_holiday(df)
    if not filtered_off_hours_holiday.empty:
        save_path_off_hours_holiday = os.path.join(
            save_dir,
            f"{DOWNLOAD_REASON_REPORT_BASE}({DOWNLOAD_REASON_OFF_HOURS_SUFFIX})_{prev_month}.xlsx",
        )
        save_excel_with_autofit(filtered_off_hours_holiday, save_path_off_hours_holiday)
        print(
            f"Results for off-hours/holiday downloads saved to: {save_path_off_hours_holiday}"
        )
    else:
        print("No records found for off-hours/holiday download check.")


def _check_download_sayu(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters records where the download reason is considered suspicious (too short/simple).
    Uses the `_unique_char_count_below_5` helper function for the check.

    Args:
        df (pd.DataFrame): DataFrame containing download records. Expected columns:
                           `COL_DOWNLOAD_REASON` (download reason),
                           `COL_EMPLOYEE_ID` (employee ID),
                           `COL_ACCESS_TIME` (access timestamp).

    Returns:
        pd.DataFrame: Filtered DataFrame with records having suspicious download reasons,
                      sorted by employee ID and access time.

    Raises:
        ValueError: If the expected download reason column (`COL_DOWNLOAD_REASON`)
                    is not found at the 5th position (index 4).
    """
    expected_reason_col_index = 4
    if df.columns[expected_reason_col_index] != COL_DOWNLOAD_REASON:
        raise ValueError(
            f"Expected '{COL_DOWNLOAD_REASON}' column at index {expected_reason_col_index}. Found: {df.columns[expected_reason_col_index]}"
        )

    # Apply the filter for unique character count in download reason.
    # Original comment: "5. 고유 문자 개수 5개 이하인 row 필터링" (Filter rows with unique char count <= 5)
    return df[df[COL_DOWNLOAD_REASON].apply(_unique_char_count_below_5)].sort_values(
        [COL_EMPLOYEE_ID, COL_ACCESS_TIME]
    )


def _filter_high_download_users(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters for users whose total number of downloaded records exceeds a defined threshold.

    Args:
        df (pd.DataFrame): DataFrame containing download records. Expected columns:
                           `COL_DOWNLOAD_COUNT` (number of records downloaded),
                           `COL_EMPLOYEE_ID` (employee ID),
                           `COL_ACCESS_TIME` (access timestamp).

    Returns:
        pd.DataFrame: DataFrame containing all download records for users who exceeded the threshold,
                      sorted by employee ID and access time.

    Raises:
        ValueError: If the expected download count column (`COL_DOWNLOAD_COUNT`)
                    is not found at the 6th position (index 5).
    """
    expected_count_col_index = 5
    if df.columns[expected_count_col_index] != COL_DOWNLOAD_COUNT:
        raise ValueError(
            f"Expected '{COL_DOWNLOAD_COUNT}' column at index {expected_count_col_index}. Found: {df.columns[expected_count_col_index]}"
        )

    # Group by employee ID and sum their download counts.
    download_sum_per_user = df.groupby(COL_EMPLOYEE_ID)[COL_DOWNLOAD_COUNT].sum()
    # Identify users who meet or exceed the download count threshold.
    target_users = download_sum_per_user[
        download_sum_per_user >= DOWNLOAD_COUNT_THRESHOLD
    ].index

    # Return all records for these identified users.
    return df[df[COL_EMPLOYEE_ID].isin(target_users)].sort_values(
        [COL_EMPLOYEE_ID, COL_ACCESS_TIME]
    )


def _filter_high_freq_download(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters for users who downloaded data with high frequency (more than a threshold number of times within an hour).

    Args:
        df (pd.DataFrame): DataFrame containing download records. Expected columns:
                           `COL_ACCESS_TIME` (access timestamp),
                           `COL_EMPLOYEE_ID` (employee ID).

    Returns:
        pd.DataFrame: DataFrame containing records that are part of a high-frequency download burst,
                      sorted by employee ID and access time. Returns an empty DataFrame if no such bursts are found.

    Raises:
        ValueError: If the input DataFrame `df` is None.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    flagged_indices = (
        set()
    )  # Store original indices of records belonging to a high-frequency burst.

    # Group by employee ID to analyze each user's download patterns.
    for _, group in df_copy.groupby(COL_EMPLOYEE_ID):
        group = group.sort_values(
            COL_ACCESS_TIME
        ).reset_index()  # Reset index to use .loc with integer indices i.

        for i in range(len(group)):
            current_download_time = group.loc[i, COL_ACCESS_TIME]
            # Define a 1-hour window from the current download time.
            window_end_time = current_download_time + pd.Timedelta(hours=1)

            # Select downloads within this 1-hour window.
            downloads_in_window = group[
                (group[COL_ACCESS_TIME] >= current_download_time)
                & (group[COL_ACCESS_TIME] <= window_end_time)
            ]

            # If the number of downloads in this window meets the frequency threshold, flag them.
            if len(downloads_in_window) >= DOWNLOAD_FREQUENCY_THRESHOLD:
                flagged_indices.update(
                    downloads_in_window["index"].tolist()
                )  # Use original index stored after reset_index()

    if flagged_indices:
        result_df = df_copy.loc[
            sorted(list(flagged_indices))
        ]  # Select using original indices
        return result_df.sort_values([COL_EMPLOYEE_ID, COL_ACCESS_TIME])
    else:
        return pd.DataFrame(
            columns=df.columns
        )  # Return empty DataFrame with same columns


def _filter_off_hour_and_holiday(df: pd.DataFrame) -> pd.DataFrame:
    """
    Filters download records that occurred outside of standard business hours or on South Korean public holidays/weekends.
    Off-hours are defined by `DOWNLOAD_OFF_HOURS_START` and `DOWNLOAD_OFF_HOURS_END`.

    Args:
        df (pd.DataFrame): DataFrame containing download records. Expected columns:
                           `COL_ACCESS_TIME` (access timestamp),
                           `COL_EMPLOYEE_ID` (employee ID).

    Returns:
        pd.DataFrame: DataFrame containing download records from off-hours or holidays/weekends,
                      sorted by employee ID and access time.

    Raises:
        ValueError: If the input DataFrame `df` is None.
    """
    if df is None:
        raise ValueError("Input DataFrame cannot be None.")

    df_copy = df.copy()
    df_copy[COL_ACCESS_TIME] = pd.to_datetime(df_copy[COL_ACCESS_TIME])

    # Initialize South Korean holidays for the relevant years.
    years = df_copy[COL_ACCESS_TIME].dt.year.unique()
    kr_holidays = holidays.KR(years=years)  # type: ignore # holidays.KR is valid

    # Extract temporal features for checks.
    weekday = df_copy[COL_ACCESS_TIME].dt.weekday
    hour = df_copy[COL_ACCESS_TIME].dt.hour
    date_only = df_copy[COL_ACCESS_TIME].dt.date  # For holiday checking

    # Define conditions for off-hours, weekend, and holiday.
    is_off_hour = (hour < DOWNLOAD_OFF_HOURS_END) | (hour >= DOWNLOAD_OFF_HOURS_START)
    is_weekend = weekday >= 5  # Monday is 0 and Sunday is 6; Saturday=5, Sunday=6.
    is_holiday = date_only.isin(kr_holidays)

    # Combine conditions: any record that is off-hour OR weekend OR holiday.
    mask = is_off_hour | is_weekend | is_holiday
    return df_copy[mask].sort_values([COL_EMPLOYEE_ID, COL_ACCESS_TIME])
