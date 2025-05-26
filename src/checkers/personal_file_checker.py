"""
This module is responsible for checking personal information access logs.
It processes Excel files containing records of user access to personal data,
merges them if multiple files are found, and then performs several checks:
- Access to '인사마스터' (HR Master) program, excluding self-access.
- High volume of data views by users.
- High volume of data saves by users.

Filtered results for each check are saved into separate Excel files, with high-volume
access reports potentially containing multiple sheets (one per user exceeding the threshold).
"""

import os
import pandas as pd
from utils import save_excel_with_autofit

# Constants for personal_file_checker.py
PERSONAL_INFO_ACCESS_LOG_PREFIX = (
    "개인정보 접속기록 조회_"  # Prefix for personal information access log files
)
MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE = "[붙임3] 개인정보 접속기록 조회_{}.xlsx"
PERSONAL_INFO_ACCESS_MASTER_SUFFIX = (
    "인사마스터"  # Suffix for report: access to '인사마스터' (HR Master) program
)
PERSONAL_INFO_ACCESS_HIGH_VOLUME_VIEWS_SUFFIX = "1000건이상조회"  # Suffix for report: users viewing more than VIEW_THRESHOLD records
PERSONAL_INFO_ACCESS_HIGH_VOLUME_SAVES_SUFFIX = (
    "100건이상저장"  # Suffix for report: users saving more than SAVE_THRESHOLD records
)
COL_ACCESS_TIME = "접근일시"  # Access Timestamp (e.g., "YYYY-MM-DD HH:MM:SS")
COL_EMPLOYEE_ID = "교번"  # Employee ID. Primarily used in download reason and personal file access logs.
COL_EMPLOYEE_ID_LOGIN = (
    "신분번호"  # Employee ID, specifically found in login history files.
)
COL_PROGRAM_NAME = "프로그램명"  # Name of the program/system accessed by the user
COL_DETAIL_CONTENT = (
    "상세내용"  # Detailed description of the activity or content accessed
)
COL_JOB_PERFORMANCE = (
    "수행업무"  # Job/Task performed by the user (e.g., '조회' - View, '저장' - Save)
)
COL_EMPLOYEE_NAME = "성명"  # Employee Name
VIEW_THRESHOLD = 1000  # For personal_file_checker: flags users viewing more than this number of records.
SAVE_THRESHOLD = 100  # For personal_file_checker: flags users saving more than this number of records.
SHEET_NAME_MAX_CHARS = (
    31  # Maximum number of characters allowed for an Excel sheet name.
)
EXCEL_EXTENSIONS = (
    ".xlsx",
    ".xls",
)  # Tuple of supported Excel file extensions for input files.


def personal_file_checker(download_dir: str, save_dir: str, prev_month: str):
    """
    Main function to check personal information access logs.

    It finds all relevant Excel files in `download_dir`, merges them into a single DataFrame,
    and then applies various filters:
    1. Access to '인사마스터' (HR Master) program, excluding cases where users access their own records.
    2. Users who have viewed an unusually high number of records (above `const.VIEW_THRESHOLD`).
    3. Users who have saved an unusually high number of records (above `const.SAVE_THRESHOLD`).

    Results for each check are saved into separate Excel files. Reports for high-volume
    access are multi-sheet Excel files, with each sheet containing all records for a specific user
    who exceeded the threshold.

    Args:
        download_dir (str): Directory containing the source personal information access log Excel files.
        save_dir (str): Directory where the generated report Excel files will be saved.
        prev_month (str): Previous month in 'YYYYMM' format, used for naming output files.

    Raises:
        EnvironmentError: If `download_dir` is not set in the environment variables.
        FileNotFoundError: If no relevant Excel files are found in `download_dir`.
        ValueError: If no valid Excel files can be merged, or if required columns are missing.
    """
    if not download_dir:
        # This check is based on the original script's structure.
        # Consider making download_dir a direct argument if it's always available.
        raise EnvironmentError("DOWNLOAD_DIR environment variable is not set.")

    # Find all Excel files matching the prefix and extensions.
    files = [
        f
        for f in os.listdir(download_dir)
        if f.startswith(PERSONAL_INFO_ACCESS_LOG_PREFIX)
        and f.lower().endswith(EXCEL_EXTENSIONS)
    ]

    if not files:
        raise FileNotFoundError(
            f"No Excel files starting with '{PERSONAL_INFO_ACCESS_LOG_PREFIX}' found in '{download_dir}'."
        )

    # Read and concatenate all found Excel files.
    dfs = []
    for filename in files:
        file_path = os.path.join(download_dir, filename)
        try:
            df_temp = pd.read_excel(file_path)
            dfs.append(df_temp)
            print(f"Successfully read: {filename}")
        except Exception as e:
            print(f"Failed to read file: {filename} - {e}")

    if not dfs:
        raise ValueError(
            "No valid Excel files could be merged. Please check file contents and format."
        )

    merged_df = pd.concat(dfs, ignore_index=True)
    print(
        f"Successfully merged {len(dfs)} files into a single DataFrame with {len(merged_df)} records."
    )

    # The original script had a commented-out section for saving the initially merged file.
    # This can be useful for debugging or intermediate checks.
    # merged_filename = MERGED_PERSONAL_INFO_ACCESS_FILENAME_TEMPLATE.format(prev_month)
    # merged_save_path = os.path.join(save_dir, merged_filename)
    # save_excel_with_autofit(merged_df, merged_save_path)
    # print(f"Merged personal information access log saved to: {merged_save_path}")

    df_to_analyze = merged_df

    # Filter 1: Access to '인사마스터' (HR Master), excluding self-access.
    # Original comment: "인사마스터에서 조회한 기록 (본인 제외)" (Records viewed in HR Master (excluding self))
    filtered_master_access = _filter_by_job_master_exclude_detail_id(df_to_analyze)
    if not filtered_master_access.empty:
        # Construct filename using the MERGED_...TEMPLATE as a base, then adding the specific suffix.
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
        print(f"HR Master access (excluding self) results saved to: {save_path_master}")
    else:
        print("No records found for HR Master access (excluding self) check.")

    # Filter 2: Users viewing a high volume of records.
    # Original comment: "조회 VIEW_THRESHOLD 이상 교번별 전체 기록 시트별 저장" (Save all records for each employee ID with views >= VIEW_THRESHOLD, by sheet)
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

    # Filter 3: Users saving a high volume of records.
    # Original comment: "저장 SAVE_THRESHOLD 이상 교번별 전체 기록 시트별 저장" (Save all records for each employee ID with saves >= SAVE_THRESHOLD, by sheet)
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
    Filters records for access to '인사마스터' (HR Master) program, excluding instances
    where the user's own ID appears in the '상세내용' (detail content) field (i.e., self-access).

    Args:
        df (pd.DataFrame): DataFrame containing personal information access logs.
                           Expected columns: `COL_PROGRAM_NAME`, `COL_EMPLOYEE_ID` (or `COL_EMPLOYEE_ID_LOGIN`),
                           `COL_ACCESS_TIME`, `COL_DETAIL_CONTENT`.

    Returns:
        pd.DataFrame: A DataFrame containing filtered records, sorted by employee ID and access time.
                      Returns an empty DataFrame if no such records are found.

    Raises:
        ValueError: If essential columns for filtering are missing.
    """
    # Determine which employee ID column to use ('교번' or '신분번호')
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

    # Filter for access to '인사마스터' (HR Master) program. '인사마스터' is a specific program name.
    # Original comment: "2. '프로그램명' == '인사마스터' 필터" (Filter for '프로그램명' == '인사마스터')
    filtered_df = df[df[COL_PROGRAM_NAME] == "인사마스터"]

    # Exclude records where the employee's ID is found within the detail content (self-access).
    # Original comment: "3. 교번이 상세내용에 포함되어 있지 않은 것만 남기기" (Keep only those where employee ID is not in detail content)
    # This lambda function checks if the string representation of the user's ID is part of the detail content string.
    # Ensure both are treated as strings for robust comparison.
    filtered_df = filtered_df[
        ~filtered_df.apply(
            lambda row: str(row[employee_id_col_to_use])
            in str(row[COL_DETAIL_CONTENT]),
            axis=1,
        )
    ]

    # Sort the results.
    # Original comment: "4. 교번, 접속일시 정렬" (Sort by employee ID, access time)
    return filtered_df.sort_values([employee_id_col_to_use, COL_ACCESS_TIME])


def _extract_and_save_by_job(
    df: pd.DataFrame, save_path: str, job: str, threshold: int, job_column_name: str
):
    """
    Identifies users who performed a specific `job` (e.g., '조회' - View, '저장' - Save)
    more than `threshold` times. For each such user, it saves all their records
    (not just those matching the job) into a separate sheet in the specified Excel file.

    Args:
        df (pd.DataFrame): The DataFrame containing all personal information access logs.
        save_path (str): The full path to the Excel file where results will be saved.
        job (str): The specific job type to count (e.g., '조회', '저장').
        threshold (int): The minimum number of times the `job` must be performed to flag a user.
        job_column_name (str): The name of the column in `df` that contains the job type (e.g., `COL_JOB_PERFORMANCE`).

    Raises:
        ValueError: If required columns (`COL_EMPLOYEE_ID` or its alternative, `COL_EMPLOYEE_NAME`, `job_column_name`) are missing.
    """
    # Determine which employee ID column to use ('교번' or '신분번호')
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

    # Filter records matching the specific job type.
    job_specific_df = df[df[job_column_name] == job]
    # Group by employee ID and count occurrences of the job.
    job_counts_per_user = job_specific_df.groupby(employee_id_col_to_use).size()
    # Identify users who meet or exceed the threshold for this job.
    target_user_ids = job_counts_per_user[
        job_counts_per_user >= threshold
    ].index.tolist()

    if not target_user_ids:
        print(
            f"No users found exceeding threshold of {threshold} for job '{job}' in '{save_path}'."
        )
        # Create an empty file or a file with a notice if that's desired.
        # For now, just prints a message. If an empty file is required:
        # with pd.ExcelWriter(save_path) as writer:
        #     pd.DataFrame().to_excel(writer, sheet_name="NoData", index=False)
        return

    # Create an ExcelWriter to save multiple sheets.
    with pd.ExcelWriter(save_path) as writer:
        for employee_id in target_user_ids:
            # Get all records for the identified user (not just the job-specific ones).
            user_all_records_df = df[df[employee_id_col_to_use] == employee_id]

            # Attempt to get the user's name for the sheet name.
            user_name = ""
            if (
                not user_all_records_df.empty
                and COL_EMPLOYEE_NAME in user_all_records_df.columns
            ):
                user_name = user_all_records_df[COL_EMPLOYEE_NAME].iloc[0]

            # Create a sheet name, ensuring it's within Excel's character limit.
            sheet_name = f"{employee_id}_{user_name}"
            if len(sheet_name) > SHEET_NAME_MAX_CHARS:
                # Truncate and ensure uniqueness if needed, though simple truncation is used here.
                sheet_name = sheet_name[:SHEET_NAME_MAX_CHARS]

            user_all_records_df.to_excel(writer, sheet_name=sheet_name, index=False)
            # Note on autofit: The current `save_excel_with_autofit` utility saves a DataFrame to a new file
            # and then autofits. Applying it directly to sheets within an ExcelWriter context
            # would require modification of the utility or post-processing of the saved multi-sheet file.
            # This was also noted in the previous docstring update.

    print(
        f"Saved {len(target_user_ids)} users' data to separate sheets in: {save_path}"
    )
