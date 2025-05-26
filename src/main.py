"""
Main script to run the data checking processes.

This script orchestrates the execution of various checker modules:
- Download Reason Checker (`sayu_checker`)
- Login Checker (`login_checker`)
- Personal File Access Checker (`personal_file_checker`)

It sets up necessary configurations such as environment variables for directories
and then calls the main function of each checker.
The script is designed to be run as the primary entry point for the application.
"""

import os
from dotenv import load_dotenv

from utils import get_prev_month_yyyymm, make_save_dir
from checkers.download_reason_checker import sayu_checker
from checkers.login_checker import login_checker
from checkers.personal_file_checker import personal_file_checker

# Load environment variables from a .env file if it exists.
# This is useful for configuring DOWNLOAD_DIR and SAVE_DIR without hardcoding.
load_dotenv()
download_dir = os.getenv(
    "DOWNLOAD_DIR"
)  # Directory where source Excel files are located.
base_save_dir = os.getenv(
    "SAVE_DIR"
)  # Base directory where output reports will be saved.


def main():
    """
    Main function to execute the data checking processes.

    It determines the target month for analysis (previous month),
    creates the necessary save directory structure, and then runs
    the selected checker functions.

    Currently, `personal_file_checker` is active, while `sayu_checker`
    and `login_checker` are commented out. Uncomment them to include them
    in the execution.
    """
    if not base_save_dir:
        print(
            "Error: SAVE_DIR environment variable is not set. Please configure it in your .env file or environment."
        )
        return
    if not download_dir:
        print(
            "Error: DOWNLOAD_DIR environment variable is not set. Please configure it in your .env file or environment."
        )
        return

    prev_month_str = get_prev_month_yyyymm()
    # Create a specific subdirectory for the previous month's reports.
    # e.g., if base_save_dir is /reports, save_dir will be /reports/YYYYMM
    reports_save_dir = make_save_dir(base_save_dir)

    print(f"Starting data checks for month: {prev_month_str}")
    print(f"Source data directory: {download_dir}")
    print(f"Reports will be saved in: {reports_save_dir}")

    # Check if directories are valid before proceeding
    if download_dir and reports_save_dir:
        # Section for checking personal information download reasons.
        # Original comment: "# 개인정보 다운로드 사유 검사" (Personal Information Download Reason Check)
        print("\n### Running Download Reason Checker ###")
        try:
            sayu_checker(download_dir, reports_save_dir, prev_month_str)
            print("Download Reason Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Download Reason Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Download Reason Checker: {e}")

        # Section for checking login IP patterns.
        # Original comment: "# 로그인 IP 검사" (Login IP Check)
        print("\n### Running Login Checker ###")
        try:
            login_checker(download_dir, reports_save_dir, prev_month_str)
            print("Login Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Login Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Login Checker: {e}")

        # Section for checking personal information access records.
        # Original comment: "# 개인정보 조회 기록 점검" (Personal Information View Record Check)
        print("\n### Running Personal File Checker ###")
        try:
            personal_file_checker(download_dir, reports_save_dir, prev_month_str)
            print("Personal File Checker completed.")
        except FileNotFoundError as e:
            print(f"Error in Personal File Checker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred in Personal File Checker: {e}")

        print("\nAll checks finished.")
    else:
        # This case should ideally be caught by the initial checks for base_save_dir and download_dir
        print("Error: Download directory or save directory is not properly configured.")


if __name__ == "__main__":
    # This ensures that main() is called only when the script is executed directly,
    # not when it's imported as a module.
    main()
