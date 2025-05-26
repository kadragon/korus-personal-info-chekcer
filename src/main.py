import os
from dotenv import load_dotenv

from utils import get_prev_month_yyyymm, make_save_dir
from mode.personal_file_download import sayu_checker
from mode.login_checker import login_checker
from mode.personal_file_checker import personal_file_checker


load_dotenv()
download_dir = os.getenv('DOWNLOAD_DIR')
base_save_dir = os.getenv('SAVE_DIR')


def main():
    prev_month = get_prev_month_yyyymm()
    save_dir = make_save_dir(base_save_dir)

    if download_dir and save_dir:
        print("# 개인정보 다운로드 사유 검사")
        # sayu_checker(
        # download_dir, save_dir, prev_month)

        print("# 로그인 IP 검사")
        # login_checker(download_dir, save_dir, prev_month)

        print("# 개인정보 조회 기록 점검")
        personal_file_checker(download_dir, save_dir, prev_month)


if __name__ == "__main__":
    main()
