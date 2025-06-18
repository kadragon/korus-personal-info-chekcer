"""
이 모듈은 날짜 조작, 디렉토리 생성, 엑셀 파일 처리(열 너비 자동 맞춤 저장),
처리할 특정 엑셀 파일 검색 및 준비와 같은 일반적인 작업을 위한 유틸리티 함수를 제공합니다.
"""

import os
from datetime import datetime
import zipfile

import openpyxl
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter

# Constants for utils.py
EXCEL_EXTENSIONS = (
    ".xlsx",
    ".xls",
)  # 입력 파일에 지원되는 Excel 파일 확장자 튜플입니다.


def get_prev_month_yyyymm() -> str:
    """
    현재 날짜로부터 이전 달을 계산하고 'YYYYMM' 형식의 문자열로 반환합니다.

    반환 값:
        str: 'YYYYMM' 형식의 이전 달입니다.
    """
    today = datetime.today()
    prev_month_date = today - relativedelta(months=1)
    return prev_month_date.strftime("%Y%m")


def make_save_dir(base_save_dir: str) -> str:
    """
    `base_save_dir` 내에 이전 달(YYYYMM)의 이름으로 하위 디렉토리를 생성합니다.
    하위 디렉토리가 이미 존재하면 아무 작업도 수행하지 않습니다.

    매개변수:
        base_save_dir (str): 새 하위 디렉토리가 생성될 기본 디렉토리입니다.
                             이 경로는 기존 디렉토리여야 합니다.

    반환 값:
        str: 생성되었거나 이미 존재하는 이전 달의 하위 디렉토리에 대한 전체 경로입니다.
    """
    prev_month_str = get_prev_month_yyyymm()
    save_dir = os.path.join(base_save_dir, prev_month_str)

    # 디렉토리가 존재하는지 확인하고, 없으면 생성합니다.
    if not os.path.exists(save_dir):
        os.makedirs(save_dir, exist_ok=True)
        print(f"Created directory: {save_dir}")

    return save_dir


def save_excel_with_autofit(df: pd.DataFrame, path: str):
    """
    Pandas DataFrame을 Excel 파일로 저장하고 열 너비를 자동으로 맞춥니다.

    매개변수:
        df (pd.DataFrame): 저장할 DataFrame입니다.
        path (str): Excel 파일이 저장될 전체 경로(파일 이름 포함)입니다.
    """
    df.to_excel(path, index=False)
    # 워크북을 로드하고 활성 시트를 선택하여 열 너비를 조정합니다.
    wb = openpyxl.load_workbook(path)
    ws = wb.active  # 활성 워크시트를 가져옵니다.

    # 자동 맞춤을 위한 최대 길이를 계산하기 위해 열을 반복합니다.
    # type: ignore # openpyxl.worksheet.worksheet.Worksheet.columns는 제너레이터입니다.
    for idx, column_cells in enumerate(ws.columns):  # type: ignore
        max_length = 0
        column_letter = get_column_letter(idx + 1)

        for cell in column_cells:
            try:
                if cell.value is not None:
                    # 셀 값 문자열 표현의 길이를 계산합니다.
                    cell_value_str = str(cell.value)
                    max_length = max(max_length, len(cell_value_str))
            except Exception as e:
                print(f"[열 너비 자동 맞춤] {cell.coordinate}에서 예외 발생: {e}")

        # 내용이 없거나 매우 짧은 경우 기본 최소 너비를 설정합니다.
        adjusted_width = max_length + 2 if max_length > 0 else 10
        # type: ignore
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(path)
    wb.close()


def find_and_prepare_excel_file(
    download_dir: str,
    file_prefix: str,
    save_dir: str,
    output_file_basename: str,
    prev_month: str,
) -> tuple[pd.DataFrame | None, str | None]:
    if not download_dir:
        raise EnvironmentError(
            "Download directory ('download_dir') is not specified.")

    excel_files = [
        f
        for f in os.listdir(download_dir)
        if f.startswith(file_prefix) and f.lower().endswith(('.xls', '.xlsx'))
    ]

    if not excel_files:
        print(
            f"Warning: No Excel file starting with '{file_prefix}' found in '{download_dir}'.")
        return None, None

    os.makedirs(save_dir, exist_ok=True)

    all_dfs = []
    for file_name in excel_files:
        file_path = os.path.join(download_dir, file_name)
        # 파일 확장자에 따라 변환 또는 바로 읽기
        if file_path.lower().endswith('.xlsx'):
            df = pd.read_excel(file_path)
        else:
            df = pd.read_excel(file_path, engine="xlrd")
        all_dfs.append(df)

    # 모든 DataFrame을 합치기 (행 단위로)
    merged_df = pd.concat(all_dfs, ignore_index=True)

    print(f"{file_prefix}: {len(merged_df)}건")

    destination_save_path = os.path.join(
        save_dir, f"{output_file_basename}_{prev_month}.xlsx"
    )
    merged_df.to_excel(destination_save_path, index=False)
    print(f"모든 파일을 합쳐서 '{destination_save_path}'로 저장했습니다.")

    return merged_df, destination_save_path


def zip_files_by_prefix(target_dir: str, prefix_list: list[str]):
    """
    파일명에서 '_' 앞부분(붙임N ... )을 zip 이름으로 하여 압축 생성
    """
    files = [f for f in os.listdir(target_dir) if f.endswith('.xlsx')]

    # 접두사별로 그룹핑
    for prefix in prefix_list:
        matched = [f for f in files if f.startswith(prefix)]
        if not matched:
            print(f"⚠️ {prefix}로 시작하는 파일 없음")
            continue

        # zip 파일명은 첫 파일의 '_' 앞까지
        group_name = matched[0].split('_')[0].split(
            '(')[0]  # 예: [붙임3] 개인정보 접속기록 조회
        zip_name = f"{group_name}.zip"
        zip_path = os.path.join(target_dir, zip_name)

        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for filename in matched:
                zipf.write(os.path.join(target_dir, filename),
                           arcname=filename)

        print(f"✅ {zip_name} 생성 ({len(matched)}개 파일 포함)")
