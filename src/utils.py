"""
이 모듈은 날짜 조작, 디렉토리 생성, 엑셀 파일 처리(열 너비 자동 맞춤 저장),
처리할 특정 엑셀 파일 검색 및 준비와 같은 일반적인 작업을 위한 유틸리티 함수를 제공합니다.
"""

import os
from datetime import datetime
import shutil

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
    for idx, column_cells in enumerate(ws.columns):  # type: ignore # openpyxl.worksheet.worksheet.Worksheet.columns는 제너레이터입니다.
        max_length = 0
        column_letter = get_column_letter(idx + 1)

        for cell in column_cells:
            try:
                if cell.value is not None:
                    # 셀 값 문자열 표현의 길이를 계산합니다.
                    cell_value_str = str(cell.value)
                    max_length = max(max_length, len(cell_value_str))
            except Exception:
                # 셀 값을 처리할 수 없는 경우 건너뜁니다.
                pass

        # 내용이 없거나 매우 짧은 경우 기본 최소 너비를 설정합니다.
        adjusted_width = max_length + 2 if max_length > 0 else 10
        ws.column_dimensions[column_letter].width = adjusted_width  # type: ignore # openpyxl.worksheet.dimensions.ColumnDimension.width는 float를 예상합니다.

    wb.save(path)
    wb.close()


def find_and_prepare_excel_file(
    download_dir: str,
    file_prefix: str,
    save_dir: str,
    output_file_basename: str,
    prev_month: str,
) -> tuple[pd.DataFrame | None, str | None]:
    """
    `download_dir`에서 `file_prefix`로 시작하는 첫 번째 Excel 파일을 찾아,
    `output_file_basename`과 `prev_month`를 사용하여 `save_dir` 내의 구조화된 경로로 복사한 후,
    이 복사된 파일을 Pandas DataFrame으로 읽어들입니다.

    복사된 파일은 원본 파일의 확장자('.xls' 또는 '.xlsx')에 관계없이 항상 '.xlsx' 확장자를 갖습니다.

    매개변수:
        download_dir (str): 원본 Excel 파일을 검색할 디렉토리입니다.
        file_prefix (str): 원본 Excel 파일이 가질 것으로 예상되는 접두사입니다 (예: "LoginHistory_").
        save_dir (str): 찾은 Excel 파일을 복사할 기본 디렉토리입니다.
                        `save_dir`가 `make_save_dir`의 결과인 경우 `prev_month`에 대한 하위 디렉토리가 암시적으로 처리될 수 있으며,
                        그렇지 않은 경우 이 함수는 `save_dir`가 존재하지 않으면 생성합니다.
        output_file_basename (str): 복사된 Excel 파일의 기본 이름입니다 (예: "LoginReport").
                                    최종 이름은 "LoginReport_YYYYMM.xlsx"와 같습니다.
        prev_month (str): 복사된 파일의 이름을 지정하는 데 사용되는 'YYYYMM' 형식의 이전 달입니다.

    반환 값:
        tuple[pd.DataFrame | None, str | None]: 다음을 포함하는 튜플입니다:
            - 복사된 Excel 파일에서 읽은 DataFrame. 파일을 찾지 못했거나 오류가 발생한 경우 None입니다.
            - 저장된(복사된) Excel 파일의 전체 경로. 파일을 찾지 못한 경우 None입니다.

    예외:
        EnvironmentError: `download_dir`가 지정되지 않은 경우(비어 있거나 None인 경우).
        RuntimeError: Excel 파일을 읽는 중 오류가 발생한 경우.
    """
    if not download_dir:
        raise EnvironmentError("Download directory ('download_dir') is not specified.")

    # 지정된 접두사로 시작하는 Excel 파일(.xlsx 및 .xls 모두)을 검색합니다.
    excel_files = [
        f
        for f in os.listdir(download_dir)
        if f.startswith(file_prefix) and f.lower().endswith(EXCEL_EXTENSIONS)
    ]

    if not excel_files:
        print(
            f"Warning: No Excel file starting with '{file_prefix}' found in '{download_dir}'."
        )
        return None, None

    # 처음 발견된 파일을 선택합니다.
    source_file_path = os.path.join(download_dir, excel_files[0])

    # 복사된 파일의 경로를 정의하고 .xlsx 확장자로 표준화합니다.
    # 예: /path/to/save_dir/OutputBaseName_YYYYMM.xlsx
    destination_save_path = os.path.join(
        save_dir, f"{output_file_basename}_{prev_month}.xlsx"
    )

    # 복사하기 전에 저장 디렉토리가 있는지 확인합니다.
    os.makedirs(save_dir, exist_ok=True)
    shutil.copy2(source_file_path, destination_save_path)
    print(f"Copied '{source_file_path}' to '{destination_save_path}'")

    try:
        # *복사된* 파일을 DataFrame으로 읽어들입니다.
        df = pd.read_excel(destination_save_path)
    except Exception as e:
        # 일반적으로 특정 예외를 잡는 것이 좋지만, 단순화를 위해:
        raise RuntimeError(
            f"Error reading the Excel file '{destination_save_path}': {e}"
        )

    return df, destination_save_path
