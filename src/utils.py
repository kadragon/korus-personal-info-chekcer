import os
from datetime import datetime

import openpyxl
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.utils import get_column_letter


def get_prev_month_yyyymm() -> str:
    today = datetime.today()
    prev_month = today - relativedelta(months=1)
    return prev_month.strftime("%Y%m")


def make_save_dir(base_save_dir) -> str:
    prev_month = get_prev_month_yyyymm()
    save_dir = os.path.join(base_save_dir, prev_month)

    # 디렉터리 존재 여부 판단
    if not os.path.exists(save_dir):
        os.makedirs(save_dir, exist_ok=True)
        print(f"폴더 생성: {save_dir}")

    return save_dir


def save_excel_with_autofit(df: pd.DataFrame, path: str):
    df.to_excel(path, index=False)
    wb = openpyxl.load_workbook(path)
    ws = wb.active

    for idx, column_cells in enumerate(ws.columns):  # type: ignore
        max_length = 0
        for cell in column_cells:
            try:
                value = str(cell.value) if cell.value is not None else ""
                max_length = max(max_length, len(value))
            except Exception:
                pass
        if max_length == 0:
            max_length = 10
        # idx+1은 1-base(즉 A=1, B=2)
        column = get_column_letter(idx + 1)
        ws.column_dimensions[column].width = max_length + 2  # type: ignore

    wb.save(path)
    wb.close()
