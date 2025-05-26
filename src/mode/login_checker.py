from typing import List
import os
import shutil

import pandas as pd
import holidays

from utils import save_excel_with_autofit


def login_checker(download_dir: str, save_dir: str, prev_month: str):
    if not download_dir:
        raise EnvironmentError("DOWNLOAD_DIR 환경변수가 .env 파일에 설정되어 있지 않습니다.")

    # 1. 폴더 내 파일 목록에서 조건에 맞는 Excel 파일 찾기
    file_prefix = "사용자접속내역_Login내역_"
    excel_extensions = ('.xlsx', '.xls')
    files = [
        f for f in os.listdir(download_dir)
        if f.startswith(file_prefix) and f.endswith(excel_extensions)
    ]

    # 2. 엑셀 파일이 존재하는지 확인
    if not files:
        raise FileNotFoundError(
            f"{download_dir} 폴더에 '{file_prefix}'로 시작하는 엑셀 파일이 없습니다.")

    # 3. 엑셀 파일 선택
    target_file = os.path.join(download_dir, files[0])  # 여러개면 리스트로 처리 가능

    # 3-1. 엑셀파일 저장해두기
    if target_file and save_dir:
        os.makedirs(save_dir, exist_ok=True)
        save_path = os.path.join(
            save_dir, f"[붙임2] 코러스 개인정보처리시스템 접속기록 점검 대장_{prev_month}.xls")
        shutil.copy2(target_file, save_path)  # 파일 복사 (원본 보존, 메타데이터도 복사)

    # 4. pandas로 Excel 파일 읽기
    try:
        df = pd.read_excel(target_file)
    except Exception as e:
        raise RuntimeError(f"엑셀 파일 읽기 실패: {e}")

    # 5. 컬럼명 확인
    col_name = df.columns[9]
    if col_name != 'IP':
        raise ValueError(f"10번째 컬럼명이 'IP'가 아닙니다. 실제 컬럼명: {col_name}")

    # 6. 60분 이내에 다른 IP 접속
    filtered = _filter_ip_switch(df)
    save_path = os.path.join(
        save_dir, f"[붙임2] 코러스 개인정보처리시스템 접속기록 점검 대장(60분IP)_{prev_month}.xlsx")

    save_excel_with_autofit(filtered, save_path)

    print(f"60분 이내 다른 IP 접속 결과를 저장했습니다.")

    # 7. 08:00~19:00 이외 접속
    filtered = _filter_off_hours(df)
    save_path = os.path.join(
        save_dir, f"[붙임2] 코러스 개인정보처리시스템 접속기록 점검 대장(업무시간외)_{prev_month}.xlsx")

    save_excel_with_autofit(filtered, save_path)

    print(f"업무시간외 접속 결과를 저장했습니다.")

    # 8. 토, 일, 공휴일 접속
    filtered = _filter_holiday_and_weekend(df)

    save_path = os.path.join(
        save_dir, f"[붙임2] 코러스 개인정보처리시스템 접속기록 점검 대장(휴일)_{prev_month}.xlsx")

    save_excel_with_autofit(filtered, save_path)

    print(f"휴일 접속 결과를 저장했습니다.")


def _filter_ip_switch(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        raise ValueError("입력된 DataFrame(df)가 None입니다.")

    df = df.copy()
    df['접근일시'] = pd.to_datetime(df['접근일시'])

    result_idx = set()

    for sid, group in df.groupby('신분번호'):
        group = group.sort_values('접근일시')
        n = len(group)
        for i in range(n):
            # 1시간 이내 window 만들기
            curr_time = group.iloc[i]['접근일시']
            window = group[(group['접근일시'] >= curr_time) &
                           (group['접근일시'] <= curr_time + pd.Timedelta(seconds=3600))]
            unique_ips = set(window['IP'])
            if len(unique_ips) >= 3:
                result_idx.update(window.index)

    if result_idx:
        result_df = df.loc[sorted(result_idx)]
        # 신분번호, 접근일시 기준 정렬
        result_df = result_df.sort_values(['신분번호', '접근일시'])
        return result_df
    else:
        return pd.DataFrame(columns=df.columns)


def _filter_off_hours(df: pd.DataFrame) -> pd.DataFrame:
    """
    '접근일시'가 07:00 이전 또는 23:00 이후인 row만 추출
    """
    if df is None:
        raise ValueError("입력된 DataFrame(df)가 None입니다.")

    df = df.copy()
    df['접근일시'] = pd.to_datetime(df['접근일시'])

    # 시(hour) 추출
    hours = df['접근일시'].dt.hour

    # 08~19시(08:00 <= x < 19:00)는 False, 나머지는 True
    mask = (hours < 7) | (hours >= 23)
    return df[mask].sort_values(['신분번호', '접근일시'])


def _filter_holiday_and_weekend(df: pd.DataFrame) -> pd.DataFrame:
    """
    대한민국 기준, 공휴일/토요일/일요일 접속 이력만 추출
    """
    if df is None:
        raise ValueError("입력된 DataFrame(df)가 None입니다.")

    df = df.copy()
    df['접근일시'] = pd.to_datetime(df['접근일시'])

    # 공휴일 객체 생성 (대한민국)
    years = df['접근일시'].dt.year.unique()
    kr_holidays = holidays.KR(years=years)  # type: ignore

    # 조건: 주말 또는 공휴일
    is_weekend = df['접근일시'].dt.weekday >= 5  # 5:토, 6:일
    is_holiday = df['접근일시'].dt.date.isin(kr_holidays)
    mask = is_weekend | is_holiday

    filtered = df[mask].sort_values(['신분번호', '접근일시'])
    return filtered
