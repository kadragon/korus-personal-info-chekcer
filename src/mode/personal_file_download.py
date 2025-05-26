import os
import shutil
import pandas as pd
import holidays

from utils import save_excel_with_autofit


def _unique_char_count_below_5(x) -> bool:
    if pd.isna(x):
        return False
    return len(set(str(x))) <= 5


def sayu_checker(download_dir: str, save_dir: str, prev_month: str):
    if not download_dir:
        raise EnvironmentError("DOWNLOAD_DIR 환경변수가 .env 파일에 설정되어 있지 않습니다.")

    # 1. 폴더 내 파일 목록에서 조건에 맞는 Excel 파일 찾기
    file_prefix = "개인정보 다운로드 사유 조회_"
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
            save_dir, f"[붙임4] 개인정보 다운로드 사유_{prev_month}.xls")
        shutil.copy2(target_file, save_path)  # 파일 복사 (원본 보존, 메타데이터도 복사)

    # 4. pandas로 Excel 파일 읽기
    try:
        df = pd.read_excel(target_file)
    except Exception as e:
        raise RuntimeError(f"엑셀 파일 읽기 실패: {e}")

    # 5. 사유를 비 정상적으로 작성한 대상자자
    filtered_df = _check_download_sayu(df)
    if filtered_df is not None and not filtered_df.empty:
        save_path = os.path.join(
            save_dir, f"[붙임4] 개인정보 다운로드 사유(사유이상)_{prev_month}.xlsx")
        save_excel_with_autofit(filtered_df, save_path)

        print(f"사유를 제대로 입력하지 않은 결과를 저장했습니다.")

    # 6. 100건 이상 개인정보 다운로드
    filtered_df = _filter_high_download_users(df)
    if filtered_df is not None and not filtered_df.empty:
        save_path = os.path.join(
            save_dir, f"[붙임4] 개인정보 다운로드 사유(100건 초과)_{prev_month}.xlsx")
        save_excel_with_autofit(filtered_df, save_path)

        print(f"개인정보 다운로드가 100건 초과한 결과를 저장했습니다.")

    # 7. 1시간 이내에 다룬로드 횟수가 20건 이상인 경우
    filtered_df = _filter_high_freq_download(df)
    if filtered_df is not None and not filtered_df.empty:
        save_path = os.path.join(
            save_dir, f"[붙임4] 개인정보 다운로드 사유(1시간20건초과)_{prev_month}.xlsx")

        save_excel_with_autofit(filtered_df, save_path)

        print(f"개인정보 다운로드가 1시간내에 20건을 초과한 결과를 저장했습니다.")

    # 8. 업무시간 외 다운로드인 경우
    filtered_df = _filter_off_hour_and_holiday(df)
    if filtered_df is not None and not filtered_df.empty:
        save_path = os.path.join(
            save_dir, f"[붙임4] 개인정보 다운로드 사유(업무시간외)_{prev_month}.xlsx")

        save_excel_with_autofit(filtered_df, save_path)

        print(f"개인정보 다운로드가 업무시간외 결과를 저장했습니다.")


def _check_download_sayu(df: pd.DataFrame) -> pd.DataFrame:
    col_name = df.columns[4]
    if col_name != '다운로드사유':
        raise ValueError(f"5번째 컬럼명이 '다운로드사유'가 아닙니다. 실제 컬럼명: {col_name}")

    # 5. 고유 문자 개수 5개 이하인 row 필터링
    return df[df['다운로드사유'].apply(_unique_char_count_below_5)].sort_values(['교번', '접속일시'])


def _filter_high_download_users(df: pd.DataFrame) -> pd.DataFrame:
    # 5번째 컬럼명 확인 (0부터 시작 → index 4)
    col_name = df.columns[5]
    if col_name != '다운로드데이터수(건)':
        raise ValueError(f"6번째 컬럼명이 '다운로드데이터수(건)'가 아닙니다. 실제 컬럼명: {col_name}")

    # '교번' 컬럼이 있다고 가정
    # 교번별 다운로드 건수 합계 계산
    download_sum = df.groupby('교번')['다운로드데이터수(건)'].sum()
    target_users = download_sum[download_sum >= 100].index

    # 해당 교번만 필터링
    return df[df['교번'].isin(target_users)].sort_values(['교번', '접속일시'])


def _filter_high_freq_download(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        raise ValueError("입력된 DataFrame(df)가 None입니다.")

    # 1. 접속일시 datetime 변환 (컬럼명이 '접속일시'라고 가정)
    df = df.copy()
    df['접속일시'] = pd.to_datetime(df['접속일시'])

    result_idx = set()

    # 2. 교번 기준 그룹화
    for sid, group in df.groupby('교번'):
        group = group.sort_values('접속일시').reset_index()
        n = len(group)
        for i in range(n):
            curr_time = pd.to_datetime(group.loc[i, '접속일시'])  # type: ignore
            window = group[
                (group['접속일시'] >= curr_time) &
                (group['접속일시'] <= curr_time + pd.Timedelta(hours=1))
            ]
            if len(window) >= 20:
                result_idx.update(window['index'].tolist())

    if result_idx:
        result_df = df.loc[sorted(result_idx)]
        result_df = result_df.sort_values(['교번', '접속일시'])
        return result_df
    else:
        return pd.DataFrame(columns=df.columns)


def _filter_off_hour_and_holiday(df: pd.DataFrame) -> pd.DataFrame:
    if df is None:
        raise ValueError("입력된 DataFrame(df)가 None입니다.")

    df = df.copy()
    df['접속일시'] = pd.to_datetime(df['접속일시'])

    # 대한민국 공휴일 계산
    years = df['접속일시'].dt.year.unique()
    kr_holidays = holidays.KR(years=years)  # type: ignore

    # 요일, 시 추출
    weekday = df['접속일시'].dt.weekday  # 0=월~6=일
    hour = df['접속일시'].dt.hour
    date_only = df['접속일시'].dt.date

    # 조건: (업무시간(07:00~22:59) & 평일 & 평일/공휴일 아님)이 아니면 필터링
    is_off_hour = (hour < 8) | (hour >= 23)
    is_weekend = weekday >= 5
    is_holiday = date_only.isin(kr_holidays)
    is_off = is_off_hour | is_weekend | is_holiday

    result_df = df[is_off].sort_values(['교번', '접속일시'])

    return result_df
