import os

import pandas as pd

from utils import save_excel_with_autofit


def personal_file_checker(download_dir: str, save_dir: str, prev_month: str):
    if not download_dir:
        raise EnvironmentError("DOWNLOAD_DIR 환경변수가 .env 파일에 설정되어 있지 않습니다.")

    # 1. 폴더 내 파일 목록에서 조건에 맞는 Excel 파일 찾기
    file_prefix = "개인정보 접속기록 조회_"
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
    dfs = []
    for filename in files:
        file_path = os.path.join(download_dir, filename)
        try:
            df = pd.read_excel(file_path)
            dfs.append(df)
        except Exception as e:
            print(f"파일 읽기 실패: {filename} - {e}")

    if not dfs:
        raise ValueError("병합할 수 있는 유효한 엑셀 파일이 없습니다.")

    # 3. 모든 데이터프레임을 하나로 합침 (행 단위)
    merged_df = pd.concat(dfs, ignore_index=True)

    merged_filename = f"[붙임3] 개인정보 접속기록 조회_{prev_month}.xlsx"

    # 4. 하나의 엑셀로 저장
    # merged_path = os.path.join(save_dir, merged_filename)
    # merged_df.to_excel(merged_path, index=False)
    # print(f"합쳐진 파일 저장: {merged_path}")

    # 5. 분석 대상 확정
    df = merged_df

    # 6. 인사마스터에서 조회한 기록록
    filtered_df = _filter_by_job_master_exclude_detail_id(df)
    if filtered_df is not None and not filtered_df.empty:
        save_path = os.path.join(
            save_dir, f"[붙임3] 개인정보 접속기록 조회(인사마스터)_{prev_month}.xlsx")

        save_excel_with_autofit(filtered_df, save_path)

        print(f"개인정보 접속기록 중 인사마스터 결과를 저장했습니다.")

    # 조회 1000회 이상 교번별 전체 기록 시트별 저장
    save_path = f"[붙임3] 개인정보 접속기록 조회(1000건이상조회)_{prev_month}.xlsx"
    save_path = os.path.join(save_dir, save_path)
    _extract_and_save_by_job(df, save_path, job='조회', threshold=1000)

    # 저장 100회 이상 교번별 전체 기록 시트별 저장
    save_path = f"[붙임3] 개인정보 접속기록 조회(100건이상저장)_{prev_month}.xlsx"
    save_path = os.path.join(save_dir, save_path)
    _extract_and_save_by_job(df, save_path, job='저장', threshold=100)


def _filter_by_job_master_exclude_detail_id(df: pd.DataFrame) -> pd.DataFrame:
    # 1. 컬럼 존재 검사
    for col in ['프로그램명', '교번', '접속일시', '상세내용']:
        if col not in df.columns:
            raise ValueError(f"'{col}' 컬럼이 데이터프레임에 없습니다.")

    # 2. '프로그램명' == '인사마스터' 필터
    filtered = df[df['프로그램명'] == '인사마스터']

    # 3. 교번이 상세내용에 포함되어 있지 않은 것만 남기기
    filtered = filtered[
        ~filtered.apply(lambda row: str(row['교번']) in str(row['상세내용']), axis=1)
    ]

    # 4. 교번, 접속일시 정렬
    filtered = filtered.sort_values(['교번', '접속일시'])
    return filtered


def _extract_and_save_by_job(df: pd.DataFrame, save_path: str,
                             job: str, threshold: int):
    """
    교번 기준으로 수행업무가 job(예: '조회' 또는 '저장')이고, 해당 횟수가 threshold 이상인 교번별로
    해당 교번의 전체 건수를 시트별로 Excel에 저장.
    """
    if '교번' not in df.columns or '성명' not in df.columns or '수행업무' not in df.columns:
        raise ValueError("'교번', '성명', '수행업무' 컬럼이 필요합니다.")
    cond = df['수행업무'] == job
    group = df[cond].groupby('교번')
    # 임계치 이상 교번 추출
    target_ids = group.size()[group.size() >= threshold].index.tolist()

    # ExcelWriter로 시트별 저장
    with pd.ExcelWriter(save_path) as writer:
        for sid in target_ids:
            subdf = df[df['교번'] == sid]
            # 시트명은 '교번_성명' 형식 (중복 회피 위해)
            name = subdf['성명'].iloc[0] if not subdf.empty else ""
            sheet_name = f"{sid}_{name}"[:31]  # 시트명 31자 제한
            subdf.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"총 {len(target_ids)}개 교번을 시트별로 저장했습니다: {save_path}")
