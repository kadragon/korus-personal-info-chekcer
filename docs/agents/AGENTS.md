## korus-personal-info-checker

### Intent

이 에이전트는 'KORUS 개인정보 처리 시스템'의 접속 기록 로그를 분석하여 개인정보 오남용 의심 사례를 탐지하고 보고서를 생성하는 작업을 수행합니다. 주요 목표는 다운로드 사유, 로그인 IP, 대량 조회/저장 패턴을 점검하는 것입니다.

### Constraints

- **환경 변수**: 실행 전 반드시 프로젝트 루트에 `.env` 파일을 생성하고 `DOWNLOAD_DIR`(로그 파일 위치)와 `SAVE_DIR`(보고서 저장 위치)를 설정해야 합니다.
- **입력 데이터**: 분석 대상 로그는 특정 형식을 가진 Excel 파일이어야 합니다. 파일명의 접두사 규칙(예: `개인정보 접속기록 조회_`)을 따릅니다.
- **실행 환경**: Python 3.12 및 `pyproject.toml`에 명시된 의존성 라이브러리 설치가 필요합니다.

### Context

#### Project Overview

- **목적**: KORUS 시스템의 Excel 로그를 분석하여 잠재적인 개인정보 오남용 사례(부적절한 접근, 대량 조회/저장 등)를 식별하는 Python 기반 CLI 도구입니다.
- **주요 기능**: 다운로드 사유 점검, 로그인 IP 패턴 분석, 인사마스터 접근 기록 분석.

#### Tech Stack

- **언어**: Python 3.12
- **핵심 라이브러리**: `pandas`, `openpyxl` (Excel 처리), `python-dotenv` (환경 변수 관리)

#### Architecture

- **`src/main.py`**: 메인 진입점. 각 검사 모듈을 순차적으로 호출합니다.
- **`src/checkers/`**: `personal_file_checker.py`, `login_checker.py` 등 핵심 분석 로직을 담은 모듈 디렉토리입니다.
- **`src/utils.py`**: 파일 시스템 처리, 날짜 계산 등 공통 유틸리티 함수를 제공합니다.

#### Setup & Execution

1.  **의존성 설치**:
    ```bash
    pip install .
    ```
2.  **환경 변수 설정**: `.env.example`을 복사하여 `.env` 파일을 만들고, `DOWNLOAD_DIR`와 `SAVE_DIR` 경로를 지정합니다.
3.  **실행**:
    ```bash
    python src/main.py
    ```

#### Development Conventions

- **정적 분석**: `ruff`(린팅/포맷팅), `mypy`(타입 체크), `bandit`(보안 스캔)을 사용합니다.
- **문서화**: 모든 주요 함수와 모듈에는 한국어로 상세한 Docstring이 작성되어 있습니다.
- **설정 관리**: 주요 설정값(파일 경로, 임계값 등)은 각 모듈의 시작 부분에 상수로 정의되어 있습니다.
