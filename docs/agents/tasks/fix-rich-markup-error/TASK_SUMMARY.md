# Rich Markup 오류 수정

**목표**: `src/display.py`의 `print_summary` 함수에서 `Text(markup=True)` 파라미터 오류를 수정하여 요약 출력이 정상 작동하도록 함.

**주요 변경사항**:
- `pyproject.toml`에 `rich>=13.0.0` 의존성 추가.
- `uv add rich`로 라이브러리 설치 (버전 14.1.0).

**테스트 및 검증**:
- Rich 라이브러리 설치 후 `Text(markup=True)` 파라미터가 정상 작동 확인.
- 오류 재현 시도: 이전에는 `TypeError: Text.__init__() got an unexpected keyword argument 'markup'` 발생, 수정 후 해결.

**커밋**: [Structural] Rich 라이브러리 추가 및 markup 오류 수정
