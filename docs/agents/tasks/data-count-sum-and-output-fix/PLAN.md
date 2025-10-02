Objective: 붙임2,3,4 파일 데이터 개수 합계 표기 및 출력 스타일 수정

Constraints: TDD 준수, 기존 기능 유지

Target Files & Changes:
- 각 checker.py: 함수가 int 반환 (데이터 개수)
- main.py: 반환값 받아 합계 계산, print_summary에 전달
- display.py: Console(markup=True)

Test/Validation cases: 실행 후 합계 출력 확인, 스타일 제대로 렌더링

Steps (1..N):
1. 각 checker 함수 수정하여 데이터 개수 반환
2. main.py 수정하여 합계 계산
3. display.py Console 설정 변경
4. 테스트 실행

Rollback: git revert

Review Hotspots: 데이터 개수 정확성, 출력 스타일

Status [ ] Step: 연구 완료, 계획 작성 중
