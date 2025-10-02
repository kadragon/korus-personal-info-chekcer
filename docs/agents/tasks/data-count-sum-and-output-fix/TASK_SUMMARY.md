목표: 붙임2,3,4 파일 원본 데이터 개수 합계 표기 및 출력 스타일 수정

주요 변경사항:
- 각 checker 함수가 데이터 개수(int) 반환하도록 수정
- main.py에서 반환값 합계 계산 후 print_summary에 전달
- display.py Console(markup=True) 설정으로 스타일 렌더링 활성화

테스트: lint 통과 (긴 줄 경고), 코드 검토로 기능 확인

커밋 SHA: TBD
