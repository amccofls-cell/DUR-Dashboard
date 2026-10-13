# DUR 품목리스트 조회 프로그램 v1.5.1

Windows용 로컬 DUR 품목리스트 조회 프로그램입니다.

## 이번 패키지 변경사항
- 프로그램 제목을 **DUR 품목리스트 조회 프로그램**으로 변경
- 기존 `DUR 약물안전성 조회`, `의약품 DUR 통합 분석 시스템` 문구 제거
- 화면 우측 하단에 **Created by Chaereen Jeong** 표시
- 검색창에 품목명의 일부를 입력하면 별도 팝업 없이 바로 아래 후보 목록 자동 표시
- 후보 선택 후 여러 품목 누적 조회
- 연령금기/비대면진료 매칭 안정화 유지
- 기존 등록파일이 있는 경우 v1.5.1 첫 실행 시 연령금기/비대면진료 인덱스 자동 보정
- SQLite 잠금 대기/재시도 로직 포함
- 병용금기 급여/비급여는 A, D, F, G, J, M, N, O 열만 인덱싱
- 화면에는 성분코드, 제품코드, 약품코드, 업체명, 원본 출처를 표시하지 않음

## GitHub 업로드
ZIP을 풀어 저장소 최상위에 아래 파일/폴더를 업로드합니다.

- `dur_dashboard.py`
- `requirements.txt`
- `.github/workflows/build-windows.yml`
- `README.md`

그 후 GitHub의 `Actions` → `Build Windows EXE` → `Run workflow`를 실행합니다.

PC에 Python을 설치할 필요는 없습니다.
