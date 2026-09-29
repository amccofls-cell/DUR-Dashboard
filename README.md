# DUR Dashboard v1.3

사용자가 승인한 UI 시안과 최대한 동일한 정보 구조로 재설계한 Windows 로컬 DUR 조회 프로그램입니다.

## v1.3 UI 변경
- 상단 브랜드/내비게이션: `홈 / 기준파일 관리 / DUR 조회 / 설정`
- 기준파일 최종 업데이트 일시를 상단 우측에 표시
- compact enterprise dashboard 형태의 흰색 패널 + 얇은 border
- Sage Green을 정상/해당 상태에, Terracotta를 선택/강조 상태에 사용
- NanumGothic 우선, Windows 한글 fallback + 숫자/영문은 시스템 렌더링
- 약품별 결과 왼쪽 목록에 품목 상태와 기준파일 최신 날짜 표시
- 7개 DUR을 compact status card로 표시
- 해당 DUR만 상세 탭으로 노출
- 병용금기 상세는 `상대 품목 / 상대 성분 / 급여구분 / 고시번호 / 고시일자` 테이블 + 선택행 상세 패널 구조
- 원본 파일/시트 등 출처 정보는 결과 화면에 표시하지 않음
- 성분코드/제품코드/약품코드/업체명도 결과 화면에 표시하지 않음
- DUR 종류별 결과는 DUR 필터별 해당 품목 목록으로 구성

## 기존 기능 유지
- 8개 Excel 파일을 한 번에 선택 → 파일명/구조를 이용해 자동 분류
- 등록한 기준파일별 마지막 등록 시각 저장
- 여러 품목 동시 입력 및 조회
- 모든 기준파일의 모든 시트 검색
- 병용금기 급여 + 비급여 동시 판정
- 읽기 실패/미등록 시 `확인불가` 처리
- 로컬 SQLite 인덱스 사용

## 병용금기 대용량 파일 최적화
모든 시트를 확인하되 Excel의 아래 열만 검색 인덱스에 저장합니다.

- A: 성분명A
- D: 제품명A
- F: 급여여부A
- G: 성분명B
- J: 제품명B
- M: 고시번호
- N: 고시일자
- O: 상세정보

검색 품목이 A측 또는 B측 어느 쪽에 있어도 반대편을 상대 품목/성분으로 표시합니다.

## GitHub Actions로 EXE 생성
1. ZIP 압축을 풉니다.
2. **ZIP 내부의 파일/폴더**를 GitHub `DUR-Dashboard` 저장소 최상위에 업로드합니다.
3. `Actions` → `Build Windows EXE` → `Run workflow`를 실행합니다.
4. 완료된 workflow의 `Artifacts`에서 `DUR-Dashboard-Windows`를 다운로드합니다.
5. 압축을 풀고 `DUR_Dashboard.exe`를 실행합니다.

사용자 PC에 Python 설치는 필요하지 않습니다.

## 로컬 저장 위치
기준 Excel 사본, 마지막 등록일 설정, SQLite 검색 인덱스는 `%LOCALAPPDATA%\DUR_Dashboard`에 저장됩니다.
