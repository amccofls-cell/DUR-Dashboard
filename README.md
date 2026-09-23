# DUR Dashboard v1.2

Windows 로컬 환경에서 DUR 기준 Excel을 인덱싱하고 여러 품목을 한 번에 조회하는 데스크톱 앱입니다.

## v1.2 UI/UX
- Sage Green + Terracotta 기반의 Warm Ivory 대시보드 테마
- Arial 중심 타이포그래피(한글은 Windows 설치 폰트 fallback)
- 기준파일 8개 일괄 등록 및 자동 분류
- 기준파일별 **마지막 등록 일시** 저장/표시
- 상단에 전체 기준파일 최종 등록 시각 표시
- 여러 품목 일괄 입력 및 약품별 / DUR 종류별 결과
- DUR 카드를 클릭하여 해당 상세정보 확인
- 상세화면에서 원본 시트/원본 파일/기준년월 등 출처성 메타데이터는 표시하지 않음
- 성분코드, 제품코드, 약품코드, 업체명 등 코드/업체 필드는 결과에서 숨김

## 병용금기 최적화
병용금기 급여/비급여 파일은 모든 시트를 확인하되 A, D, F, G, J, M, N, O 열만 인덱싱합니다.
검색 품목이 A측 또는 B측 어느 쪽에 있어도 상대 품목/성분을 반대편에서 표시합니다.

## GitHub Actions로 Windows EXE 만들기
1. 이 ZIP의 **내용물**을 GitHub 저장소 최상위에 업로드합니다.
2. GitHub `Actions` → `Build Windows EXE` → `Run workflow`를 실행합니다.
3. 빌드 완료 후 Artifacts의 `DUR-Dashboard-Windows`를 다운로드합니다.
4. 압축을 풀고 `DUR_Dashboard.exe`를 실행합니다.

사용자 PC에 Python 설치는 필요하지 않습니다.

## 로컬 데이터
등록한 Excel과 검색 인덱스는 사용자 PC의 `%LOCALAPPDATA%\\DUR_Dashboard` 아래에 저장됩니다. Excel 원본을 GitHub에 올릴 필요가 없습니다.
