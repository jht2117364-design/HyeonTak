# ChoHyeonTakV3 업데이트 및 배포 안내

## 기존 파일을 유지하며 확장판 적용

저장소: https://github.com/jht2117364-design/ChoHyeonTakV3  
웹앱: https://jht2117364-design.github.io/ChoHyeonTakV3/

1. 기존 저장소의 **Code → Download ZIP**으로 먼저 백업합니다.
2. 제공된 **`ChoHyeonTakV3_업데이트_코드만.zip`**을 다운로드해서 **압축을 풉니다**.
3. 저장소 최상위에 V4 코드의 아래 파일들을 업로드해 기존 파일을 교체합니다.

   - `index.html`, `style.css`, `app.js`, `worker.js`, `model.py`
   - `README.md`, `VALIDATION.md`, `UPLOAD_GUIDE.md`, `.nojekyll`
   - `assets/dashboard-preview.png`, `assets/network-preview.png`, `assets/mobile-preview.png`

4. **현재 업로드한 Excel 4개는 그대로 두세요.** V4는 다음 두 경로를 자동으로 확인합니다.

   - `./data/io-2023-large.xlsx` (폴더 유지된 신규 저장소)
   - `./io-2023-large.xlsx` (현재 ChoHyeonTakV3 저장소)

5. Commit changes로 저장한 뒤 **Settings → Pages → Source: Deploy from a branch, Branch: main, Folder: /(root)** 설정을 확인합니다.
6. **Actions → Pages build and deployment** 성공을 확인한 다음 `Ctrl + Shift + R`로 새로고침합니다.

> 기존 이름과 같은 파일을 웹 업로드 화면에서 바로 덮어쓸 수 없는 경우, GitHub Desktop 또는 로컬 Git을 이용해 교체 후 Commit/Push하세요. ZIP 자체를 업로드하는 것은 안 됩니다.

## 기능 검증 6가지

통합대분류 Excel을 불러오고 다음을 실행합니다.

- [ ] **파급효과 순위**: 33개 산업 전체 순위가 표시되고 CSV 다운로드 가능
- [ ] **히트맵**: `Ad`/`Am`/`L` 전환, 셀 클릭 및 33×33 행렬 표시
- [ ] **상세 보고서**: F 건설 선택, 205.84억원 생산유발액, 인쇄/PDF 가능
- [ ] **직접·간접 분해**: F 건설 100억원 → 직접 100.00 / 1차 54.68 / 2차 이상 51.16억원
- [ ] **정책 시나리오**: F 건설·P 교육서비스, 총수요 1000억원에서 4개 배분 시나리오 표시
- [ ] **산업연관 네트워크**: F 건설 중심, 공급/사용 연결선 클릭, 계수 표시, CSV 다운로드

기존의 산업별 거래표 조회, 다중 산업 유발효과 시뮬레이션과 산업 간 비교 기능도 추가로 점검하세요.

## 배포 오류가 난다면

- **404 페이지:** Pages의 `main` 및 `/(root)`, 루트의 `index.html`, Actions 배포 성공 여부 확인
- **Excel 다운로드 404:** Excel 4개가 `data/` 폴더 또는 루트 중 한 위치에 있는지 확인
- **Pyodide/CDN 오류:** 네트워크와 확장 프로그램의 CDN 차단 여부 점검
- **이전 화면이 계속 나옴:** `Ctrl + Shift + R`로 강력 새로고침

참고: [GitHub Pages 배포 안내](https://docs.github.com/en/pages/getting-started-with-github-pages/creating-a-github-pages-site) · [한국은행 원본 자료](https://www.bok.or.kr/portal/bbs/B0000501/view.do?menuNo=201264&nttId=10093668)
