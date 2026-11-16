# PDF 문서 번역기

PDF를 올리면 내용과 양식을 읽어 **문서 유형(문제집/덤프 · 발주서 · 견적서 · 계약서 · 일반 문서)** 을 판별하고,
유형에 맞는 양식으로 보여 준 뒤 선택한 언어로 번역해 **PDF · Excel · Word · 텍스트 · HTML** 로 내려받는 정적 웹앱입니다.
서버가 없고, 모든 처리는 사용자의 브라우저 안에서 이루어집니다. (원본: document-translator v22 재설계)

## 번역 방식 (키·플러그인 불필요)

| 방식 | 동작 | 조건 |
|---|---|---|
| 브라우저 내장 번역 | Chrome/Edge의 기기 내 Translator API로 자동 번역 | 최신 Chrome 138+ 또는 이를 지원하는 Edge, 데스크톱 |
| Copilot 채팅 번역 | 앱이 묶음별 번역 요청문을 만들고 → Microsoft 365 Copilot 채팅에 붙여 넣기 → 답변(JSON)을 앱에 다시 붙여 넣기 | 회사 계정으로 Copilot 채팅 사용 가능 |

> API 키나 비밀값을 코드에 넣지 않습니다. 공개 저장소에 올려도 안전합니다.

## 주요 기능
- 규칙 기반 문서 유형 자동 판별(판별 근거 표시) + 유형 수동 변경
- 유형별 양식: 문항·보기·정답·해설 / 기본 정보 + 품목 표 / 조항 목록 / 제목·본문
- 페이지 범위 지정, 여러 파일 동시 관리
- 보기 전환: 번역 · 번역+원문 · 원문
- 내보내기: PDF(.pdf), Excel(.xlsx, 시트별 정리), Word(.doc), 텍스트(.txt), HTML(.html)

## 로컬 실행
```bash
npm install
npm run dev        # http://localhost:5173
npm run build      # dist/ 생성
npm run preview    # 빌드 결과 확인
```
Node.js 20 이상이 필요합니다.

## GitHub Pages 배포
1. GitHub에서 새 저장소를 만들고 이 폴더 내용을 그대로 올립니다.
   ```bash
   git init
   git add .
   git commit -m "PDF 문서 번역기"
   git branch -M main
   git remote add origin https://github.com/<계정>/<저장소>.git
   git push -u origin main
   ```
2. 저장소 **Settings → Pages → Build and deployment → Source** 를 **GitHub Actions** 로 바꿉니다.
3. `main`에 push하면 `.github/workflows/deploy.yml` 이 자동으로 빌드·배포합니다.
   주소: `https://<계정>.github.io/<저장소>/`

`vite.config.ts` 의 `base: "./"` 덕분에 저장소 이름과 상관없이 동작합니다.
GitHub Pages 대신 사내 웹 서버, Azure Static Web Apps, SharePoint 등 정적 호스팅에 `dist/` 폴더만 올려도 됩니다.

## 폴더 구조
```
src/
  pages/home.tsx         화면 전체(업로드, 유형, 번역, 내보내기, 양식 렌더링)
  lib/pdf-parse.ts       PDF 텍스트 추출, 머리글/바닥글 제거, 유형 판별, 구조화
  lib/translate.ts       언어 목록, 브라우저 번역, Copilot 요청문 생성·답변 해석
  lib/exporters.ts       TXT / HTML / Word / Excel 생성
  lib/pdf-export.ts      결과 화면 → A4 PDF
  components/ui/         버튼·선택 등 UI 부품(shadcn/ui 기반)
.github/workflows/deploy.yml   GitHub Pages 자동 배포
```

## 유형 판별 규칙 손보기
`src/lib/pdf-parse.ts` 의 `detectKind` 점수 규칙과 상단 정규식(`Q_START`, `CHOICE`, `ANSWER`, `CLAUSE`, `FIELD` 등)을 바꾸면 됩니다.
점수가 4 미만이면 "일반 문서"로 분류합니다.

## 제한 사항
- 스캔 이미지로만 된 PDF(글자 레이어 없음)는 읽지 못합니다(OCR 미포함).
- 브라우저 내장 번역은 브라우저·OS·회사 정책에 따라 꺼져 있을 수 있습니다. 그때는 Copilot 채팅 번역을 쓰세요.
- PDF 내보내기는 화면을 이미지로 담기 때문에 PDF 안의 글자는 선택되지 않습니다. 편집이 필요하면 Word·HTML을 쓰세요.
- 공개 저장소의 GitHub Pages는 누구나 접속할 수 있습니다. 문서는 서버로 전송되지 않지만, 사내 전용으로 쓰려면 비공개 저장소 + 사내 호스팅을 권장합니다.
