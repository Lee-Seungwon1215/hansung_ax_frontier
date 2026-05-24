# Hansung AX Frontier

한성대학교 SW중심대학 사업단 운영위원회 업무 시연을 위한 인트라넷/전자결재/AI 주간보고서 생성 데모입니다.

## 주요 기능

- 운영위원회 전자결재함 시연
- 공지사항, 업무 양식, 담당 연락처 조회
- 결재 건 기반 AI 주간보고서 초안 생성
- 사용자가 AI 초안을 직접 수정
- Word/PPT 저장 및 다운로드
- 한성대학교 규정 PDF와 사업단 참고자료 기반 검색
- Gmail 자료 연동 기반 보고서 생성 보조

## 필요 환경

- Python 3.10 이상
- Node.js 18 이상 권장
- Gemini API Key
- GEMINI_API_KEY=AIzaSyBUq_RUTYrQkfwtz4cu1U4AJaVCR2UTT8A

## 프로젝트 구조

```txt
hsu_ai/
├─ hansung-back/
│  ├─ main.py
│  ├─ app/
│  │  ├─ factory.py
│  │  ├─ legacy.py
│  │  └─ routers/
│  ├─ rules/
│  ├─ templates/
│  ├─ hansung_data.txt
│  ├─ 소중대_data.txt
│  └─ requirements.txt
└─ hansung-front/
   ├─ public/
   ├─ src/
   └─ package.json
```

## 1. 저장소 클론

```bash
git clone https://github.com/Lee-Seungwon1215/hansung_ax_frontier.git
cd hansung_ax_frontier/hsu_ai
```

## 2. 백엔드 설정

```bash
cd hsu_ai/hansung-back
pip install -r requirements.txt
```

`hansung-back/.env` 파일을 만들고 위에 제공한 Gemini API Key를 입력합니다.

```env
GEMINI_API_KEY=AIzaSyBUq_RUTYrQkfwtz4cu1U4AJaVCR2UTT8A
```

백엔드를 실행합니다.

```bash
uvicorn main:app --reload --port 8001
```

백엔드 주소:

```txt
http://localhost:8001
```

## 3. 프론트엔드 설정

새 터미널에서 실행합니다.

```bash
cd hsu_ai/hansung-front
npm install
npm start
```

프론트엔드 주소:

```txt
http://localhost:3000
```

## 4. 실행 시 참고사항

- `hansung_demo.db`는 GitHub에 포함하지 않습니다.
- 백엔드 실행 시 SQLite DB가 자동 생성됩니다.
- `templates/` 폴더에는 Word 양식 파일이 들어 있습니다.
- `rules/` 폴더에는 AI 참고용 한성대학교 규정 PDF가 들어 있습니다.
- Gmail 연동을 사용하려면 별도의 Google OAuth 설정이 필요합니다.
- Gmail 연동 없이도 인트라넷, 전자결재, AI 보고서 기본 기능은 테스트할 수 있습니다.

## 5. 주요 API

```txt
GET  /intranet/dashboard
POST /intranet/approvals
POST /weekly-reports/compose-request
POST /generate
POST /save-word
POST /save-ppt
GET  /generated-files
GET  /generated-files/{file_id}/download
GET  /rules/search
```

## 6. 저장 방식

- 초안/최종본 텍스트는 SQLite DB에 저장됩니다.
- Word/PPT 생성 파일도 DB의 `generated_files` 테이블에 저장됩니다.
- `output_*.docx`, `output_*.pptx` 같은 로컬 결과 파일은 새로 생성하지 않도록 구현되어 있습니다.

## 7. 주의사항

- `.env`, `hansung_demo.db`, `output_*`, `node_modules`, `build`, `__pycache__`는 Git에 올리지 않습니다.
- 백엔드 포트는 프론트 코드에서 `http://localhost:8001`로 연결되어 있습니다.
- 서버를 재시작한 뒤 화면이 비어 보이면 프론트 브라우저를 새로고침 또는 링크를 복사해서 크롬으로 접속합니다.
