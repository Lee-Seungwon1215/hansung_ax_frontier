from datetime import datetime
from io import BytesIO
from typing import List, Optional
import base64
import os

import chromadb
from docx import Document
from dotenv import load_dotenv
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from google import genai
from pptx import Presentation
from pydantic import BaseModel, Field

try:
    from pypdf import PdfReader
except ImportError:
    try:
        from PyPDF2 import PdfReader
    except ImportError:
        PdfReader = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

load_dotenv()

gemini = genai.Client(api_key=os.getenv("GEMINI_API_KEY"))

app = FastAPI(title="Hansung AI Committee Report API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:3000"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
chroma_client = chromadb.Client()


def load_data(collection_name: str, filename: str):
    collection = chroma_client.get_or_create_collection(name=collection_name)
    filepath = os.path.join(BASE_DIR, filename)
    if not os.path.exists(filepath):
        return collection

    with open(filepath, "r", encoding="utf-8") as file:
        lines = [line.strip() for line in file.readlines() if line.strip()]

    if lines and len(collection.get()["ids"]) == 0:
        collection.add(documents=lines, ids=[f"doc_{index}" for index in range(len(lines))])

    return collection


collections = {
    "committee": load_data("committee", "hansung_data.txt"),
    "sojungdae": load_data("sojungdae", "소중대_data.txt"),
    "ai_university": load_data("ai-university", "aiuniv_data.txt"),
    "job_plus": load_data("job-plus", "jobplus_data.txt"),
}


class AttachmentSource(BaseModel):
    name: str
    mimeType: Optional[str] = None
    type: Optional[str] = None
    text: Optional[str] = None
    data: Optional[str] = None


class EmailSource(BaseModel):
    subject: str
    from_: Optional[str] = Field(default=None, alias="from")
    date: Optional[str] = None
    body: Optional[str] = None
    attachments: List[AttachmentSource] = Field(default_factory=list)


class UploadedFile(BaseModel):
    name: str
    type: Optional[str] = None
    text: Optional[str] = None
    data: Optional[str] = None


class WorkRequest(BaseModel):
    work_content: str = ""
    target_committee: str = "사업단운영위원회"
    meeting_focus: str = "성과 및 추진현황"
    email_sources: List[EmailSource] = Field(default_factory=list)
    uploaded_files: List[UploadedFile] = Field(default_factory=list)


class SaveRequest(BaseModel):
    content: str
    filename: str


def limit_text(text: Optional[str], max_length: int = 12000) -> str:
    if not text:
        return ""
    return text[:max_length]


def search_data(query: str) -> str:
    collection = collections["committee"]
    if not query.strip():
        return "직접 입력된 업무 내용이 없어 기본 사업단 자료를 참고합니다."

    results = collection.query(query_texts=[query], n_results=4)
    if results["documents"] and results["documents"][0]:
        return "\n".join(results["documents"][0])
    return "관련 참고자료 없음"


def decode_file_data(data: str) -> bytes:
    normalized = data.strip()
    if "," in normalized:
        normalized = normalized.split(",", 1)[1]
    padding = "=" * (-len(normalized) % 4)
    return base64.urlsafe_b64decode(normalized + padding)


def extract_pdf_text(file_bytes: bytes) -> str:
    if PdfReader is None:
        return "(PDF 본문 추출 실패: pypdf 또는 PyPDF2가 설치되어 있지 않습니다.)"

    reader = PdfReader(BytesIO(file_bytes))
    pages = [page.extract_text() or "" for page in reader.pages]
    return "\n".join(pages).strip() or "(PDF에서 추출된 텍스트가 없습니다.)"


def extract_docx_text(file_bytes: bytes) -> str:
    doc = Document(BytesIO(file_bytes))
    paragraphs = [paragraph.text for paragraph in doc.paragraphs if paragraph.text.strip()]
    table_rows = []
    for table in doc.tables:
        for row in table.rows:
            values = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if values:
                table_rows.append(" | ".join(values))
    return "\n".join(paragraphs + table_rows).strip() or "(Word 문서에서 추출된 텍스트가 없습니다.)"


def extract_xlsx_text(file_bytes: bytes) -> str:
    if openpyxl is None:
        return "(Excel 본문 추출 실패: openpyxl이 설치되어 있지 않습니다.)"

    workbook = openpyxl.load_workbook(BytesIO(file_bytes), read_only=True, data_only=True)
    blocks = []
    for sheet in workbook.worksheets:
        rows = []
        for row in sheet.iter_rows(values_only=True):
            values = [str(value) for value in row if value is not None and str(value).strip()]
            if values:
                rows.append(" | ".join(values))
        if rows:
            blocks.append(f"[{sheet.title}]\n" + "\n".join(rows))
    workbook.close()
    return "\n\n".join(blocks).strip() or "(Excel 문서에서 추출된 텍스트가 없습니다.)"


def extract_pptx_text(file_bytes: bytes) -> str:
    prs = Presentation(BytesIO(file_bytes))
    slides = []
    for index, slide in enumerate(prs.slides, start=1):
        texts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                texts.append(shape.text.strip())
        if texts:
            slides.append(f"[슬라이드 {index}]\n" + "\n".join(texts))
    return "\n\n".join(slides).strip() or "(PPT 문서에서 추출된 텍스트가 없습니다.)"


def extract_document_text(name: str, mime_type: Optional[str], data: Optional[str]) -> str:
    if not data:
        return ""

    try:
        file_bytes = decode_file_data(data)
        lower_name = name.lower()
        mime = mime_type or ""

        if lower_name.endswith(".pdf") or mime == "application/pdf":
            return extract_pdf_text(file_bytes)
        if lower_name.endswith(".docx") or "wordprocessingml" in mime:
            return extract_docx_text(file_bytes)
        if lower_name.endswith(".xlsx") or "spreadsheetml" in mime:
            return extract_xlsx_text(file_bytes)
        if lower_name.endswith(".pptx") or "presentationml" in mime:
            return extract_pptx_text(file_bytes)

        return "(지원하지 않는 문서 형식입니다. 파일명과 형식만 참고합니다.)"
    except Exception as exc:
        return f"(문서 본문 추출 실패: {exc})"


def source_text(name: str, mime_type: Optional[str], text: Optional[str], data: Optional[str]) -> str:
    if text:
        return text
    return extract_document_text(name, mime_type, data)


def format_email_sources(email_sources: List[EmailSource]) -> str:
    if not email_sources:
        return "Gmail 참고자료 없음"

    blocks = []
    for index, email in enumerate(email_sources, start=1):
        attachments = "\n".join(
            [
                f"  - {file.name} ({file.mimeType or file.type or 'unknown'}): "
                f"{limit_text(source_text(file.name, file.mimeType or file.type, file.text, file.data), 5000)}"
                for file in email.attachments
            ]
        )
        blocks.append(
            f"""[메일 {index}]
제목: {email.subject}
발신자: {email.from_ or ''}
일자: {email.date or ''}
본문:
{limit_text(email.body)}
첨부:
{attachments or '첨부 없음'}"""
        )
    return "\n\n".join(blocks)


def format_uploaded_files(uploaded_files: List[UploadedFile]) -> str:
    if not uploaded_files:
        return "사용자 첨부자료 없음"

    return "\n\n".join(
        [
            f"""[첨부자료 {index}]
파일명: {file.name}
형식: {file.type or 'unknown'}
내용:
{limit_text(source_text(file.name, file.type, file.text, file.data))}"""
            for index, file in enumerate(uploaded_files, start=1)
        ]
    )


@app.post("/generate")
def generate_report(req: WorkRequest):
    today = datetime.now().strftime("%Y년 %m월 %d일")
    related = search_data(req.work_content)
    email_section = format_email_sources(req.email_sources)
    file_section = format_uploaded_files(req.uploaded_files)

    prompt = f"""당신은 한성대학교 사업단운영위원회에 제출할 주간업무보고 초안을 작성하는 실무 보좌관입니다.
보고서는 위원들이 짧은 시간 안에 사업 진행 현황, 의사결정 필요 사항, 후속 조치를 파악할 수 있도록 작성합니다.

작성 원칙:
- 반드시 제공된 참고자료, Gmail 본문, Gmail 첨부자료, 사용자 첨부자료, 사용자가 입력한 업무 내용에 근거해 작성합니다.
- 자료에서 확인되지 않는 수치, 일정, 성과는 임의로 만들지 말고 "(확인 필요)"라고 표시합니다.
- 캘린더 일정은 사용하지 않습니다.
- 문체는 공식적이고 간결하게 유지합니다.
- 운영위원회 안건으로 검토가 필요한 내용은 별도 항목에 분리합니다.

[대상 회의]
{req.target_committee}

[보고 초점]
{req.meeting_focus}

[기본 사업단 참고자료]
{related}

[Gmail 참고자료]
{email_section}

[사용자 첨부자료]
{file_section}

[사용자 입력 업무 내용]
{req.work_content or '직접 입력 없음'}

아래 형식을 지켜 작성하세요.

사업단운영위원회 주간업무보고

작성일: {today}
보고대상: {req.target_committee}
보고초점: {req.meeting_focus}

1. 금주 주요 추진 실적
- 핵심 실적을 업무 단위로 정리
- 근거가 되는 메일/첨부/입력 내용이 있으면 자연스럽게 반영

2. 세부 진행 현황
- 사업 운영, 프로그램, 예산, 행정, 대외협력 등 관련 범주로 구분
- 진행률이나 수치는 자료에 있는 경우에만 작성

3. 운영위원회 검토 필요 안건
- 의결, 승인, 공유, 협의가 필요한 사항을 구분
- 안건이 명확하지 않으면 "해당 없음"으로 작성

4. 리스크 및 대응 방안
- 지연, 누락, 추가 확인이 필요한 사항
- 대응 계획 또는 담당 확인 필요 사항

5. 차주 추진 계획
- 후속 조치와 준비 사항 중심으로 작성

6. 확인 필요 사항
- 자료만으로 확정할 수 없는 항목을 목록화
"""

    response = gemini.models.generate_content(model="gemini-2.5-flash-lite", contents=prompt)
    return {"draft": response.text}


@app.post("/save-word")
def save_word(req: SaveRequest):
    doc = Document()
    doc.add_heading("사업단운영위원회 주간업무보고", level=1)
    doc.add_paragraph("")
    for line in req.content.split("\n"):
        doc.add_paragraph(line)

    filepath = os.path.join(BASE_DIR, f"output_{datetime.now().strftime('%Y%m%d%H%M%S')}.docx")
    doc.save(filepath)
    return FileResponse(
        path=filepath,
        filename=req.filename,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.post("/save-ppt")
def save_ppt(req: SaveRequest):
    prs = Presentation()
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = "사업단운영위원회 주간업무보고"
    title_slide.placeholders[1].text = datetime.now().strftime("%Y년 %m월 %d일")

    lines = [line for line in req.content.split("\n") if line.strip()]
    chunk_size = 9
    for index in range(0, len(lines), chunk_size):
        chunk = lines[index : index + chunk_size]
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = chunk[0][:42]
        text_frame = slide.placeholders[1].text_frame
        text_frame.text = chunk[0]
        for line in chunk[1:]:
            text_frame.add_paragraph().text = line

    filepath = os.path.join(BASE_DIR, f"output_{datetime.now().strftime('%Y%m%d%H%M%S')}.pptx")
    prs.save(filepath)
    return FileResponse(
        path=filepath,
        filename=req.filename,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
