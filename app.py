"""Local web UI for batch invoice processing."""

from __future__ import annotations

import io
import os
import re
import tempfile
import time
import zipfile
from pathlib import Path, PurePosixPath

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from dotenv import load_dotenv

from processor import process, valid_uscc


BASE = Path(__file__).resolve().parent
load_dotenv(BASE / ".env")
MAX_FILES = 1000
MAX_UPLOAD_BYTES = 200 * 1024 * 1024
MAX_FILE_BYTES = 30 * 1024 * 1024
JOBS = {}
JOB_TTL_SECONDS = 24 * 60 * 60

app = FastAPI(title="发票整理台", docs_url=None, redoc_url=None)
app.mount("/static", StaticFiles(directory=BASE / "static"), name="static")


def _prune_jobs():
    now = time.monotonic()
    for job_id, (job, created) in list(JOBS.items()):
        if now - created > JOB_TTL_SECONDS:
            job.cleanup()
            del JOBS[job_id]


@app.get("/", response_class=HTMLResponse)
def index():
    return (BASE / "templates" / "index.html").read_text(encoding="utf-8")


def _safe_name(name: str) -> str:
    name = name.replace("\\", "/").replace("\x00", "")
    parts = PurePosixPath(name).parts
    if not parts or any(part in (".", "..") for part in parts) or name.startswith("/"):
        raise HTTPException(400, "文件路径不安全")
    if re.match(r"^[A-Za-z]:", name):
        raise HTTPException(400, "文件路径不安全")
    return "/".join(parts)


def _store_pdf(stream, name: str, directory: Path, index: int):
    safe_name = _safe_name(name)
    if not safe_name.lower().endswith(".pdf") or PurePosixPath(safe_name).name.startswith("._"):
        return None
    destination = directory / f"{index:05d}.pdf"
    size = 0
    with destination.open("wb") as target:
        while chunk := stream.read(1024 * 1024):
            size += len(chunk)
            if size > MAX_FILE_BYTES:
                destination.unlink(missing_ok=True)
                raise HTTPException(413, f"文件过大：{safe_name}")
            target.write(chunk)
    return destination, safe_name


@app.post("/api/process")
async def process_upload(
    files: list[UploadFile] = File(...),
    buyer_names: str = Form(""),
    tax_ids: str = Form(""),
    use_llm: bool = Form(False),
):
    _prune_jobs()
    if use_llm and (not os.getenv("LLM_API_KEY") or not os.getenv("LLM_BASE_URL")):
        raise HTTPException(400, "大模型辅助分类未配置：请在本机 .env 中填写服务地址和密钥，或取消勾选")
    if len(files) > MAX_FILES:
        raise HTTPException(413, f"最多上传 {MAX_FILES} 个文件")
    job = tempfile.TemporaryDirectory(prefix="tax-project-")
    root = Path(job.name)
    inputs = root / "input"
    inputs.mkdir()
    collected = []
    total_size = 0
    extracted_size = 0
    try:
        for uploaded in files:
            name = _safe_name(uploaded.filename or "")
            if name.lower().endswith(".zip"):
                archive_data = await uploaded.read(MAX_UPLOAD_BYTES + 1)
                total_size += len(archive_data)
                if total_size > MAX_UPLOAD_BYTES:
                    raise HTTPException(413, "上传文件总大小超过 200 MB")
                try:
                    with zipfile.ZipFile(io.BytesIO(archive_data)) as archive:
                        members = [info for info in archive.infolist() if info.filename.lower().endswith(".pdf") and not info.filename.startswith("__MACOSX/")]
                        if len(members) + len(collected) > MAX_FILES:
                            raise HTTPException(413, f"最多处理 {MAX_FILES} 份 PDF")
                        for info in members:
                            if info.file_size > MAX_FILE_BYTES:
                                raise HTTPException(413, f"文件过大：{info.filename}")
                            source_name = info.filename
                            try:
                                source_name = source_name.encode("cp437").decode("utf-8") if not info.flag_bits & 0x800 else source_name
                            except UnicodeError:
                                pass
                            with archive.open(info) as source:
                                stored = _store_pdf(source, source_name, inputs, len(collected))
                            if stored:
                                extracted_size += stored[0].stat().st_size
                                if extracted_size > MAX_UPLOAD_BYTES:
                                    raise HTTPException(413, "解压后 PDF 总大小超过 200 MB")
                                collected.append(stored)
                except zipfile.BadZipFile:
                    raise HTTPException(400, f"压缩包损坏：{name}")
            elif name.lower().endswith(".pdf"):
                stored = _store_pdf(uploaded.file, name, inputs, len(collected))
                if stored:
                    total_size += stored[0].stat().st_size
                    extracted_size += stored[0].stat().st_size
                    if total_size > MAX_UPLOAD_BYTES:
                        raise HTTPException(413, "上传文件总大小超过 200 MB")
                    if extracted_size > MAX_UPLOAD_BYTES:
                        raise HTTPException(413, "PDF 总大小超过 200 MB")
                    collected.append(stored)
        if not collected:
            raise HTTPException(400, "没有找到 PDF，请上传文件夹或 ZIP 压缩包")
        if len(collected) > MAX_FILES:
            raise HTTPException(413, f"最多处理 {MAX_FILES} 份 PDF")
        names = [x.strip() for x in re.split(r"[,，;；\n]", buyer_names) if x.strip()]
        ids = [x.strip() for x in re.split(r"[,，;；\n]", tax_ids) if x.strip()]
        invalid_ids = [value for value in ids if len(value) == 18 and not valid_uscc(value)]
        if invalid_ids:
            raise HTTPException(400, "设定的统一社会信用代码校验位错误，请检查输入")
        archive, summary = process(collected, root / "result", names, ids, use_llm)
        job_id = os.urandom(16).hex()
        JOBS[job_id] = (job, time.monotonic())
        return JSONResponse({"job_id": job_id, "summary": summary})
    except HTTPException:
        job.cleanup()
        raise
    except Exception as exc:
        job.cleanup()
        raise HTTPException(500, f"处理失败：{exc}") from exc


@app.get("/api/download/{job_id}")
def download(job_id: str):
    _prune_jobs()
    stored = JOBS.get(job_id)
    if not stored:
        raise HTTPException(404, "结果不存在或服务已重启")
    job, _created = stored
    archive = Path(job.name) / "发票分类结果.zip"
    if not archive.is_file():
        raise HTTPException(404, "结果文件不存在")
    return FileResponse(archive, filename="发票分类结果.zip", media_type="application/zip")
