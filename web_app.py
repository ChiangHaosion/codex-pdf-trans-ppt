#!/usr/bin/env python3
from __future__ import annotations

import io
import json
import os
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import fitz
from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from pdf_text_to_ppt import APP_VERSION, ConversionSummary, convert_pdf_to_pptx


BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PDF = BASE_DIR / "mypaper.pdf"
RUNTIME_DIR = BASE_DIR / "runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_DIR = RUNTIME_DIR / "outputs"
MAX_UPLOAD_SIZE_MB = 80
ALLOWED_LAYOUTS = {"clean", "positioned"}


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE_MB * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("PDF_TRANS_PPT_SECRET", "local-dev-secret")


def ensure_runtime_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


def human_file_size(size: int) -> str:
    value = float(size)
    for unit in ("B", "KB", "MB", "GB"):
        if value < 1024 or unit == "GB":
            return f"{value:.1f} {unit}" if unit != "B" else f"{int(value)} B"
        value /= 1024
    return f"{value:.1f} GB"


def parse_skip_from_page() -> tuple[int | None, str]:
    mode = request.form.get("skip_mode", "default")
    if mode == "default":
        return 29, "第 29 页开始跳过"
    if mode == "auto":
        return 0, "自动检测答案页"
    if mode == "all":
        return None, "保留全部页面"
    if mode == "custom":
        raw_value = request.form.get("skip_from_page", "").strip()
        if not raw_value:
            raise ValueError("请输入自定义跳过页码")
        value = int(raw_value)
        if value < 1:
            raise ValueError("自定义跳过页码必须大于 0")
        return value, f"第 {value} 页开始跳过"
    raise ValueError("未知的跳过策略")


def save_uploaded_pdf(job_id: str, upload: FileStorage) -> tuple[Path, str]:
    filename = secure_filename(upload.filename or "")
    if not filename:
        raise ValueError("请选择 PDF 文件")
    if Path(filename).suffix.lower() != ".pdf":
        raise ValueError("只支持上传 PDF 文件")

    input_path = UPLOAD_DIR / f"{job_id}_{filename}"
    upload.save(input_path)
    return input_path, filename


def resolve_input_pdf(job_id: str) -> tuple[Path, str, str]:
    upload = request.files.get("pdf_file")
    use_sample = request.form.get("source") == "sample"

    if upload and upload.filename:
        input_path, filename = save_uploaded_pdf(job_id, upload)
        return input_path, filename, "uploaded"

    if use_sample and SAMPLE_PDF.exists():
        return SAMPLE_PDF, SAMPLE_PDF.name, "sample"

    raise ValueError("请上传 PDF，或选择使用项目内置样例")


def metadata_path(job_id: str) -> Path:
    return OUTPUT_DIR / f"{job_id}.json"


def output_path_for(job_id: str, original_name: str) -> tuple[Path, str]:
    stem = secure_filename(Path(original_name).stem) or "converted"
    download_name = f"{stem}_v0.2.pptx"
    return OUTPUT_DIR / f"{job_id}_{download_name}", download_name


def skip_label(summary: ConversionSummary, requested_label: str) -> str:
    if summary.skipped_from_page is None:
        return "保留全部页面"
    if requested_label == "自动检测答案页":
        return f"自动检测到第 {summary.skipped_from_page} 页开始跳过"
    return requested_label


def build_job_metadata(
    job_id: str,
    summary: ConversionSummary,
    original_name: str,
    input_source: str,
    input_path: Path,
    download_name: str,
    requested_skip_label: str,
) -> dict:
    output_size = summary.pptx_path.stat().st_size
    created_at = datetime.now().isoformat(timespec="seconds")
    return {
        "job_id": job_id,
        "created_at": created_at,
        "created_at_display": created_at.replace("T", " "),
        "input_source": input_source,
        "input_path": str(input_path.resolve()),
        "original_name": original_name,
        "download_name": download_name,
        "output_path": str(summary.pptx_path.resolve()),
        "output_size": output_size,
        "output_size_display": human_file_size(output_size),
        "source_pages": summary.source_pages,
        "processed_pages": summary.processed_pages,
        "skipped_pages": max(summary.source_pages - summary.processed_pages, 0),
        "skip_label": skip_label(summary, requested_skip_label),
        "layout": summary.layout,
        "slides": summary.slides,
    }


def save_job_metadata(metadata: dict) -> None:
    metadata_path(metadata["job_id"]).write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")


def load_job(job_id: str) -> dict:
    path = metadata_path(job_id)
    if not path.exists():
        abort(404)
    metadata = json.loads(path.read_text(encoding="utf-8"))
    if not Path(metadata["output_path"]).exists():
        abort(404)
    return metadata


def recent_jobs(limit: int = 5) -> list[dict]:
    jobs: list[dict] = []
    if not OUTPUT_DIR.exists():
        return jobs
    for path in OUTPUT_DIR.glob("*.json"):
        try:
            metadata = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        if Path(metadata.get("output_path", "")).exists():
            jobs.append(metadata)
    return sorted(jobs, key=lambda item: item.get("created_at", ""), reverse=True)[:limit]


def render_page(result: dict | None = None):
    return render_template(
        "index.html",
        app_version=APP_VERSION,
        max_upload_size_mb=MAX_UPLOAD_SIZE_MB,
        sample_exists=SAMPLE_PDF.exists(),
        result=result,
        recent_jobs=recent_jobs(),
    )


def safe_pdf_path(path_value: str) -> Path:
    path = Path(path_value).resolve()
    if path == SAMPLE_PDF.resolve():
        return path
    if UPLOAD_DIR.resolve() in path.parents:
        return path
    abort(404)


def render_pdf_preview(pdf_path: Path):
    if not pdf_path.exists():
        abort(404)
    with fitz.open(pdf_path) as doc:
        if doc.page_count == 0:
            abort(404)
        pixmap = doc[0].get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
        image = io.BytesIO(pixmap.tobytes("png"))
    image.seek(0)
    return send_file(image, mimetype="image/png", max_age=60)


@app.get("/")
def index():
    return render_page()


@app.post("/convert")
def convert():
    ensure_runtime_dirs()
    job_id = uuid4().hex[:12]

    try:
        layout = request.form.get("layout", "clean")
        if layout not in ALLOWED_LAYOUTS:
            raise ValueError("未知的排版模式")

        skip_from_page, requested_skip_label = parse_skip_from_page()
        input_path, original_name, input_source = resolve_input_pdf(job_id)
        output_path, download_name = output_path_for(job_id, original_name)

        summary = convert_pdf_to_pptx(input_path, output_path, skip_from_page=skip_from_page, layout=layout)
        metadata = build_job_metadata(
            job_id=job_id,
            summary=summary,
            original_name=original_name,
            input_source=input_source,
            input_path=input_path,
            download_name=download_name,
            requested_skip_label=requested_skip_label,
        )
        save_job_metadata(metadata)
    except ValueError as error:
        flash(str(error), "error")
        return redirect(url_for("index"))

    return redirect(url_for("job", job_id=job_id))


@app.get("/jobs/<job_id>")
def job(job_id: str):
    return render_page(load_job(job_id))


@app.get("/download/<job_id>")
def download(job_id: str):
    metadata = load_job(job_id)
    return send_file(
        Path(metadata["output_path"]),
        as_attachment=True,
        download_name=metadata["download_name"],
    )


@app.get("/preview/sample")
def sample_preview():
    return render_pdf_preview(SAMPLE_PDF)


@app.get("/preview/<job_id>")
def job_preview(job_id: str):
    metadata = load_job(job_id)
    return render_pdf_preview(safe_pdf_path(metadata["input_path"]))


if __name__ == "__main__":
    ensure_runtime_dirs()
    debug = os.environ.get("FLASK_DEBUG") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug)
