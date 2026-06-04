#!/usr/bin/env python3
from __future__ import annotations

import io
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from pathlib import Path
from threading import Lock
from uuid import uuid4

import fitz
from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.datastructures import FileStorage
from werkzeug.utils import secure_filename

from pdf_text_to_ppt import APP_VERSION, CONVERSION_PROFILES, ConversionSummary, convert_pdf_to_pptx


BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PDF = BASE_DIR / "mypaper.pdf"
RUNTIME_DIR = BASE_DIR / "runtime"
UPLOAD_DIR = RUNTIME_DIR / "uploads"
OUTPUT_DIR = RUNTIME_DIR / "outputs"
MAX_UPLOAD_SIZE_MB = 80
ALLOWED_LAYOUTS = {"clean", "positioned"}
ALLOWED_PROFILES = set(CONVERSION_PROFILES)
ALLOWED_ANSWER_MODES = {"inline", "skip"}
PROFILE_LABELS = {
    "workbook": "练习册 / 答案页",
    "generic": "通用 PDF",
}
ANSWER_MODE_LABELS = {
    "inline": "题目下方显示答案",
    "skip": "不显示答案",
}
MAX_RECENT_JOBS = 20


app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE_MB * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("PDF_TRANS_PPT_SECRET", "local-dev-secret")
app.config["SYNC_JOBS"] = os.environ.get("PDF_TRANS_PPT_SYNC_JOBS") == "1"

JOB_EXECUTOR = ThreadPoolExecutor(max_workers=max(1, int(os.environ.get("PDF_TRANS_PPT_WORKERS", "2"))))
JOB_METADATA_LOCK = Lock()


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
        return 29, skip_from_page_label(29)
    if mode == "auto":
        return 0, "自动检测答案页"
    if mode == "all":
        return None, "保留全部页面"
    if mode == "custom":
        raw_value = request.form.get("skip_from_page", "").strip()
        if not raw_value:
            raise ValueError("请输入自定义起始页")
        value = int(raw_value)
        if value < 1:
            raise ValueError("自定义起始页必须大于 0")
        return value, skip_from_page_label(value)
    raise ValueError("未知的跳过策略")


def parse_profile() -> str:
    profile = request.form.get("profile", "workbook")
    if profile not in ALLOWED_PROFILES:
        raise ValueError("未知的转换配置")
    return profile


def parse_answer_mode() -> str:
    answer_mode = request.form.get("answer_mode", "inline")
    if answer_mode not in ALLOWED_ANSWER_MODES:
        raise ValueError("未知的答案展示方式")
    return answer_mode


def display_filename(filename: str) -> str:
    return filename.rsplit("\\", 1)[-1].rsplit("/", 1)[-1]


def storage_stem(filename: str) -> str:
    return secure_filename(Path(display_filename(filename)).stem) or "uploaded"


def save_uploaded_pdf(job_id: str, upload: FileStorage) -> tuple[Path, str]:
    original_name = display_filename(upload.filename or "")
    if not original_name:
        raise ValueError("请选择 PDF 文件")
    if Path(original_name).suffix.lower() != ".pdf":
        raise ValueError("只支持上传 PDF 文件")

    input_path = UPLOAD_DIR / f"{job_id}_{storage_stem(original_name)}.pdf"
    upload.save(input_path)
    return input_path, original_name


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


def app_version_suffix() -> str:
    major_minor = ".".join(APP_VERSION.split(".")[:2])
    return f"v{major_minor}"


def output_path_for(job_id: str, original_name: str, answer_mode: str = "skip") -> tuple[Path, str]:
    download_stem = Path(display_filename(original_name)).stem.strip() or "converted"
    storage_name = secure_filename(download_stem) or "converted"
    mode_suffix = "_answers" if answer_mode == "inline" else ""
    version_suffix = app_version_suffix()
    download_name = f"{download_stem}{mode_suffix}_{version_suffix}.pptx"
    return OUTPUT_DIR / f"{job_id}_{storage_name}{mode_suffix}_{version_suffix}.pptx", download_name


def skip_from_page_label(page: int) -> str:
    return f"从第 {page} 页起跳过（含该页及之后）"


def display_skip_label(label: str) -> str:
    if label == "第 29 页开始跳过":
        return skip_from_page_label(29)
    match = re.fullmatch(r"第 (\d+) 页开始跳过", label)
    if match:
        return skip_from_page_label(int(match.group(1)))
    match = re.fullmatch(r"自动检测到第 (\d+) 页开始跳过", label)
    if match:
        return f"自动检测到第 {match.group(1)} 页起跳过（含该页及之后）"
    return label


def skip_label(summary: ConversionSummary, requested_label: str) -> str:
    if summary.answer_mode == "inline":
        if summary.skipped_from_page is None:
            return "答案已放在题下"
        return f"第 {summary.skipped_from_page} 页起作为答案来源"
    if summary.skipped_from_page is None:
        return "保留全部页面"
    if requested_label == "自动检测答案页":
        return f"自动检测到第 {summary.skipped_from_page} 页起跳过（含该页及之后）"
    return display_skip_label(requested_label)


def profile_label(profile: str) -> str:
    return PROFILE_LABELS.get(profile, profile)


def answer_mode_label(answer_mode: str) -> str:
    return ANSWER_MODE_LABELS.get(answer_mode, answer_mode)


def answer_match_label(metadata: dict) -> str:
    if metadata.get("answer_mode") != "inline":
        return ""
    matched = metadata.get("matched_answers", 0)
    unmatched = metadata.get("unmatched_topics", 0)
    total = matched + unmatched
    if total <= 0:
        return "答案匹配：待生成"
    return f"答案匹配：{matched}/{total}"


def build_initial_job_metadata(
    job_id: str,
    original_name: str,
    input_source: str,
    input_path: Path,
    output_path: Path,
    download_name: str,
    skip_from_page: int | None,
    requested_skip_label: str,
    layout: str,
    profile: str,
    answer_mode: str,
) -> dict:
    created_at = datetime.now().isoformat(timespec="seconds")
    return {
        "job_id": job_id,
        "status": "queued",
        "error_message": "",
        "created_at": created_at,
        "created_at_display": created_at.replace("T", " "),
        "updated_at": created_at,
        "input_source": input_source,
        "input_path": str(input_path.resolve()),
        "original_name": original_name,
        "download_name": download_name,
        "output_path": str(output_path.resolve()),
        "output_size": 0,
        "output_size_display": "待生成",
        "source_pages": 0,
        "processed_pages": 0,
        "skipped_pages": 0,
        "skip_from_page": skip_from_page,
        "requested_skip_label": requested_skip_label,
        "skip_label": requested_skip_label,
        "layout": layout,
        "profile": profile,
        "profile_label": profile_label(profile),
        "answer_mode": answer_mode,
        "answer_mode_label": answer_mode_label(answer_mode),
        "matched_answers": 0,
        "unmatched_topics": 0,
        "unused_answers": 0,
        "answer_match_label": "答案匹配：待生成" if answer_mode == "inline" else "",
        "slides": 0,
    }


def apply_summary_to_metadata(metadata: dict, summary: ConversionSummary) -> dict:
    output_size = summary.pptx_path.stat().st_size
    metadata.update(
        {
            "status": "done",
            "error_message": "",
            "updated_at": datetime.now().isoformat(timespec="seconds"),
            "output_size": output_size,
            "output_size_display": human_file_size(output_size),
            "source_pages": summary.source_pages,
            "processed_pages": summary.processed_pages,
            "skipped_pages": max(summary.source_pages - summary.processed_pages, 0),
            "skip_label": skip_label(summary, metadata.get("requested_skip_label", "")),
            "layout": summary.layout,
            "answer_mode": summary.answer_mode,
            "answer_mode_label": answer_mode_label(summary.answer_mode),
            "matched_answers": summary.matched_answers,
            "unmatched_topics": summary.unmatched_topics,
            "unused_answers": summary.unused_answers,
            "slides": summary.slides,
        }
    )
    metadata["answer_match_label"] = answer_match_label(metadata)
    return metadata


def save_job_metadata(metadata: dict) -> None:
    with JOB_METADATA_LOCK:
        metadata_path(metadata["job_id"]).write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding="utf-8")


def load_job(job_id: str) -> dict:
    path = metadata_path(job_id)
    if not path.exists():
        abort(404)
    metadata = json.loads(path.read_text(encoding="utf-8"))
    output_path_value = metadata.get("output_path")
    output_exists = bool(output_path_value) and Path(output_path_value).exists()
    if "status" not in metadata:
        metadata["status"] = "done" if output_exists else "failed"
    metadata.setdefault("profile", "workbook")
    metadata.setdefault("profile_label", profile_label(metadata["profile"]))
    metadata.setdefault("answer_mode", "skip")
    metadata.setdefault("answer_mode_label", answer_mode_label(metadata["answer_mode"]))
    metadata.setdefault("error_message", "")
    metadata["skip_label"] = display_skip_label(metadata.get("skip_label", ""))
    metadata["answer_match_label"] = answer_match_label(metadata)
    if metadata.get("status") == "done" and not output_exists:
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
        output_path_value = metadata.get("output_path")
        output_exists = bool(output_path_value) and Path(output_path_value).exists()
        status = metadata.get("status", "done" if output_exists else "")
        if status == "done" and output_exists:
            metadata.setdefault("profile_label", profile_label(metadata.get("profile", "workbook")))
            metadata.setdefault("answer_mode", "skip")
            metadata.setdefault("answer_mode_label", answer_mode_label(metadata["answer_mode"]))
            metadata["skip_label"] = display_skip_label(metadata.get("skip_label", ""))
            metadata["answer_match_label"] = answer_match_label(metadata)
            jobs.append(metadata)
    return sorted(jobs, key=lambda item: item.get("created_at", ""), reverse=True)[:limit]


def path_inside(path: Path, parent: Path) -> bool:
    try:
        return parent.resolve() in path.resolve().parents
    except OSError:
        return False


def safe_unlink_runtime_file(path_value: str | None) -> None:
    if not path_value:
        return
    path = Path(path_value)
    if path.exists() and path.is_file() and (path_inside(path, UPLOAD_DIR) or path_inside(path, OUTPUT_DIR)):
        path.unlink()


def cleanup_old_jobs(keep: int = MAX_RECENT_JOBS) -> None:
    if not OUTPUT_DIR.exists():
        return
    metadata_files = sorted(
        OUTPUT_DIR.glob("*.json"),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    for path in metadata_files[keep:]:
        try:
            metadata = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            metadata = {}
        safe_unlink_runtime_file(metadata.get("output_path"))
        safe_unlink_runtime_file(metadata.get("input_path"))
        path.unlink(missing_ok=True)


def delete_job_files(job_id: str) -> None:
    path = metadata_path(job_id)
    if not path.exists():
        abort(404)
    try:
        metadata = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        metadata = {}
    safe_unlink_runtime_file(metadata.get("output_path"))
    safe_unlink_runtime_file(metadata.get("input_path"))
    path.unlink(missing_ok=True)


def render_page(result: dict | None = None):
    return render_template(
        "index.html",
        app_version=APP_VERSION,
        max_upload_size_mb=MAX_UPLOAD_SIZE_MB,
        sample_exists=SAMPLE_PDF.exists(),
        profiles=PROFILE_LABELS,
        answer_modes=ANSWER_MODE_LABELS,
        result=result,
        recent_jobs=recent_jobs(MAX_RECENT_JOBS),
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
    try:
        with fitz.open(pdf_path) as doc:
            if doc.page_count == 0:
                abort(404)
            pixmap = doc[0].get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
            image = io.BytesIO(pixmap.tobytes("png"))
    except fitz.FileDataError:
        abort(404)
    image.seek(0)
    return send_file(image, mimetype="image/png", max_age=60)


def conversion_error_message(error: Exception) -> str:
    if isinstance(error, fitz.FileDataError):
        return "无法打开 PDF 文件，请确认文件未损坏且不是加密文件"
    if isinstance(error, ValueError):
        return str(error)
    return "转换失败，请检查 PDF 内容后重试"


def run_conversion_job(job_id: str) -> None:
    with app.app_context():
        metadata = load_job(job_id)
        metadata["status"] = "running"
        metadata["updated_at"] = datetime.now().isoformat(timespec="seconds")
        save_job_metadata(metadata)

        try:
            summary = convert_pdf_to_pptx(
                Path(metadata["input_path"]),
                Path(metadata["output_path"]),
                skip_from_page=metadata.get("skip_from_page"),
                layout=metadata.get("layout", "clean"),
                profile=metadata.get("profile", "workbook"),
                answer_mode=metadata.get("answer_mode", "skip"),
            )
            metadata = apply_summary_to_metadata(metadata, summary)
        except Exception as error:
            app.logger.exception("Conversion failed for job %s", job_id)
            metadata["status"] = "failed"
            metadata["error_message"] = conversion_error_message(error)
            metadata["updated_at"] = datetime.now().isoformat(timespec="seconds")
        save_job_metadata(metadata)


def enqueue_conversion(job_id: str) -> None:
    if app.config.get("SYNC_JOBS"):
        run_conversion_job(job_id)
        return
    JOB_EXECUTOR.submit(run_conversion_job, job_id)


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

        profile = parse_profile()
        answer_mode = parse_answer_mode()
        skip_from_page, requested_skip_label = parse_skip_from_page()
        input_path, original_name, input_source = resolve_input_pdf(job_id)
        output_path, download_name = output_path_for(job_id, original_name, answer_mode)

        metadata = build_initial_job_metadata(
            job_id=job_id,
            original_name=original_name,
            input_source=input_source,
            input_path=input_path,
            output_path=output_path,
            download_name=download_name,
            skip_from_page=skip_from_page,
            requested_skip_label=requested_skip_label,
            layout=layout,
            profile=profile,
            answer_mode=answer_mode,
        )
        save_job_metadata(metadata)
        enqueue_conversion(job_id)
        cleanup_old_jobs()
    except ValueError as error:
        flash(str(error), "error")
        return redirect(url_for("index"))

    return redirect(url_for("job", job_id=job_id))


@app.get("/jobs/<job_id>")
def job(job_id: str):
    return render_page(load_job(job_id))


@app.get("/status/<job_id>")
def job_status(job_id: str):
    metadata = load_job(job_id)
    return jsonify(
        {
            "job_id": metadata["job_id"],
            "status": metadata.get("status", "done"),
            "error_message": metadata.get("error_message", ""),
            "updated_at": metadata.get("updated_at", ""),
        }
    )


@app.get("/download/<job_id>")
def download(job_id: str):
    metadata = load_job(job_id)
    if metadata.get("status") != "done":
        abort(404)
    return send_file(
        Path(metadata["output_path"]),
        as_attachment=True,
        download_name=metadata["download_name"],
    )


@app.post("/jobs/<job_id>/delete")
def delete_job(job_id: str):
    delete_job_files(job_id)
    return redirect(url_for("index"))


@app.get("/preview/sample")
def sample_preview():
    return render_pdf_preview(SAMPLE_PDF)


@app.get("/preview/<job_id>")
def job_preview(job_id: str):
    metadata = load_job(job_id)
    return render_pdf_preview(safe_pdf_path(metadata["input_path"]))


if __name__ == "__main__":
    ensure_runtime_dirs()
    cleanup_old_jobs()
    debug = os.environ.get("FLASK_DEBUG") == "1"
    app.run(host="127.0.0.1", port=5000, debug=debug)
