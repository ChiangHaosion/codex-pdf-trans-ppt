from __future__ import annotations

import io
import json
from pathlib import Path

import fitz
import pytest

import web_app
from pdf_text_to_ppt import ConversionSummary


@pytest.fixture()
def client_runtime(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    runtime_dir = tmp_path / "runtime"
    upload_dir = runtime_dir / "uploads"
    output_dir = runtime_dir / "outputs"

    monkeypatch.setattr(web_app, "RUNTIME_DIR", runtime_dir)
    monkeypatch.setattr(web_app, "UPLOAD_DIR", upload_dir)
    monkeypatch.setattr(web_app, "OUTPUT_DIR", output_dir)
    monkeypatch.setattr(web_app, "SAMPLE_PDF", tmp_path / "missing-sample.pdf")
    web_app.app.config.update(TESTING=True, SYNC_JOBS=True)
    web_app.ensure_runtime_dirs()

    yield web_app.app.test_client(), output_dir

    web_app.app.config.update(TESTING=False, SYNC_JOBS=False)


def read_metadata(output_dir: Path, location: str) -> dict:
    job_id = location.rstrip("/").rsplit("/", 1)[-1]
    return json.loads((output_dir / f"{job_id}.json").read_text(encoding="utf-8"))


def make_pdf_bytes(text: str = "demo") -> bytes:
    doc = fitz.open()
    page = doc.new_page(width=320, height=220)
    page.insert_text((40, 72), text)
    data = doc.tobytes()
    doc.close()
    return data


def test_chinese_pdf_filename_upload_is_accepted(client_runtime, monkeypatch: pytest.MonkeyPatch) -> None:
    client, output_dir = client_runtime

    def fake_convert(pdf_path: Path, pptx_path: Path, skip_from_page, layout: str, profile: str, answer_mode: str) -> ConversionSummary:
        assert pdf_path.name.endswith("_uploaded.pdf")
        assert profile == "workbook"
        assert answer_mode == "inline"
        pptx_path.write_bytes(b"pptx")
        return ConversionSummary(
            pdf_path=pdf_path,
            pptx_path=pptx_path,
            source_pages=1,
            processed_pages=1,
            skipped_from_page=skip_from_page,
            slides=1,
            layout=layout,
            answer_mode=answer_mode,
            matched_answers=1,
        )

    monkeypatch.setattr(web_app, "convert_pdf_to_pptx", fake_convert)

    response = client.post(
        "/convert",
        data={
            "pdf_file": (io.BytesIO(b"not used by fake converter"), "测试.pdf"),
            "layout": "clean",
            "profile": "workbook",
            "answer_mode": "inline",
            "skip_mode": "all",
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 302
    metadata = read_metadata(output_dir, response.headers["Location"])
    assert metadata["status"] == "done"
    assert metadata["original_name"] == "测试.pdf"
    assert metadata["download_name"] == "测试_answers_v0.5.pptx"
    assert metadata["answer_mode"] == "inline"
    assert metadata["answer_match_label"] == "答案匹配：1/1"


def test_index_does_not_load_sample_preview_by_default(client_runtime, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    client, _output_dir = client_runtime
    sample_pdf = tmp_path / "sample.pdf"
    sample_pdf.write_bytes(b"not rendered")
    monkeypatch.setattr(web_app, "SAMPLE_PDF", sample_pdf)

    response = client.get("/")

    assert response.status_code == 200
    assert b"/preview/sample" not in response.data
    assert "使用项目样例".encode() in response.data


def test_skip_options_use_clear_answer_page_labels(client_runtime) -> None:
    client, _output_dir = client_runtime

    response = client.get("/")

    assert response.status_code == 200
    assert "答案页处理".encode() in response.data
    assert "29 页起跳过".encode() in response.data
    assert "含之后页面".encode() in response.data
    assert "预检查".encode() in response.data
    assert "<legend>跳过页</legend>".encode() not in response.data


def test_precheck_endpoint_returns_inline_answer_summary(client_runtime, monkeypatch: pytest.MonkeyPatch) -> None:
    client, _output_dir = client_runtime

    def fake_precheck(doc, skip_from_page, profile: str) -> dict:
        assert doc.page_count == 1
        assert skip_from_page == 29
        assert profile == "workbook"
        return {
            "source_pages": 1,
            "answer_start_page": 1,
            "processed_pages": 0,
            "topic_count": 2,
            "answer_count": 2,
            "matched_answers": 2,
            "unmatched_topics": 0,
            "unused_answers": 0,
            "confidence_counts": {"high": 2, "medium": 0, "low": 0, "none": 0},
            "matches": [],
            "unmatched_items": [],
            "unused_answer_items": [],
            "warnings": [],
        }

    monkeypatch.setattr(web_app, "precheck_inline_answer_document", fake_precheck)

    response = client.post(
        "/precheck",
        data={
            "pdf_file": (io.BytesIO(make_pdf_bytes()), "demo.pdf"),
            "profile": "workbook",
            "answer_mode": "inline",
            "skip_mode": "default",
        },
        content_type="multipart/form-data",
    )

    payload = response.get_json()
    assert response.status_code == 200
    assert payload["ok"] is True
    assert payload["original_name"] == "demo.pdf"
    assert payload["precheck"]["matched_answers"] == 2


def test_generated_files_are_rendered_in_separate_list(client_runtime) -> None:
    client, output_dir = client_runtime
    output_path = output_dir / "job123_demo_v0.3.pptx"
    output_path.write_bytes(b"pptx")
    metadata = {
        "job_id": "job123",
        "status": "done",
        "created_at": "2026-06-04T13:00:00",
        "created_at_display": "2026-06-04 13:00:00",
        "original_name": "demo.pdf",
        "download_name": "demo_v0.3.pptx",
        "output_path": str(output_path),
        "output_size_display": "10 B",
        "processed_pages": 2,
        "skip_label": "保留全部页面",
        "layout": "clean",
        "profile": "generic",
        "profile_label": "通用 PDF",
        "answer_mode": "skip",
        "answer_mode_label": "不显示答案",
        "slides": 3,
    }
    (output_dir / "job123.json").write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")

    response = client.get("/")

    assert response.status_code == 200
    assert "已生成文件".encode() in response.data
    assert b"demo_v0.3.pptx" in response.data
    assert "删除".encode() in response.data


def test_delete_job_removes_metadata_output_and_uploaded_pdf(client_runtime) -> None:
    client, output_dir = client_runtime
    upload_dir = output_dir.parent / "uploads"
    input_path = upload_dir / "job123_demo.pdf"
    output_path = output_dir / "job123_demo_v0.3.pptx"
    input_path.write_bytes(b"pdf")
    output_path.write_bytes(b"pptx")
    metadata = {
        "job_id": "job123",
        "status": "done",
        "created_at": "2026-06-04T13:00:00",
        "created_at_display": "2026-06-04 13:00:00",
        "original_name": "demo.pdf",
        "download_name": "demo_v0.3.pptx",
        "input_path": str(input_path),
        "output_path": str(output_path),
        "output_size_display": "10 B",
        "processed_pages": 2,
        "skip_label": "保留全部页面",
        "layout": "clean",
        "profile": "generic",
        "profile_label": "通用 PDF",
        "answer_mode": "skip",
        "answer_mode_label": "不显示答案",
        "slides": 3,
    }
    metadata_path = output_dir / "job123.json"
    metadata_path.write_text(json.dumps(metadata, ensure_ascii=False), encoding="utf-8")

    response = client.post("/jobs/job123/delete")

    assert response.status_code == 302
    assert response.headers["Location"] == "/"
    assert not metadata_path.exists()
    assert not output_path.exists()
    assert not input_path.exists()


def test_invalid_pdf_failure_is_saved_on_job(client_runtime, monkeypatch: pytest.MonkeyPatch) -> None:
    client, output_dir = client_runtime

    def fake_convert(pdf_path: Path, pptx_path: Path, skip_from_page, layout: str, profile: str, answer_mode: str) -> ConversionSummary:
        raise fitz.FileDataError("bad pdf")

    monkeypatch.setattr(web_app, "convert_pdf_to_pptx", fake_convert)

    response = client.post(
        "/convert",
        data={
            "pdf_file": (io.BytesIO(b"bad pdf"), "broken.pdf"),
            "layout": "clean",
            "profile": "workbook",
            "answer_mode": "inline",
            "skip_mode": "default",
        },
        content_type="multipart/form-data",
    )

    assert response.status_code == 302
    metadata = read_metadata(output_dir, response.headers["Location"])
    assert metadata["status"] == "failed"
    assert "无法打开 PDF" in metadata["error_message"]
