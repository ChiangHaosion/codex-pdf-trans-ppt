#!/usr/bin/env python3
from __future__ import annotations

import argparse
import io
import math
import re
from pathlib import Path

import fitz
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.util import Inches, Pt


EMU_PER_INCH = 914400
PDF_POINTS_PER_INCH = 72
CONTENT_MARGIN_X_PT = 56
CONTENT_MARGIN_Y_PT = 28
TEXT_INSET_X_PT = 8
TEXT_INSET_Y_PT = 3
CAPACITY_SAFETY = 0.98
IMAGE_AREA_HEIGHT_PT = 190


def pt_to_emu(value: float) -> int:
    return int(value / PDF_POINTS_PER_INCH * EMU_PER_INCH)


def int_to_rgb(value: int) -> RGBColor:
    return RGBColor((value >> 16) & 255, (value >> 8) & 255, value & 255)


def span_is_bold(span: dict) -> bool:
    font = span.get("font", "").lower()
    return "bold" in font or "black" in font or "heavy" in font


def span_is_italic(span: dict) -> bool:
    font = span.get("font", "").lower()
    return "italic" in font or "oblique" in font


def clean_font_name(font_name: str) -> str:
    for marker in ("-Bold", "-Italic", "-Oblique", ",Bold", ",Italic"):
        font_name = font_name.replace(marker, "")
    return font_name.strip() or "Arial"


def add_text_block(
    slide,
    block: dict,
    scale: float,
    offset_x: float = 0,
    offset_y: float = 0,
    margin_left: float = 0,
    margin_top: float = 0,
) -> None:
    x0, y0, x1, y1 = block["bbox"]
    left = pt_to_emu((x0 - offset_x) * scale) + pt_to_emu(margin_left)
    top = pt_to_emu((y0 - offset_y) * scale) + pt_to_emu(margin_top)
    width = max(pt_to_emu((x1 - x0) * scale), pt_to_emu(6))
    height = max(pt_to_emu((y1 - y0) * scale * 1.08), pt_to_emu(6))

    box = slide.shapes.add_textbox(left, top, width, height)
    frame = box.text_frame
    frame.clear()
    frame.margin_left = 0
    frame.margin_right = 0
    frame.margin_top = 0
    frame.margin_bottom = 0
    frame.word_wrap = False
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    frame.vertical_anchor = MSO_ANCHOR.TOP

    first_paragraph = True
    for line in block.get("lines", []):
        paragraph = frame.paragraphs[0] if first_paragraph else frame.add_paragraph()
        first_paragraph = False
        paragraph.alignment = PP_ALIGN.LEFT
        paragraph.space_before = Pt(0)
        paragraph.space_after = Pt(0)
        paragraph.line_spacing = 1.0

        for span in line.get("spans", []):
            text = span.get("text", "")
            if not text:
                continue
            run = paragraph.add_run()
            run.text = text
            font = run.font
            font.size = Pt(max(span.get("size", 8) * scale, 3))
            font.name = clean_font_name(span.get("font", "Arial"))
            font.bold = span_is_bold(span)
            font.italic = span_is_italic(span)
            font.color.rgb = int_to_rgb(span.get("color", 0))


def has_text(block: dict) -> bool:
    return any(span.get("text", "").strip() for line in block.get("lines", []) for span in line.get("spans", []))


def split_blocks_by_vertical_gaps(blocks: list[dict], min_gap: float = 60) -> list[list[dict]]:
    if not blocks:
        return []

    sorted_blocks = sorted(blocks, key=lambda block: (block["bbox"][1], block["bbox"][0]))
    segments: list[list[dict]] = []
    current: list[dict] = []
    current_bottom: float | None = None

    for block in sorted_blocks:
        top = block["bbox"][1]
        bottom = block["bbox"][3]
        if current and current_bottom is not None and top - current_bottom >= min_gap:
            segments.append(current)
            current = []
        current.append(block)
        current_bottom = max(current_bottom or bottom, bottom)

    if current:
        segments.append(current)
    return segments


def segment_bbox(blocks: list[dict], page_rect: fitz.Rect) -> tuple[float, float, float, float]:
    x0 = min(block["bbox"][0] for block in blocks)
    y0 = min(block["bbox"][1] for block in blocks)
    x1 = max(block["bbox"][2] for block in blocks)
    y1 = max(block["bbox"][3] for block in blocks)

    pad_x = 4
    pad_y = 6
    return (
        max(0, x0 - pad_x),
        max(0, y0 - pad_y),
        min(page_rect.width, x1 + pad_x),
        min(page_rect.height, y1 + pad_y),
    )


def block_text(block: dict) -> str:
    if "_text" in block:
        return block["_text"]

    lines: list[str] = []
    for line in block.get("lines", []):
        text = "".join(span.get("text", "") for span in line.get("spans", [])).strip()
        if text:
            lines.append(text)
    return normalize_text(join_pdf_lines(lines))


def should_start_new_line(previous: str, current: str) -> bool:
    if not previous:
        return False
    if previous == "阅读答题技巧小纸条":
        return True
    if previous.startswith(("【", "一、", "二、", "三、")):
        return True
    if current.startswith(("【", "一、", "二、", "三、", "四、", "五、", "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩", "⑪", "⑫", "⑬", "任务")):
        return True
    if previous.endswith(("。", "？", "?", "！", "!", "：", ":")) and len(previous) < 42:
        return True
    return False


def join_pdf_lines(lines: list[str]) -> str:
    paragraphs: list[str] = []
    current = ""
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if current and should_start_new_line(current, line):
            paragraphs.append(current)
            current = line
        else:
            current = f"{current}{line}" if current else line
    if current:
        paragraphs.append(current)
    return "\n".join(paragraphs).strip()


def normalize_text(text: str) -> str:
    text = re.sub(r"([（(])\s*\n\s*", r"\1", text)
    text = re.sub(r"\s*\n\s*([）)])", r"\1", text)
    text = re.sub(r"([\u4e00-\u9fff])\n([\u4e00-\u9fff])", r"\1\2", text)
    text = re.sub(r"([（(])\s+", r"\1", text)
    text = re.sub(r"\s+([）)])", r"\1", text)
    return text


def blocks_text_length(blocks: list[dict]) -> int:
    return sum(len(block_text(block)) for block in blocks)


def make_text_block_like(source: dict, text: str) -> dict:
    return {"bbox": source["bbox"], "lines": source.get("lines", []), "_text": text, "_span": first_span(source) or {}}


def is_table_or_chart_text(text: str) -> bool:
    compact = " ".join(text.split())
    percent_count = compact.count("%")
    table_keywords = (
        "统计表",
        "调查表",
        "图表",
        "比例",
        "用途 所占比例",
        "阅读方式",
        "每天阅读用时",
        "调查年级及人数",
        "经常使用",
        "偶尔使用",
        "从不使用",
    )
    digit_count = sum(char.isdigit() for char in compact)
    return percent_count >= 3 or digit_count >= 10 or any(keyword in compact for keyword in table_keywords)


def replace_table_text_with_placeholders(blocks: list[dict]) -> list[dict]:
    output: list[dict] = []
    table_mode = 0
    for block in blocks:
        text = block_text(block)
        if table_mode and "阅读答题技巧小纸条" in text:
            table_mode = 0
        if is_table_or_chart_text(text):
            if not output or block_text(output[-1]) != "【图表占位：请参照原 PDF 图表】":
                output.append(make_text_block_like(block, "【图表占位：请参照原 PDF 图表】"))
            table_mode = 8
            continue
        if table_mode:
            compact = " ".join(text.split())
            digit_count = sum(char.isdigit() for char in compact)
            looks_like_table_row = len(compact) < 90 or digit_count >= 4 or "%" in compact
            if looks_like_table_row:
                table_mode -= 1
                continue
            table_mode = 0
        output.append(block)
    return output


def should_merge_text_blocks(previous_text: str, current_text: str) -> bool:
    if not previous_text or not current_text:
        return False
    previous_tail = previous_text[-1]
    current_head = current_text[0]
    if previous_tail in "。！？!?：:；;）)】》”’…":
        return False
    if current_head in "【一二三四五六七八九十①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬任务":
        return False
    return "\u4e00" <= previous_tail <= "\u9fff" and "\u4e00" <= current_head <= "\u9fff"


def merge_adjacent_text_blocks(blocks: list[dict]) -> list[dict]:
    merged: list[dict] = []
    for block in blocks:
        text = block_text(block)
        if merged and should_merge_text_blocks(block_text(merged[-1]), text):
            previous = merged[-1]
            previous_text = block_text(previous)
            merged[-1] = make_text_block_like(previous, f"{previous_text}{text}")
        else:
            merged.append(block)
    return merged


def visual_units(text: str) -> float:
    units = 0.0
    for char in text:
        if char.isspace():
            units += 0.25
        elif ord(char) < 128:
            units += 0.55
        else:
            units += 1.0
    return units


def estimate_wrapped_lines(text: str, font_size: float, width_pt: float) -> int:
    chars_per_line = max(width_pt / (font_size * 0.92), 8)
    lines = 0
    for raw_line in text.splitlines() or [text]:
        line = raw_line.strip()
        if not line:
            continue
        lines += max(1, math.ceil(visual_units(line) / chars_per_line))
    return max(lines, 1)


def block_font_size(block: dict, base_size: int) -> int:
    text = block_text(block)
    span = first_span(block) or {}
    source_size = span.get("size", 10)
    is_heading = source_size >= 13 or span_is_bold(span)
    return base_size + 2 if is_heading and len(text) < 90 else base_size


def estimate_blocks_height(blocks: list[dict], width_pt: float, base_size: int = 18) -> float:
    height = 0.0
    for block in blocks:
        text = block_text(block)
        if not text:
            continue
        font_size = block_font_size(block, base_size)
        height += estimate_wrapped_lines(text, font_size, width_pt) * font_size * 1.28
        height += 7
    return height


def split_text_by_height(text: str, width_pt: float, max_height_pt: float, font_size: int) -> list[str]:
    chunks: list[str] = []
    current_lines: list[str] = []

    def current_height(lines: list[str]) -> float:
        return estimate_wrapped_lines("\n".join(lines), font_size, width_pt) * font_size * 1.28 + 7

    for raw_line in split_text_into_sentences(text):
        line = raw_line.strip()
        if not line:
            continue
        pieces: list[str] = []
        chars_per_piece = max(int(width_pt / (font_size * 0.92) * 5), 35)
        if current_height([line]) <= max_height_pt:
            pieces.append(line)
        else:
            for start in range(0, len(line), chars_per_piece):
                pieces.append(line[start : start + chars_per_piece])

        for piece in pieces:
            candidate = [*current_lines, piece]
            if current_lines and current_height(candidate) > max_height_pt:
                chunks.append("\n".join(current_lines))
                current_lines = [piece]
            else:
                current_lines = candidate

    if current_lines:
        chunks.append("\n".join(current_lines))
    return chunks


def split_text_into_sentences(text: str) -> list[str]:
    parts: list[str] = []
    current = ""
    for char in text:
        current += char
        if char in "。！？!?；;" or char in "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬" and current.strip()[:-1]:
            parts.append(current.strip())
            current = ""
    if current.strip():
        parts.append(current.strip())
    return parts


def split_block_by_capacity(block: dict, width_pt: float, max_height_pt: float, base_size: int = 18) -> list[dict]:
    if estimate_blocks_height([block], width_pt, base_size) <= max_height_pt:
        return [block]
    span = first_span(block) or {}
    font_size = block_font_size(block, base_size)
    return [
        {"bbox": block["bbox"], "lines": block.get("lines", []), "_text": chunk, "_span": span}
        for chunk in split_text_by_height(block_text(block), width_pt, max_height_pt, font_size)
    ]


def split_segments_by_capacity(
    segments: list[list[dict]],
    width_pt: float,
    max_height_pt: float,
    base_size: int = 18,
) -> list[list[dict]]:
    output: list[list[dict]] = []
    for segment in segments:
        current: list[dict] = []
        for block in segment:
            for chunk in split_block_by_capacity(block, width_pt, max_height_pt, base_size):
                candidate = [*current, chunk]
                if current and estimate_blocks_height(candidate, width_pt, base_size) > max_height_pt:
                    output.append(current)
                    current = [chunk]
                else:
                    current = candidate
        if current:
            output.append(current)
    return output


def merge_short_segments(
    segments: list[list[dict]],
    width_pt: float,
    max_height_pt: float,
    base_size: int = 15,
    min_chars: int = 90,
) -> list[list[dict]]:
    merged: list[list[dict]] = []
    for segment in segments:
        segment_chars = blocks_text_length(segment)
        if (
            merged
            and segment_chars < min_chars
            and estimate_blocks_height([*merged[-1], *segment], width_pt, base_size) <= max_height_pt * 1.25
        ):
            merged[-1].extend(segment)
        else:
            merged.append(segment)
    if len(merged) < 2:
        return merged

    output: list[list[dict]] = []
    index = 0
    while index < len(merged):
        segment = merged[index]
        segment_text = "\n".join(block_text(block) for block in segment)
        is_placeholder_only = "图表占位" in segment_text and len(segment_text.strip()) <= 24
        if (
            index + 1 < len(merged)
            and blocks_text_length(segment) < min_chars
            and estimate_blocks_height([*segment, *merged[index + 1]], width_pt, base_size) <= max_height_pt * 1.25
        ):
            output.append([*segment, *merged[index + 1]])
            index += 2
        elif is_placeholder_only and output:
            output[-1].extend(segment)
            index += 1
        else:
            output.append(segment)
            index += 1
    return output


def merge_short_slide_specs(
    specs: list[tuple[list[dict], list[dict]]],
    width_pt: float,
    max_height_pt: float,
    base_size: int = 15,
    min_chars: int = 90,
) -> list[tuple[list[dict], list[dict]]]:
    merged: list[tuple[list[dict], list[dict]]] = []
    for segment, images in specs:
        segment_text = "\n".join(block_text(block) for block in segment).strip()
        segment_chars = len(segment_text)
        is_short = segment_chars < min_chars
        is_placeholder_only = "图表占位" in segment_text and segment_chars <= 30
        if merged and (is_short or is_placeholder_only):
            previous_segment, previous_images = merged[-1]
            candidate = [*previous_segment, *segment]
            capacity_multiplier = 1.7 if is_short else 1.25
            max_merged_chars = 900 if is_short else 780
            if blocks_text_length(candidate) <= max_merged_chars and estimate_blocks_height(candidate, width_pt, base_size) <= max_height_pt * capacity_multiplier:
                merged[-1] = (candidate, [*previous_images, *images])
                continue
        merged.append((segment, images))
    return merged


def split_block_for_rebalance(block: dict, max_prefix_chars: int) -> tuple[dict | None, dict | None]:
    text = block_text(block)
    if len(text) <= max_prefix_chars:
        return block, None

    lines = text.splitlines()
    prefix_lines: list[str] = []
    suffix_lines: list[str] = []
    current_chars = 0
    for index, line in enumerate(lines):
        line_len = len(line)
        if prefix_lines and current_chars + line_len > max_prefix_chars:
            suffix_lines = lines[index:]
            break
        if not prefix_lines and line_len > max_prefix_chars:
            split_at = max_prefix_chars
            prefix_lines = [line[:split_at]]
            suffix_lines = [line[split_at:], *lines[index + 1 :]]
            break
        prefix_lines.append(line)
        current_chars += line_len

    if not suffix_lines and len("\n".join(prefix_lines)) == len(text):
        return block, None

    prefix_text = "\n".join(line for line in prefix_lines if line).strip()
    suffix_text = "\n".join(line for line in suffix_lines if line).strip()
    prefix = make_text_block_like(block, prefix_text) if prefix_text else None
    suffix = make_text_block_like(block, suffix_text) if suffix_text else None
    return prefix, suffix


def rebalance_short_slide_specs(
    specs: list[tuple[list[dict], list[dict]]],
    width_pt: float,
    max_height_pt: float,
    base_size: int = 15,
    min_chars: int = 380,
    target_chars: int = 650,
    max_chars: int = 780,
) -> list[tuple[list[dict], list[dict]]]:
    balanced = [(list(segment), list(images)) for segment, images in specs]
    index = 0
    while index + 1 < len(balanced):
        segment, images = balanced[index]
        next_segment, next_images = balanced[index + 1]
        current_text = "\n".join(block_text(block) for block in segment).strip()
        next_text = "\n".join(block_text(block) for block in next_segment).strip()
        current_tail = current_text[-1:] if current_text else ""
        next_head = next_text[:1]
        current_incomplete = current_tail not in "。！？!?：:；;）)】》”’…"
        next_continues = next_head and next_head not in "【一二三四五六七八九十①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬任务"
        should_rebalance = current_incomplete and next_continues and blocks_text_length(segment) < min_chars

        if images or not next_segment or not should_rebalance:
            index += 1
            continue

        moved_any = False
        while next_segment and blocks_text_length(segment) < target_chars:
            remaining_capacity = max_chars - blocks_text_length(segment)
            if remaining_capacity <= 80:
                break
            prefix, suffix = split_block_for_rebalance(next_segment[0], remaining_capacity)
            if prefix is None:
                break
            candidate = [*segment, prefix]
            if estimate_blocks_height(candidate, width_pt, base_size) > max_height_pt * 1.25:
                break
            segment = candidate
            moved_any = True
            if suffix is None:
                next_segment = next_segment[1:]
            else:
                next_segment = [suffix, *next_segment[1:]]
                break

        if moved_any:
            balanced[index] = (segment, images)
            if next_segment:
                balanced[index + 1] = (next_segment, next_images)
            else:
                if next_images:
                    balanced[index] = (segment, [*images, *next_images])
                del balanced[index + 1]
                continue
        index += 1
    return balanced


def split_long_segments(segments: list[list[dict]], max_chars: int = 850) -> list[list[dict]]:
    output: list[list[dict]] = []
    for segment in segments:
        current: list[dict] = []
        current_chars = 0
        for block in split_oversized_blocks(segment, max_chars):
            text_len = len(block_text(block))
            if current and current_chars + text_len > max_chars:
                output.append(current)
                current = []
                current_chars = 0
            current.append(block)
            current_chars += text_len
        if current:
            output.append(current)
    return output


def split_text_by_size(text: str, max_chars: int) -> list[str]:
    chunks: list[str] = []
    current = ""
    pieces: list[str] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        start = 0
        while start < len(line):
            pieces.append(line[start : start + max_chars])
            start += max_chars

    for piece in pieces:
        separator = "\n" if current else ""
        if current and len(current) + len(separator) + len(piece) > max_chars:
            chunks.append(current)
            current = piece
        else:
            current = f"{current}{separator}{piece}" if current else piece
    if current:
        chunks.append(current)
    return chunks


def split_oversized_blocks(blocks: list[dict], max_chars: int) -> list[dict]:
    output: list[dict] = []
    for block in blocks:
        text = block_text(block)
        if len(text) <= max_chars:
            output.append(block)
            continue
        span = first_span(block) or {}
        for chunk in split_text_by_size(text, max_chars):
            output.append({"bbox": block["bbox"], "lines": block.get("lines", []), "_text": chunk, "_span": span})
    return output


def first_span(block: dict) -> dict | None:
    if "_span" in block:
        return block["_span"]

    for line in block.get("lines", []):
        for span in line.get("spans", []):
            if span.get("text", "").strip():
                return span
    return None


def choose_font_size(chars: int) -> int:
    if chars <= 180:
        return 22
    if chars <= 320:
        return 18
    return 15


def add_clean_text_slide(slide, segment: list[dict], slide_width_pt: float, slide_height_pt: float, images: list[dict] | None = None) -> None:
    images = images or []
    margin_x = CONTENT_MARGIN_X_PT
    margin_y = CONTENT_MARGIN_Y_PT
    image_area_height = 0
    if images:
        image_area_height = min(IMAGE_AREA_HEIGHT_PT, slide_height_pt * 0.36)
    box = slide.shapes.add_textbox(
        pt_to_emu(margin_x),
        pt_to_emu(margin_y),
        pt_to_emu(slide_width_pt - margin_x * 2),
        pt_to_emu(slide_height_pt - margin_y * 2 - image_area_height),
    )
    frame = box.text_frame
    frame.clear()
    frame.margin_left = pt_to_emu(TEXT_INSET_X_PT)
    frame.margin_right = pt_to_emu(TEXT_INSET_X_PT)
    frame.margin_top = pt_to_emu(TEXT_INSET_Y_PT)
    frame.margin_bottom = pt_to_emu(TEXT_INSET_Y_PT)
    frame.word_wrap = True
    frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    frame.vertical_anchor = MSO_ANCHOR.TOP

    base_size = choose_font_size(blocks_text_length(segment))
    first_paragraph = True
    previous_paragraph = None
    for block in segment:
        text = block_text(block)
        if not text:
            continue
        if previous_paragraph is not None and should_merge_text_blocks(previous_paragraph.text.strip(), text):
            paragraph = previous_paragraph
        else:
            paragraph = frame.paragraphs[0] if first_paragraph else frame.add_paragraph()
            first_paragraph = False
            paragraph.alignment = PP_ALIGN.LEFT
            paragraph.space_before = Pt(0)
            paragraph.space_after = Pt(5)
            paragraph.line_spacing = 1.0
            previous_paragraph = paragraph

        span = first_span(block) or {}
        run = paragraph.add_run()
        run.text = text
        font = run.font
        source_size = span.get("size", 10)
        is_heading = source_size >= 13 or span_is_bold(span)
        font.name = clean_font_name(span.get("font", "Arial"))
        font.size = Pt(base_size + 3 if is_heading and len(text) < 90 else base_size)
        font.bold = is_heading
        font.italic = span_is_italic(span)
        font.color.rgb = int_to_rgb(span.get("color", 0))

    add_images_to_slide(slide, images, slide_width_pt, slide_height_pt, margin_x, margin_y)


def add_images_to_slide(slide, image_blocks: list[dict], slide_width_pt: float, slide_height_pt: float, margin_x: float, margin_y: float) -> None:
    if not image_blocks:
        return

    gap = 12
    max_total_width = slide_width_pt - margin_x * 2
    max_height = min(175, slide_height_pt * 0.32)
    slots = len(image_blocks)
    slot_width = (max_total_width - gap * (slots - 1)) / slots
    rendered: list[tuple[dict, float, float]] = []

    for image_block in image_blocks:
        image_bytes = image_block.get("image")
        if not image_bytes:
            continue
        width = image_block.get("width") or 1
        height = image_block.get("height") or 1
        scale = min(slot_width / width, max_height / height, 1.0)
        rendered.append((image_block, width * scale, height * scale))

    if not rendered:
        return

    total_width = sum(width for _, width, _ in rendered) + gap * (len(rendered) - 1)
    left = (slide_width_pt - total_width) / 2
    top = slide_height_pt - margin_y - max(height for _, _, height in rendered)
    for image_block, width, height in rendered:
        slide.shapes.add_picture(
            io.BytesIO(image_block["image"]),
            pt_to_emu(left),
            pt_to_emu(top + (max_height - height) / 2),
            width=pt_to_emu(width),
            height=pt_to_emu(height),
        )
        left += width + gap


def detect_answer_start_page(doc: fitz.Document, fallback_page: int | None = 29) -> int | None:
    for index, page in enumerate(doc, 1):
        if index < 20:
            continue
        text = page.get_text()
        if "巩固训练答案" in text or text.count("答案") >= 2:
            return index
    return fallback_page


def convert_pdf_to_pptx(pdf_path: Path, pptx_path: Path, skip_from_page: int | None = 29, layout: str = "clean") -> None:
    doc = fitz.open(pdf_path)
    if doc.page_count == 0:
        raise ValueError(f"PDF has no pages: {pdf_path}")

    slide_width_pt = 13.333333 * PDF_POINTS_PER_INCH
    slide_height_pt = 7.5 * PDF_POINTS_PER_INCH

    presentation = Presentation()
    presentation.slide_width = Inches(13.333333)
    presentation.slide_height = Inches(7.5)
    blank_layout = presentation.slide_layouts[6]

    if skip_from_page == 0:
        skip_from_page = detect_answer_start_page(doc)

    slide_specs: list[tuple[list[dict], list[dict]]] = []

    for page_number, page in enumerate(doc, 1):
        if skip_from_page is not None and page_number >= skip_from_page:
            break
        page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_LIGATURES | fitz.TEXT_PRESERVE_WHITESPACE)
        image_page_dict = page.get_text("dict")
        text_blocks = [block for block in page_dict.get("blocks", []) if block.get("type") == 0 and has_text(block)]
        text_blocks = replace_table_text_with_placeholders(text_blocks)
        text_blocks = merge_adjacent_text_blocks(text_blocks)
        image_blocks = [block for block in image_page_dict.get("blocks", []) if block.get("type") == 1]
        has_images = bool(image_blocks)
        content_width_pt = slide_width_pt - CONTENT_MARGIN_X_PT * 2 - TEXT_INSET_X_PT * 2
        visible_height_pt = (
            slide_height_pt
            - CONTENT_MARGIN_Y_PT * 2
            - TEXT_INSET_Y_PT * 2
            - (IMAGE_AREA_HEIGHT_PT if has_images else 0)
        )
        content_height_pt = visible_height_pt * CAPACITY_SAFETY
        segments = split_segments_by_capacity(
            split_blocks_by_vertical_gaps(text_blocks),
            width_pt=content_width_pt,
            max_height_pt=content_height_pt,
            base_size=15,
        )
        segments = merge_short_segments(segments, content_width_pt, content_height_pt, base_size=15)

        for segment_index, segment in enumerate(segments):
            slide_specs.append((segment, image_blocks if segment_index == len(segments) - 1 else []))

    slide_specs = merge_short_slide_specs(
        slide_specs,
        width_pt=slide_width_pt - CONTENT_MARGIN_X_PT * 2 - TEXT_INSET_X_PT * 2,
        max_height_pt=(slide_height_pt - CONTENT_MARGIN_Y_PT * 2 - TEXT_INSET_Y_PT * 2) * CAPACITY_SAFETY,
        base_size=15,
    )
    slide_specs = rebalance_short_slide_specs(
        slide_specs,
        width_pt=slide_width_pt - CONTENT_MARGIN_X_PT * 2 - TEXT_INSET_X_PT * 2,
        max_height_pt=(slide_height_pt - CONTENT_MARGIN_Y_PT * 2 - TEXT_INSET_Y_PT * 2) * CAPACITY_SAFETY,
        base_size=15,
        min_chars=380,
        target_chars=520,
        max_chars=760,
    )
    slide_specs = merge_short_slide_specs(
        slide_specs,
        width_pt=slide_width_pt - CONTENT_MARGIN_X_PT * 2 - TEXT_INSET_X_PT * 2,
        max_height_pt=(slide_height_pt - CONTENT_MARGIN_Y_PT * 2 - TEXT_INSET_Y_PT * 2) * CAPACITY_SAFETY,
        base_size=15,
    )

    for segment, images in slide_specs:
            slide = presentation.slides.add_slide(blank_layout)
            if layout == "clean":
                add_clean_text_slide(
                    slide,
                    segment,
                    slide_width_pt,
                    slide_height_pt,
                    images,
                )
                continue

            x0, y0, x1, y1 = segment_bbox(segment, page.rect)
            segment_width = max(x1 - x0, 1)
            segment_height = max(y1 - y0, 1)
            slide_margin_x = 34
            slide_margin_y = 28
            scale = min(
                (slide_width_pt - slide_margin_x * 2) / segment_width,
                (slide_height_pt - slide_margin_y * 2) / segment_height,
            )
            scale = min(scale, 2.2)
            used_width = segment_width * scale
            used_height = segment_height * scale
            margin_left = max(slide_margin_x, (slide_width_pt - used_width) / 2)
            margin_top = max(slide_margin_y, (slide_height_pt - used_height) / 2)

            for block in segment:
                add_text_block(slide, block, scale, x0, y0, margin_left, margin_top)

    presentation.save(pptx_path)


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert copyable PDF text into an editable PPTX.")
    parser.add_argument("pdf", nargs="?", default="mypaper.pdf", help="Input PDF path")
    parser.add_argument("pptx", nargs="?", default="mypaper.pptx", help="Output PPTX path")
    parser.add_argument("--skip-from-page", type=int, default=29, help="Skip this PDF page and all following pages. Use 0 to auto-detect, -1 to keep all pages.")
    parser.add_argument("--layout", choices=("clean", "positioned"), default="clean", help="clean avoids overlap; positioned preserves original coordinates.")
    args = parser.parse_args()

    skip_from_page = None if args.skip_from_page < 0 else args.skip_from_page
    convert_pdf_to_pptx(Path(args.pdf), Path(args.pptx), skip_from_page=skip_from_page, layout=args.layout)


if __name__ == "__main__":
    main()
