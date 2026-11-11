"""
Image Analyzer - AI-Powered Image and PDF Analysis Tool
Supports: Local LM Studio and OpenRouter multimodal models
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import json
import base64
import csv
import io
import re
import datetime
import subprocess
from pathlib import Path
from PIL import Image, ExifTags
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from pypdf import PdfReader
from tkinterdnd2 import DND_FILES, TkinterDnD

# ── Config ────────────────────────────────────────────────────────────────────
LMSTUDIO_BASE  = os.environ.get("IMAGE_ANALYZER_LOCAL_URL", "http://127.0.0.1:8080/v1")
LMSTUDIO_MODEL = "Qwen3.6-35B-A3B"
LLAMA_START_SCRIPT = Path(os.environ.get("IMAGE_ANALYZER_LLAMA_START", "D:/AI/llama.cpp/Start-Llama.ps1"))
LOCAL_MODELS = {
    "qwen": (LMSTUDIO_BASE, LLAMA_START_SCRIPT, ["-NoBrowser"]),
    "gemma": ("http://127.0.0.1:8081/v1", Path("D:/AI/llama.cpp/Start-Gemma4-E4B.ps1"), []),
}
OPENROUTER_BASE = "https://openrouter.ai/api/v1"
OPENROUTER_KEY  = (os.environ.get("OPENROUTER_KEY") or
                   os.environ.get("OPENROUTER_API_KEY", ""))
GEOCODING_KEY   = os.environ.get("GEOCODING_KEY", "")

OPENROUTER_MODELS = [
    {
        "label": "Gemini 3.1 Flash Lite",
        "id": "google/gemini-3.1-flash-lite",
        "price": "$0.25/M in / $1.50/M out",
        "image": True,
    },
    {
        "label": "Qwen3 VL 8B Instruct",
        "id": "qwen/qwen3-vl-8b-instruct",
        "price": "$0.08/M in / $0.50/M out",
        "image": True,
    },
    {
        "label": "Qwen2.5 VL 72B Instruct",
        "id": "qwen/qwen2.5-vl-72b-instruct",
        "price": "$0.25/M in / $0.75/M out",
        "image": True,
    },
    {
        "label": "GPT-5.6 Luna",
        "id": "openai/gpt-5.6-luna",
        "price": "$1.00/M in / $6.00/M out",
        "image": True,
    },
    {
        "label": "Kimi K3",
        "id": "moonshotai/kimi-k3",
        "price": "$3.00/M in / $15.00/M out",
        "image": True,
    },
    {
        "label": "Tencent Hy3 (PDF text only)",
        "id": "tencent/hy3",
        "price": "$0.14/M in / $0.58/M out",
        "image": False,
    },
    {
        "label": "Nemotron 3 Ultra 550B (PDF text only)",
        "id": "nvidia/nemotron-3-ultra-550b-a55b",
        "price": "$0.50/M in / $2.20/M out",
        "image": False,
    },
]

IMAGE_EXT = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".tiff", ".tif"}
PDF_EXT = {".pdf"}
SUPPORTED_EXT = IMAGE_EXT | PDF_EXT

PDF_ENGINES = {
    "Auto (native when available)": None,
    "Cloudflare AI (free text parser)": "cloudflare-ai",
    "Mistral OCR (best for scans; extra cost)": "mistral-ocr",
    "Native model processing": "native",
}

RESULT_COLUMNS = [
    ("Filename", "Filename", 150, False),
    ("Description", "Description", 280, True),
    ("Text Content", "Text Content", 250, True),
    ("Translation", "Translation", 220, True),
    ("Path", "Path", 280, True),
    ("Dimensions", "Dims / Pages", 120, False),
    ("Color Mode", "Mode", 90, False),
    ("File Size", "Size", 90, False),
    ("Photographer", "Photographer", 150, False),
    ("Location", "Location", 200, True),
    ("GPS Coordinates", "GPS Coordinates", 150, False),
    ("Date Taken", "Date Taken", 140, False),
    ("Subject Tags", "Subject Tags", 190, True),
    ("Quality Notes", "Quality Notes", 220, True),
    ("Metadata", "Metadata", 280, True),
]
RESULTS_WIDTH = sum(column[2] for column in RESULT_COLUMNS)
RESULT_ROW_HEIGHT = 76

def build_analysis_prompt(translate_text=False):
    translation_rule = (
        '  "translation": "English translation of visible non-English text, '
        'preserving line breaks where useful, or null",\n'
        if translate_text else
        '  "translation": null,\n'
    )
    translation_instruction = (
        "If visible text is in any non-English language, translate it into "
        "natural English in the translation field. Keep text_content as the "
        "original visible text/OCR, without translating it. "
        if translate_text else
        "Do not translate visible text; set translation to JSON null. "
    )
    return (
    "Analyze this image and respond with ONLY a single valid JSON object. "
    "No markdown, no code fences, no explanation before or after. "
    "Keep every string value under 900 characters. "
    "Preserve visible text exactly when readable; summarize only if there is too much. "
    "Always fill description, subject_tags, colors, and quality_notes with your best visual analysis. "
    "Do not leave those fields blank unless the image cannot be inspected. "
    f"{translation_instruction}"
    "Extract metadata only if present (EXIF or visible image content). Include GPS coordinates if available. "
    "Do not invent a location from coordinates; set location to null unless a place name is visible or embedded metadata provides it. "
    "Use JSON null (not the word null in quotes) for missing data. "
    "The JSON must have exactly these 10 keys:\n"
    "{\n"
    '  "description": "clear, information-dense description (subject, setting, action, notable details) under 400 chars",\n'
    '  "text_content": "visible text/OCR from the image, preserving line breaks where useful, or null",\n'
    f"{translation_rule}"
    '  "photographer": null,\n'
    '  "location": null,\n'
    '  "gps_coordinates": "latitude, longitude or null",\n'
    '  "date_taken": null,\n'
    '  "subject_tags": "5–10 specific tags (objects, scene type, activity, style)",\n'
    '  "colors": "3–6 dominant colors if useful for search/filtering, otherwise null",\n'
    '  "quality_notes": "brief lighting and composition note"\n'
    "}"
)


ANALYSIS_PROMPT = build_analysis_prompt(False)

PDF_DOCUMENT_PROMPT = """
Reconstruct this PDF as a structured, editable document. Return ONLY one valid
JSON object, without markdown fences or commentary. Preserve wording, reading
order, headings, lists, tables, page boundaries, and visible emphasis. Do not
summarize or invent content. The exact schema is:
{
  "title": "document title or empty string",
  "subtitle": "subtitle or empty string",
  "style": {
    "page_width_inches": 8.5,
    "page_height_inches": 11.0,
    "margin_top_inches": 0.75,
    "margin_right_inches": 0.75,
    "margin_bottom_inches": 0.75,
    "margin_left_inches": 0.75,
    "base_font": "Arial",
    "base_font_size": 10.5,
    "title_color": "000000",
    "heading_color": "000000"
  },
  "pages": [
    {
      "page_number": 1,
      "blocks": [
        {"type": "heading", "level": 1, "text": "Heading", "alignment": "left"},
        {"type": "paragraph", "text": "Paragraph text", "alignment": "left",
         "runs": [{"text": "optional styled span", "bold": false,
                    "italic": false, "underline": false}]},
        {"type": "bullet_list", "items": ["Item one", "Item two"]},
        {"type": "numbered_list", "items": ["First", "Second"]},
        {"type": "table", "headers": ["Column 1"], "rows": [["Value"]]},
        {"type": "note", "text": "Visually distinct note or callout"}
      ]
    }
  ]
}
Use only the listed block types. The runs field is optional; when used, its text
must collectively replace the block text without duplication. Put each table
row in its own array with the same number of cells as headers. Keep page breaks
faithful to the source. Style values should approximate the visible PDF, while
remaining readable in Word.
""".strip()

# ── Palette ───────────────────────────────────────────────────────────────────
BG      = "#0f1720"
BG2     = "#17212c"
BG3     = "#243241"
ACCENT  = "#2dd4bf"
ACCENT2 = "#67e8d7"
SUCCESS = "#86efac"
WARN    = "#facc15"
ERR     = "#fecaca"
TEXT    = "#edf4f8"
MUTED   = "#a9b8c6"
BORDER  = "#354657"
ERR_BG  = "#38252d"
ERR_HOVER = "#49303a"


# ── Helpers ───────────────────────────────────────────────────────────────────

def human_size(n):
    for unit in ("B", "KB", "MB", "GB"):
        if n < 1024:
            return f"{n:.1f} {unit}"
        n /= 1024
    return f"{n:.1f} TB"


def _plain_meta_value(value):
    if isinstance(value, bytes):
        return f"<{len(value)} bytes>"
    if isinstance(value, tuple):
        return [_plain_meta_value(v) for v in value]
    if isinstance(value, list):
        return [_plain_meta_value(v) for v in value]
    if isinstance(value, dict):
        return {str(k): _plain_meta_value(v) for k, v in value.items()}
    try:
        json.dumps(value)
        return value
    except TypeError:
        return str(value)


def _stringify_metadata_keys(value):
    if isinstance(value, dict):
        return {
            str(k): _stringify_metadata_keys(v)
            for k, v in value.items()
        }
    if isinstance(value, list):
        return [_stringify_metadata_keys(v) for v in value]
    if isinstance(value, tuple):
        return [_stringify_metadata_keys(v) for v in value]
    return value


def _gps_decimal(gps):
    try:
        def pick(name, num):
            return gps.get(name) if name in gps else gps.get(num)

        def dms(vals):
            d, m, s = [float(x) for x in vals]
            return d + m / 60 + s / 3600

        lat = dms(pick("GPSLatitude", 2))
        lon = dms(pick("GPSLongitude", 4))
        if pick("GPSLatitudeRef", 1) == "S":
            lat = -lat
        if pick("GPSLongitudeRef", 3) == "W":
            lon = -lon
        return f"{lat:.6f}, {lon:.6f}"
    except Exception:
        return ""


def _split_gps_coordinates(value):
    match = re.search(r"(-?\d+(?:\.\d+)?)\s*,\s*(-?\d+(?:\.\d+)?)", value or "")
    if not match:
        return None
    try:
        lat, lon = float(match.group(1)), float(match.group(2))
    except ValueError:
        return None
    if not (-90 <= lat <= 90 and -180 <= lon <= 180):
        return None
    return lat, lon


def reverse_geocode(gps_coordinates, api_key, cache=None):
    coords = _split_gps_coordinates(gps_coordinates)
    api_key = (api_key or "").strip()
    if not coords or not api_key:
        return ""
    lat, lon = coords
    cache_key = f"{lat:.6f},{lon:.6f}"
    if cache is not None and cache_key in cache:
        return cache[cache_key]
    try:
        r = requests.get(
            "https://maps.googleapis.com/maps/api/geocode/json",
            params={"latlng": cache_key, "key": api_key, "language": "en"},
            timeout=12,
        )
        r.raise_for_status()
        data = r.json()
        if data.get("status") != "OK":
            location = ""
        else:
            results = data.get("results") or []
            location = (results[0].get("formatted_address") if results else "") or ""
    except Exception:
        location = ""
    if cache is not None:
        cache[cache_key] = location
    return location


def _clean_metadata_json(data):
    return json.dumps(
        _stringify_metadata_keys(data),
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )


def get_exif_data(path, img=None):
    meta = {"photographer": "", "date_taken": "", "gps": "", "details": {}}
    try:
        close_img = img is None
        if img is None:
            img = Image.open(path)
        details = {
            "file": {
                "format": img.format or Path(path).suffix.lower().lstrip("."),
                "dimensions_px": [img.width, img.height],
                "mode": img.mode,
                "size_bytes": os.path.getsize(path),
            }
        }
        if img.info:
            details["image_info"] = {
                str(k): _plain_meta_value(v) for k, v in img.info.items()
                if k.lower() not in {"exif", "icc_profile"}
            }
            if "dpi" in img.info:
                details["file"]["dpi"] = _plain_meta_value(img.info.get("dpi"))
            if "icc_profile" in img.info:
                details["color"] = {
                    "icc_profile_bytes": len(img.info.get("icc_profile") or b"")
                }

        exif_obj = img.getexif()
        exif_raw = dict(exif_obj.items()) if exif_obj else {}
        if not exif_raw and hasattr(img, "_getexif"):
            exif_raw = img._getexif() or {}
        if exif_raw:
            exif = {ExifTags.TAGS.get(k, k): _plain_meta_value(v)
                    for k, v in exif_raw.items()}
            ifd_tags = getattr(ExifTags, "IFD", None)
            for ifd_name, ifd_id in (
                ("ExifIFD", getattr(ifd_tags, "Exif", None)),
                ("GPSInfo", getattr(ifd_tags, "GPSInfo", None)),
            ):
                if ifd_id is None:
                    continue
                try:
                    ifd = exif_obj.get_ifd(ifd_id)
                except Exception:
                    ifd = {}
                if ifd:
                    tag_map = ExifTags.GPSTAGS if ifd_name == "GPSInfo" else ExifTags.TAGS
                    exif[ifd_name] = {
                        tag_map.get(k, k): _plain_meta_value(v) for k, v in ifd.items()
                    }
            gps = exif.get("GPSInfo")
            if gps and isinstance(gps, dict):
                gps = {ExifTags.GPSTAGS.get(k, k): _plain_meta_value(v)
                       for k, v in gps.items()}
                exif["GPSInfo"] = gps
            details["exif"] = exif
            exif_ifd = exif.get("ExifIFD") if isinstance(exif.get("ExifIFD"), dict) else {}
            meta["photographer"] = str(exif.get("Artist") or exif.get("Copyright") or "")
            meta["date_taken"] = str(
                exif_ifd.get("DateTimeOriginal") or exif.get("DateTimeOriginal") or
                exif.get("DateTime") or ""
            )
            if isinstance(gps, dict):
                meta["gps"] = _gps_decimal(gps)

        meta["details"] = details
        if close_img:
            img.close()
    except Exception:
        pass
    return meta


def build_metadata_column(exif, ai):
    details = dict(exif.get("details") or {})
    ai_colors = (ai.get("colors") or "").strip()
    if ai_colors:
        details.setdefault("visual", {})["dominant_colors"] = ai_colors
    text = _clean_metadata_json(details)
    return text[:32000]


# Excel cannot store a cell longer than 32767 characters, and the Windows
# clipboard/Tk both degrade on enormous blobs. Keep values plain strings,
# strip control characters openpyxl rejects, and cap length so every export
# (Excel, CSV, copy-paste) succeeds reliably even for very heavy rows.
EXPORT_CELL_LIMIT = 32000
_EXCEL_BAD_CHARS = set(chr(c) for c in range(32)) - {"\t", "\n", "\r"}


def _export_value(value):
    """Coerce a result value to a clean, bounded plain string for export."""
    if value is None:
        return ""
    if not isinstance(value, str):
        value = str(value)
    if _EXCEL_BAD_CHARS:
        value = "".join(ch for ch in value if ch not in _EXCEL_BAD_CHARS)
    if len(value) > EXPORT_CELL_LIMIT:
        value = value[:EXPORT_CELL_LIMIT] + "\n…[truncated—full result available in the app window]"
    return value


def _export_dict(row):
    """Return a copy of a result row whose export column values are clean."""
    out = {}
    for key in EXPORT_COLS:
        out[key] = _export_value(row.get(key, ""))
    return out


EXPORT_COLS = ["Filename", "Description", "Text Content", "Translation",
               "Path", "Dimensions", "Color Mode", "File Size", "Photographer",
               "Location", "GPS Coordinates", "Date Taken", "Subject Tags",
               "Quality Notes", "Metadata"]


def encode_image(path):
    with open(path, "rb") as f:
        data = f.read()
    ext = Path(path).suffix.lower().lstrip(".")
    mime_map = {"jpg": "image/jpeg", "jpeg": "image/jpeg", "png": "image/png",
                "gif": "image/gif", "bmp": "image/bmp", "webp": "image/webp",
                "tiff": "image/tiff", "tif": "image/tiff"}
    return base64.b64encode(data).decode(), mime_map.get(ext, "image/jpeg")


def encode_pdf(path):
    with open(path, "rb") as pdf_file:
        return base64.b64encode(pdf_file.read()).decode()


def get_pdf_info(path):
    info = {"pages": 0, "metadata": {}}
    try:
        reader = PdfReader(path)
        info["pages"] = len(reader.pages)
        info["metadata"] = {
            str(key).lstrip("/"): str(value)
            for key, value in (reader.metadata or {}).items()
            if value is not None
        }
    except Exception:
        pass
    return info


# ── Robust JSON parser ────────────────────────────────────────────────────────

def _try_load(text):
    """Attempt json.loads; return dict or None."""
    try:
        obj = json.loads(text)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass
    return None


def _fix_truncated(text):
    """
    Best-effort repair of a truncated JSON object by closing any open
    string then appending missing closing brackets/braces.
    """
    text = text.rstrip().rstrip(",")
    if text.count('"') % 2 != 0:
        text += '"'
    text += "]" * max(0, text.count("[") - text.count("]"))
    text += "}" * max(0, text.count("{") - text.count("}"))
    return text


def parse_ai_response(raw):
    """
    Robustly extract a JSON dict from model output that may be:
    - clean JSON
    - JSON wrapped in ```json ... ```
    - JSON preceded/followed by prose
    - Truncated JSON (token limit hit mid-string)
    - Python-style None/True/False instead of null/true/false
    """
    if not raw:
        return _empty_ai()

    # 1. Strip markdown fences
    text = re.sub(r"```(?:json)?", "", raw).strip("`").strip()

    # 2. Normalise Python literals that some models emit
    text = re.sub(r'\bNone\b',  'null',  text)
    text = re.sub(r'\bTrue\b',  'true',  text)
    text = re.sub(r'\bFalse\b', 'false', text)

    # 3. Try the whole cleaned string as-is
    result = _try_load(text)
    if result:
        return _sanitise(result)

    # 4. Find the outermost { ... } block
    # Use a brace-counting scan so nested braces are handled correctly
    start = text.find("{")
    if start != -1:
        depth = 0
        end   = start
        for i, ch in enumerate(text[start:], start):
            if ch == "{":
                depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    end = i
                    break
        candidate = text[start: end + 1]
        result = _try_load(candidate)
        if result:
            return _sanitise(result)

        # 5. Repair and retry
        repaired = _fix_truncated(candidate)
        result = _try_load(repaired)
        if result:
            return _sanitise(result)

        # 6. If braces were never balanced, try from { to end of string
        partial = text[start:]
        repaired = _fix_truncated(partial)
        result = _try_load(repaired)
        if result:
            return _sanitise(result)

    # 7. Last resort: surface the raw text so nothing is silently lost
    return {**_empty_ai(), "description": raw[:500]}


def _empty_ai():
    return {"description": "", "text_content": "", "translation": "",
            "photographer": "", "location": "", "gps_coordinates": "",
            "date_taken": "", "subject_tags": "", "colors": "",
            "quality_notes": ""}


def _sanitise(d):
    """Ensure all expected keys exist and values are plain, bounded strings."""
    out = _empty_ai()
    for key in out:
        val = d.get(key)
        if val is None:
            out[key] = ""
        else:
            # Never pass dicts/lists to Excel
            if not isinstance(val, str):
                val = json.dumps(val) if isinstance(val, (dict, list)) else str(val)
            
            # Strip brackets and quotes for specific fields
            if key in ["subject_tags", "colors"]:
                # Remove brackets and quotes
                val = re.sub(r'[\[\]"\'\s]+', ' ', val)
                # Replace multiple commas/spaces with a single comma and space
                val = re.sub(r'\s*,\s*', ', ', val)
                val = val.strip(', ')

            limit = 2400 if key in ("text_content", "translation") else 900
            out[key] = val[:limit]
    return out


def _weak_ai_result(ai):
    required = ("description", "subject_tags", "quality_notes")
    if any(not (ai.get(k) or "").strip() for k in required):
        return True
    joined = " ".join((ai.get(k) or "") for k in required).lower()
    if len(joined.strip()) < 40:
        return True
    repeated = re.search(r"\b(\w{3,})\b(?:\W+\1\b){4,}", joined)
    return bool(repeated)


def _extract_json_object(raw):
    """Extract a JSON object from clean, fenced, or prose-wrapped output."""
    if not raw:
        return None
    text = re.sub(r"```(?:json)?", "", raw).strip("`").strip()
    direct = _try_load(text)
    if direct:
        return direct
    start = text.find("{")
    if start < 0:
        return None
    depth = 0
    in_string = False
    escaped = False
    for pos, char in enumerate(text[start:], start):
        if escaped:
            escaped = False
            continue
        if char == "\\" and in_string:
            escaped = True
            continue
        if char == '"':
            in_string = not in_string
        elif not in_string and char == "{":
            depth += 1
        elif not in_string and char == "}":
            depth -= 1
            if depth == 0:
                return _try_load(text[start:pos + 1])
    return _try_load(_fix_truncated(text[start:]))


def parse_document_response(raw):
    """Validate and bound the editable document structure returned by a model."""
    data = _extract_json_object(raw)
    if not isinstance(data, dict):
        raise ValueError("The model did not return a valid document structure")

    raw_pages = data.get("pages")
    if not isinstance(raw_pages, list) or not raw_pages:
        raw_blocks = data.get("blocks")
        if isinstance(raw_blocks, list):
            raw_pages = [{"page_number": 1, "blocks": raw_blocks}]
        else:
            raise ValueError("The document response did not contain any pages")

    pages = []
    allowed = {"heading", "paragraph", "bullet_list", "numbered_list",
               "table", "note"}
    for page_index, raw_page in enumerate(raw_pages[:500], 1):
        if not isinstance(raw_page, dict):
            continue
        blocks = []
        for raw_block in (raw_page.get("blocks") or [])[:1000]:
            if not isinstance(raw_block, dict):
                continue
            block_type = str(raw_block.get("type") or "paragraph").lower()
            if block_type not in allowed:
                block_type = "paragraph"
            block = {"type": block_type}
            if block_type in ("bullet_list", "numbered_list"):
                block["items"] = [str(item)[:12000] for item in
                                  (raw_block.get("items") or [])[:500]]
            elif block_type == "table":
                block["headers"] = [str(cell)[:5000] for cell in
                                    (raw_block.get("headers") or [])[:30]]
                block["rows"] = [
                    [str(cell)[:5000] for cell in row[:30]]
                    for row in (raw_block.get("rows") or [])[:1000]
                    if isinstance(row, list)
                ]
            else:
                block["text"] = str(raw_block.get("text") or "")[:30000]
                alignment = str(raw_block.get("alignment") or "left").lower()
                block["alignment"] = alignment if alignment in {
                    "left", "center", "right", "justify"} else "left"
                runs = raw_block.get("runs")
                if isinstance(runs, list):
                    block["runs"] = [
                        {
                            "text": str(run.get("text") or "")[:12000],
                            "bold": bool(run.get("bold")),
                            "italic": bool(run.get("italic")),
                            "underline": bool(run.get("underline")),
                        }
                        for run in runs[:500] if isinstance(run, dict)
                    ]
                if block_type == "heading":
                    try:
                        block["level"] = min(3, max(1, int(raw_block.get("level", 1))))
                    except (TypeError, ValueError):
                        block["level"] = 1
            blocks.append(block)
        pages.append({"page_number": page_index, "blocks": blocks})

    if not any(page["blocks"] for page in pages):
        raise ValueError("The document response contained no editable content")

    style = data.get("style") if isinstance(data.get("style"), dict) else {}
    return {
        "title": str(data.get("title") or "")[:1000],
        "subtitle": str(data.get("subtitle") or "")[:2000],
        "style": style,
        "pages": pages,
    }


def document_plain_text(document):
    parts = []
    for page in document.get("pages", []):
        for block in page.get("blocks", []):
            if "text" in block and block["text"]:
                parts.append(block["text"])
            parts.extend(block.get("items") or [])
            if block.get("headers"):
                parts.append(" | ".join(block["headers"]))
            for row in block.get("rows") or []:
                parts.append(" | ".join(row))
    return "\n".join(parts)


# ── API callers ───────────────────────────────────────────────────────────────

def call_openai_compatible(img_path, base_url, model, prompt, headers=None):
    b64, mime = encode_image(img_path)
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            {"type": "text", "text": prompt},
        ]}],
        "max_tokens": 2400,
        "temperature": 0,
    }
    r = requests.post(f"{base_url.rstrip('/')}/chat/completions",
                      json=payload, headers=headers or {}, timeout=180)
    r.raise_for_status()
    return r.json()["choices"][0]["message"]["content"], 0, 0


def get_first_model(base_url, fallback):
    try:
        r = requests.get(f"{base_url.rstrip('/')}/models", timeout=5)
        r.raise_for_status()
        models = r.json().get("data") or []
        if models:
            model_id = models[0].get("id")
            if model_id:
                return model_id
    except Exception:
        pass
    return fallback


def ensure_local_server(base_url, local_model="qwen"):
    """Start the configured llama.cpp server only for its default local endpoint."""
    base = base_url.rstrip("/")
    default_url, start_script, start_args = LOCAL_MODELS[local_model]
    try:
        response = requests.get(f"{base}/models", timeout=3)
        response.raise_for_status()
        return
    except requests.RequestException as exc:
        if base != default_url.rstrip("/") or not start_script.is_file():
            raise RuntimeError(f"Local vision server is unavailable at {base}. Start it or check --base-url.") from exc
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
             "-File", str(start_script), *start_args],
            capture_output=True, text=True, timeout=150,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
        if result.returncode:
            raise RuntimeError((result.stderr or result.stdout).strip()[-1000:])
        response = requests.get(f"{base}/models", timeout=5)
        response.raise_for_status()
    except (OSError, subprocess.SubprocessError, requests.RequestException) as exc:
        raise RuntimeError(f"Could not start llama.cpp at {base}: {exc}") from exc


def call_lmstudio(img_path, base_url, prompt):
    model = get_first_model(base_url, LMSTUDIO_MODEL)
    return call_openai_compatible(img_path, base_url, model, prompt)


def call_openrouter(img_path, model, api_key, prompt):
    b64, mime = encode_image(img_path)
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{b64}"}},
            {"type": "text", "text": prompt},
        ]}],
        "max_tokens": 2400,
        "temperature": 0,
    }
    headers = {
        "Authorization": f"Bearer {api_key}",
        "HTTP-Referer": "https://image-analyzer-tool",
        "X-Title": "Image Analyzer",
    }
    r = requests.post(f"{OPENROUTER_BASE}/chat/completions",
                      json=payload, headers=headers, timeout=180)
    r.raise_for_status()
    resp  = r.json()
    usage = resp.get("usage", {})
    return (resp["choices"][0]["message"]["content"],
            usage.get("prompt_tokens", 0),
            usage.get("completion_tokens", 0))


def call_openrouter_pdf(pdf_path, model, api_key, engine=None):
    encoded = encode_pdf(pdf_path)
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": [
            {"type": "file", "file": {
                "filename": os.path.basename(pdf_path),
                "file_data": f"data:application/pdf;base64,{encoded}",
            }},
            {"type": "text", "text": PDF_DOCUMENT_PROMPT},
        ]}],
        "max_tokens": 16000,
        "temperature": 0,
    }
    if engine:
        payload["plugins"] = [{"id": "file-parser", "pdf": {"engine": engine}}]
    headers = {
        "Authorization": f"Bearer {api_key}",
        "HTTP-Referer": "https://image-analyzer-tool",
        "X-Title": "Image Analyzer",
    }
    response = requests.post(
        f"{OPENROUTER_BASE}/chat/completions",
        json=payload,
        headers=headers,
        timeout=600,
    )
    response.raise_for_status()
    data = response.json()
    message = data["choices"][0]["message"]
    usage = data.get("usage", {})
    return (message.get("content") or "",
            usage.get("prompt_tokens", 0),
            usage.get("completion_tokens", 0),
            message.get("annotations") or [])


# ── Cost estimation ───────────────────────────────────────────────────────────

def estimate_tokens(img_path):
    try:
        if Path(img_path).suffix.lower() == ".pdf":
            page_count = get_pdf_info(img_path)["pages"] or 1
            return 900 * page_count, min(1400 * page_count, 16000)
        sz = os.path.getsize(img_path)
        return 800 + min(int(sz / 1200), 1500), 350
    except Exception:
        return 1150, 350


def parse_price(price_str):
    try:
        nums = re.findall(r"\$?([\d.]+)", price_str)
        return float(nums[0]), float(nums[1])
    except Exception:
        return 0.0, 0.0


def openrouter_label(model):
    return f"{model['label']}  ({model['price']})"


def openrouter_model_from_label(label):
    for model in OPENROUTER_MODELS:
        if label == openrouter_label(model) or label == model["label"]:
            return model
    return OPENROUTER_MODELS[0]


# =============================================================================
#  GUI Application
# =============================================================================

class ImageAnalyzerApp(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()
        self.title("Image Analyzer AI")
        self.geometry("1320x816")
        self.minsize(980, 700)
        self.configure(bg=BG)
        self._setup_styles()

        self.input_paths = []   # queued image and PDF paths
        self.results     = []   # processed result dicts (never auto-cleared)
        self.running     = False
        self.stop_flag   = False
        self.geocode_cache = {}
        self.openrouter_key = OPENROUTER_KEY
        self.geocoding_key = GEOCODING_KEY

        self._build_ui()

    # ── Styles ────────────────────────────────────────────────────────────────

    def _setup_styles(self):
        s = ttk.Style(self)
        s.theme_use("clam")
        s.configure(".", background=BG, foreground=TEXT, font=("Segoe UI", 10))
        s.configure("TFrame",        background=BG)
        s.configure("TLabel",        background=BG, foreground=TEXT)
        s.configure("TCombobox",     background=BG3, foreground=TEXT,
                    fieldbackground=BG3, selectbackground=ACCENT,
                    arrowcolor=ACCENT2)
        s.map("TCombobox", fieldbackground=[("readonly", BG3)])
        s.configure("Accent.TButton", background=ACCENT, foreground="#071b1a",
                    font=("Segoe UI Semibold", 10), padding=(16, 9),
                    relief="flat", borderwidth=0)
        s.map("Accent.TButton",
              background=[("active", ACCENT2), ("pressed", "#20b8a5")],
              foreground=[("active", "#071b1a"), ("pressed", "#071b1a")])
        s.configure("Ghost.TButton", background=BG3, foreground=TEXT,
                    font=("Segoe UI Semibold", 9), padding=(14, 8),
                    relief="flat", borderwidth=0)
        s.map("Ghost.TButton",
              background=[("active", BORDER), ("pressed", BG2)])
        s.configure("Compact.TButton", background=BG3, foreground=TEXT,
                    font=("Segoe UI", 9), padding=(8, 3),
                    relief="flat", borderwidth=0)
        s.map("Compact.TButton",
              background=[("active", BORDER), ("pressed", BG2)])
        s.configure("Danger.TButton", background=ERR_BG, foreground=ERR,
                    font=("Segoe UI Semibold", 9), padding=(14, 8),
                    relief="flat", borderwidth=0)
        s.map("Danger.TButton", background=[("active", ERR_HOVER)])
        s.configure("TProgressbar", troughcolor=BG3, background=ACCENT,
                    lightcolor=ACCENT2, darkcolor=ACCENT,
                    borderwidth=0, thickness=7)
        for orientation in ("Vertical", "Horizontal"):
            style_name = f"Theme.{orientation}.TScrollbar"
            s.configure(style_name, background=BG3, troughcolor=BG2,
                        bordercolor=BORDER, arrowcolor=ACCENT2,
                        lightcolor=BG3, darkcolor=BG3, relief="flat",
                        borderwidth=0)
            s.map(style_name,
                  background=[("active", BORDER), ("pressed", ACCENT2)],
                  arrowcolor=[("active", ACCENT)])

    # ── UI ────────────────────────────────────────────────────────────────────

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG2, height=56)
        hdr.pack(fill="x", side="top")
        hdr.pack_propagate(False)
        tk.Label(hdr, text="⬡  Image Analyzer AI", bg=BG2, fg=TEXT,
                 font=("Segoe UI Semibold", 15)).pack(side="left", padx=22, pady=12)
        tk.Label(hdr, text="llama.cpp  ·  OpenRouter  ·  Images and PDFs",
                 bg=BG2, fg=MUTED, font=("Segoe UI", 9)).pack(side="left")

        body = tk.Frame(self, bg=BG)
        body.pack(fill="both", expand=True, padx=18, pady=12)

        left_shell = tk.Frame(body, bg=BG, width=326)
        left_shell.pack(side="left", fill="y", padx=(0, 14))
        left_shell.pack_propagate(False)
        self.left_canvas = tk.Canvas(
            left_shell, bg=BG, highlightthickness=0, bd=0, width=306)
        left_scroll = ttk.Scrollbar(
            left_shell, orient="vertical", command=self.left_canvas.yview,
            style="Theme.Vertical.TScrollbar")
        self.left_canvas.configure(yscrollcommand=left_scroll.set)
        left_scroll.pack(side="right", fill="y")
        self.left_canvas.pack(side="left", fill="both", expand=True)
        left = tk.Frame(self.left_canvas, bg=BG, width=306)
        self._left_window = self.left_canvas.create_window(
            (0, 0), window=left, anchor="nw")
        left.bind(
            "<Configure>",
            lambda _event: self.left_canvas.configure(
                scrollregion=self.left_canvas.bbox("all")))
        self.left_canvas.bind(
            "<Configure>",
            lambda event: self.left_canvas.itemconfigure(
                self._left_window, width=event.width))
        self._build_left(left)

        right = tk.Frame(body, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_right(right)

    def _card(self, parent, title=None, side="top"):
        outer = tk.Frame(parent, bg=BORDER)
        outer.pack(fill="x", pady=(0, 0 if side == "bottom" else 10), side=side)
        inner = tk.Frame(outer, bg=BG2)
        inner.pack(fill="both", expand=True, padx=1, pady=1)
        if title:
            tk.Label(inner, text=title.upper(), bg=BG2, fg=MUTED,
                     font=("Segoe UI", 8, "bold")).pack(anchor="w",
                     padx=14, pady=(10, 4))
        return inner

    def _dark_checkbutton(self, parent, text, variable):
        return tk.Checkbutton(parent, text=text, variable=variable,
                              bg=BG2, fg=TEXT, selectcolor=BG3,
                              activebackground=BG2, activeforeground=TEXT,
                              cursor="hand2", font=("Segoe UI", 9),
                              indicatoron=True, highlightthickness=0)

    def _build_left(self, parent):
        # Provider
        c = self._card(parent, "AI Provider")
        self.provider_var = tk.StringVar(value="local")
        row = tk.Frame(c, bg=BG2)
        row.pack(padx=14, pady=(0, 8), fill="x")
        for label, val in [("OpenRouter", "openrouter"),
                           ("llama.cpp", "local")]:
            tk.Radiobutton(
                row, text=label, variable=self.provider_var, value=val,
                indicatoron=False, bg=BG3, fg=MUTED, selectcolor=ACCENT2,
                activebackground=BORDER, activeforeground=TEXT,
                cursor="hand2", font=("Segoe UI Semibold", 9), relief="flat",
                bd=0, padx=10, pady=6, width=12,
            ).pack(side="left", padx=(0, 5), fill="x", expand=True)
        self.provider_var.trace_add("write", self._on_provider_change)

        # Local pane
        self.local_frame = tk.Frame(c, bg=BG2)
        tk.Label(self.local_frame, text="Base URL:", bg=BG2, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(self.local_frame, text="Local model:", bg=BG2, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")
        self.local_model_var = tk.StringVar(value="qwen")
        self.local_model_labels = {"Qwen3.6 35B A3B": "qwen", "Gemma 4 E4B": "gemma"}
        self.local_model_display = tk.StringVar(value="Qwen3.6 35B A3B")
        local_combo = ttk.Combobox(self.local_frame, textvariable=self.local_model_display,
                                   state="readonly", values=tuple(self.local_model_labels))
        local_combo.pack(fill="x", pady=(0, 5))
        def select_local_model(_event):
            self.local_model_var.set(self.local_model_labels[self.local_model_display.get()])
            self.lmstudio_base_var.set(LOCAL_MODELS[self.local_model_var.get()][0])
        local_combo.bind("<<ComboboxSelected>>", select_local_model)
        self.lmstudio_base_var = tk.StringVar(value=LMSTUDIO_BASE)
        tk.Entry(self.local_frame, textvariable=self.lmstudio_base_var, bg=BG3,
                 fg=TEXT, relief="flat", font=("Segoe UI", 9),
                 insertbackground=TEXT, bd=0
                 ).pack(fill="x", ipady=5)

        # OpenRouter pane (hidden initially)
        self.or_frame = tk.Frame(c, bg=BG2)
        tk.Label(self.or_frame, text="Vision Model:", bg=BG2, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w")
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(self.or_frame, textvariable=self.model_var,
                                        state="readonly", font=("Segoe UI", 9))
        self.model_combo["values"] = [openrouter_label(m) for m in OPENROUTER_MODELS]
        self.model_combo.current(0)
        self.model_combo.pack(fill="x")
        self.model_combo.bind("<<ComboboxSelected>>", lambda _: self._update_cost())
        tk.Label(self.or_frame, text="PDF Parser:", bg=BG2, fg=MUTED,
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(7, 0))
        self.pdf_engine_var = tk.StringVar(value=next(iter(PDF_ENGINES)))
        self.pdf_engine_combo = ttk.Combobox(
            self.or_frame, textvariable=self.pdf_engine_var, state="readonly",
            values=list(PDF_ENGINES), font=("Segoe UI", 9),
        )
        self.pdf_engine_combo.pack(fill="x")
        self.local_frame.pack(fill="x", padx=14, pady=(0, 8))

        # Test connection
        tc = tk.Frame(c, bg=BG2)
        tc.pack(fill="x", padx=14, pady=(0, 4))
        ttk.Button(tc, text="Test Connection", style="Compact.TButton",
                   command=self._test_connection).pack(fill="x")
        self.conn_label = tk.Label(c, text="", bg=BG2, fg=MUTED,
                                   font=("Segoe UI", 9))

        self.translate_var = tk.BooleanVar(value=True)
        self.geocode_var = tk.BooleanVar(value=True)

        # Drop zone
        c2 = self._card(parent, "Image / PDF Input")
        self.drop_zone = tk.Frame(c2, bg=BG3, height=110, cursor="hand2")
        self.drop_zone.pack(fill="x", padx=14, pady=(0, 6))
        self.drop_zone.pack_propagate(False)
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_zone.bind("<Button-1>", lambda _e: self._browse())
        tk.Label(self.drop_zone,
                 text="⬇\nDrop images, PDFs, or a folder\nor click to browse",
                 bg=BG3, fg=MUTED, font=("Segoe UI", 10),
                 justify="center").place(relx=0.5, rely=0.5, anchor="center")

        b2 = tk.Frame(c2, bg=BG2)
        b2.pack(fill="x", padx=14, pady=(0, 6))
        ttk.Button(b2, text="Browse Folder", style="Ghost.TButton",
                   command=self._browse_folder).pack(side="left", fill="x",
                   expand=True, padx=(0, 4))
        ttk.Button(b2, text="Clear Queue", style="Danger.TButton",
                   command=self._clear_queue).pack(side="left")

        self.queue_label = tk.Label(c2, text="No files queued",
                                    bg=BG2, fg=MUTED, font=("Segoe UI", 9))
        self.queue_label.pack(anchor="w", padx=14, pady=(0, 4))
        self._dark_checkbutton(
            c2, "Translate to English", self.translate_var
        ).pack(anchor="w", padx=14, pady=(0, 2))
        self._dark_checkbutton(
            c2, "Lookup GPS location", self.geocode_var
        ).pack(anchor="w", padx=14, pady=(0, 4))
        tk.Label(c2, text="API keys: OPENROUTER_KEY (or OPENROUTER_API_KEY) and GEOCODING_KEY",
                 bg=BG2, fg=MUTED, font=("Segoe UI", 8),
                 wraplength=260, justify="left").pack(anchor="w", padx=14,
                                                       pady=(0, 8))

        # Cost estimate
        c3 = self._card(parent, "Cost Estimate")
        self.cost_label = tk.Label(c3, text="—", bg=BG2, fg=TEXT,
                                   font=("Segoe UI Semibold", 14))
        self.cost_label.pack(anchor="w", padx=14, pady=(0, 2))
        self.cost_detail = tk.Label(c3, text="Add images to estimate",
                                    bg=BG2, fg=MUTED, font=("Segoe UI", 9),
                                    wraplength=260, justify="left")
        self.cost_detail.pack(anchor="w", padx=14, pady=(0, 10))

        # Run / Stop
        c4 = self._card(parent, side="bottom")
        self.run_btn = ttk.Button(c4, text="▶  Analyze Files",
                                  style="Accent.TButton", command=self._run)
        self.run_btn.pack(fill="x", padx=14, pady=(8, 4))
        self.stop_btn = ttk.Button(c4, text="■  Stop", style="Ghost.TButton",
                                   command=self._stop, state="disabled")
        self.stop_btn.pack(fill="x", padx=14, pady=(0, 7))

    def _build_right(self, parent):
        # Toolbar is packed first so its lower edge always aligns with the
        # bottom of the left-side Analyze / Stop card.
        toolbar = tk.Frame(parent, bg=BG2, height=48)
        toolbar.pack(fill="x", side="bottom")
        toolbar.pack_propagate(False)
        ttk.Button(toolbar, text="Copy", style="Ghost.TButton",
                   command=self._copy_table).pack(side="left", padx=(10, 4), pady=8)
        ttk.Button(toolbar, text="Excel", style="Accent.TButton",
                   command=self._download_excel).pack(side="left", pady=8)
        ttk.Button(toolbar, text="CSV", style="Ghost.TButton",
                   command=self._download_csv).pack(side="left", padx=4, pady=8)
        ttk.Button(toolbar, text="Clear", style="Danger.TButton",
                   command=self._clear_results).pack(side="left", padx=(16, 0), pady=8)
        self.status_bar = tk.Label(toolbar, text="", bg=BG2, fg=MUTED,
                                   font=("Segoe UI", 9))
        self.status_bar.pack(side="right", padx=14)

        # Progress bar sits directly above the shared bottom toolbar.
        prog_wrap = tk.Frame(parent, bg=BG)
        prog_wrap.pack(fill="x", side="bottom", pady=(6, 0))
        ph = tk.Frame(prog_wrap, bg=BG)
        ph.pack(fill="x")
        self.prog_label = tk.Label(ph, text="Ready", bg=BG, fg=MUTED,
                                   font=("Segoe UI", 9))
        self.prog_label.pack(side="left")
        self.prog_pct = tk.Label(ph, text="", bg=BG, fg=ACCENT2,
                                 font=("Segoe UI Semibold", 9))
        self.prog_pct.pack(side="right")
        self.progressbar = ttk.Progressbar(prog_wrap, mode="determinate",
                                           style="TProgressbar")
        self.progressbar.pack(fill="x", pady=(3, 0))

        # One synchronized horizontal viewport for both the header and rows.
        grid = tk.Frame(parent, bg=BG)
        grid.pack(fill="both", expand=True)
        self.header_canvas = tk.Canvas(grid, bg=BG3, height=32,
                                       highlightthickness=0, bd=0)
        self.header_canvas.pack(fill="x", side="top")
        header_frame = tk.Frame(self.header_canvas, bg=BG3,
                                width=RESULTS_WIDTH, height=32)
        header_frame.pack_propagate(False)
        x_pos = 0
        for _key, label, width, _scrollable in RESULT_COLUMNS:
            tk.Label(header_frame, text=label, bg=BG3, fg=MUTED,
                     font=("Segoe UI", 8, "bold"), anchor="w",
                     padx=8).place(x=x_pos, y=0, width=width, height=32)
            x_pos += width
        self._header_window = self.header_canvas.create_window(
            (0, 0), window=header_frame, anchor="nw", width=RESULTS_WIDTH,
            height=32,
        )
        self.header_canvas.configure(scrollregion=(0, 0, RESULTS_WIDTH, 32))

        cv_wrap = tk.Frame(grid, bg=BG)
        cv_wrap.pack(fill="both", expand=True, pady=(2, 0))
        self.canvas = tk.Canvas(cv_wrap, bg=BG, highlightthickness=0, bd=0)
        vertical = ttk.Scrollbar(
            cv_wrap, orient="vertical", command=self.canvas.yview,
            style="Theme.Vertical.TScrollbar",
        )
        self.horizontal = ttk.Scrollbar(
            cv_wrap, orient="horizontal", command=self._scroll_x,
            style="Theme.Horizontal.TScrollbar",
        )
        self.canvas.configure(yscrollcommand=vertical.set,
                              xscrollcommand=self._set_xscroll)
        vertical.pack(side="right", fill="y")
        self.horizontal.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.results_frame = tk.Frame(self.canvas, bg=BG, width=RESULTS_WIDTH)
        self._cwin = self.canvas.create_window(
            (0, 0), window=self.results_frame, anchor="nw", width=RESULTS_WIDTH,
        )
        self.results_frame.bind("<Configure>", self._on_frame_cfg)
        self.canvas.bind("<Configure>", self._on_canvas_cfg)
        self.canvas.bind_all("<MouseWheel>", self._on_wheel)
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_shift_wheel)

    # ── Canvas helpers ────────────────────────────────────────────────────────

    def _on_frame_cfg(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_cfg(self, e):
        content_width = max(RESULTS_WIDTH, e.width)
        self.canvas.itemconfig(self._cwin, width=content_width)
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_wheel(self, e):
        self.canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")

    def _on_shift_wheel(self, e):
        self._scroll_x("scroll", int(-1 * (e.delta / 120)), "units")
        return "break"

    def _set_xscroll(self, first, last):
        self.horizontal.set(first, last)
        self.header_canvas.xview_moveto(float(first))

    def _scroll_x(self, *args):
        self.canvas.xview(*args)
        self.header_canvas.xview(*args)

    def _on_cell_wheel(self, event, widget):
        widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
        return "break"

    def _show_row_details(self, row):
        window = tk.Toplevel(self)
        window.title(f"Result details - {row.get('Filename', '')}")
        window.geometry("900x600")
        window.configure(bg=BG2)
        text_wrap = tk.Frame(window, bg=BG2)
        text_wrap.pack(fill="both", expand=True, padx=12, pady=12)
        scrollbar = ttk.Scrollbar(text_wrap, orient="vertical",
                                  style="Theme.Vertical.TScrollbar")
        detail = tk.Text(text_wrap, bg=BG3, fg=TEXT, insertbackground=TEXT,
                         selectbackground=ACCENT, relief="flat", bd=0,
                         wrap="word", font=("Segoe UI", 10), padx=12, pady=10,
                         yscrollcommand=scrollbar.set)
        scrollbar.configure(command=detail.yview)
        scrollbar.pack(side="right", fill="y")
        detail.pack(side="left", fill="both", expand=True)
        for key, label, _width, _scrollable in RESULT_COLUMNS:
            detail.insert("end", f"{label}\n", "heading")
            detail.insert("end", f"{row.get(key, '') or '—'}\n\n")
        detail.tag_configure("heading", foreground=ACCENT2,
                             font=("Segoe UI Semibold", 10))
        detail.configure(state="disabled")

    # ── Provider toggle ───────────────────────────────────────────────────────

    def _on_provider_change(self, *_):
        provider = self.provider_var.get()
        self.local_frame.pack_forget()
        self.or_frame.pack_forget()
        if provider == "local":
            self.local_frame.pack(fill="x", padx=14, pady=(0, 8))
        else:
            self.or_frame.pack(fill="x", padx=14, pady=(0, 8))
        self._update_cost()

    # ── Image loading ─────────────────────────────────────────────────────────

    def _on_drop(self, e):
        paths = []
        for token in re.findall(r'\{[^}]+\}|\S+', e.data):
            p = token.strip("{}")
            if os.path.isdir(p):
                paths.extend(self._collect_dir(p))
            elif os.path.isfile(p) and Path(p).suffix.lower() in SUPPORTED_EXT:
                paths.append(p)
        self._add_paths(paths)

    def _browse(self):
        files = filedialog.askopenfilenames(
            title="Select Images or PDFs",
            filetypes=[("Supported files", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.tiff *.tif *.pdf"),
                       ("PDF documents", "*.pdf"),
                       ("Images", "*.jpg *.jpeg *.png *.gif *.bmp *.webp *.tiff *.tif"),
                       ("All files", "*.*")])
        if files:
            self._add_paths(list(files))

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder")
        if folder:
            self._add_paths(self._collect_dir(folder))

    def _collect_dir(self, folder):
        out = []
        for root, _, files in os.walk(folder):
            for f in files:
                if Path(f).suffix.lower() in SUPPORTED_EXT:
                    out.append(os.path.join(root, f))
        return out

    def _add_paths(self, paths):
        existing = set(self.input_paths)
        new = [p for p in paths if p not in existing]
        self.input_paths.extend(new)
        n = len(self.input_paths)
        self.queue_label.config(
            text=f"{n} file{'s' if n != 1 else ''} queued  (+{len(new)} new)")
        self._update_cost()

    def _clear_queue(self):
        """Clear input queue only — results table is untouched."""
        self.input_paths = []
        self.queue_label.config(text="No files queued")
        self._update_cost()

    def _clear_results(self):
        """Clear results table only — image queue is untouched."""
        self.results = []
        for w in self.results_frame.winfo_children():
            w.destroy()
        self.progressbar["value"] = 0
        self.prog_label.config(text="Ready")
        self.prog_pct.config(text="")
        self.status_bar.config(text="Results cleared")

    # ── Cost estimate ─────────────────────────────────────────────────────────

    def _update_cost(self):
        n = len(self.input_paths)
        if n == 0:
            self.cost_label.config(text="—")
            self.cost_detail.config(text="Add images to estimate")
            return
        if self.provider_var.get() == "local":
            self.cost_label.config(text="$0.00  (local)", fg=SUCCESS)
            self.cost_detail.config(text="Running on llama.cpp - no API cost. PDFs require OpenRouter.")
            return
        model = openrouter_model_from_label(self.model_var.get())
        model_id = model["id"]
        in_r, out_r = parse_price(model["price"])
        total_in = total_out = 0
        for p in self.input_paths:
            ti, to = estimate_tokens(p)
            total_in  += ti
            total_out += to
        cost = (total_in * in_r + total_out * out_r) / 1_000_000
        self.cost_label.config(text=f"≈ ${cost:.4f}", fg=WARN)
        self.cost_detail.config(
            text=(f"{n} file{'s' if n != 1 else ''}  ·  "
                  f"~{total_in // 1000}K in / {total_out // 1000}K out tokens\n"
                  f"{model['label']}"
                  f"{'  ·  parser fees may apply' if any(Path(p).suffix.lower() == '.pdf' for p in self.input_paths) else ''}"))

    # ── Connection test ───────────────────────────────────────────────────────

    def _set_conn_status(self, text, fg):
        if not self.conn_label.winfo_ismapped():
            self.conn_label.pack(anchor="w", padx=14, pady=(0, 4))
        self.conn_label.config(text=text, fg=fg)

    def _test_connection(self):
        self._set_conn_status("Testing...", MUTED)
        threading.Thread(target=self._do_test, daemon=True).start()

    def _do_test(self):
        try:
            provider = self.provider_var.get()
            if provider == "local":
                base_url = self.lmstudio_base_var.get().strip().rstrip("/") or LMSTUDIO_BASE
                r = requests.get(f"{base_url}/models", timeout=5)
                r.raise_for_status()
                self.after(0, lambda: self._set_conn_status(
                    "OK: local server connected", SUCCESS))
            else:
                api_key = self.openrouter_key
                if not api_key:
                    raise RuntimeError("OPENROUTER_KEY is not set")
                headers = {
                    "Authorization": f"Bearer {api_key}",
                    "HTTP-Referer": "https://image-analyzer-tool",
                    "X-Title": "Image Analyzer",
                }
                r = requests.get(f"{OPENROUTER_BASE}/models",
                                 headers=headers, timeout=8)
                r.raise_for_status()
                self.after(0, lambda: self._set_conn_status(
                    "OK: OpenRouter connected", SUCCESS))
        except Exception as ex:
            msg = str(ex)[:58]
            self.after(0, lambda: self._set_conn_status(
                f"Error: {msg}", ERR))

    # ── Analysis ──────────────────────────────────────────────────────────────

    def _run(self):
        if not self.input_paths:
            messagebox.showwarning("No Files", "Please add images or PDFs first.")
            return
        provider = self.provider_var.get()
        has_pdf = any(Path(path).suffix.lower() == ".pdf" for path in self.input_paths)
        has_images = any(Path(path).suffix.lower() in IMAGE_EXT for path in self.input_paths)
        if has_pdf and provider != "openrouter":
            messagebox.showwarning(
                "OpenRouter Required",
                "PDF reconstruction uses OpenRouter. Select OpenRouter and try again.",
            )
            return
        model = openrouter_model_from_label(self.model_var.get())
        if provider == "openrouter" and not self.openrouter_key:
            messagebox.showwarning(
                "Missing API Key",
                "Set OPENROUTER_KEY or OPENROUTER_API_KEY before using OpenRouter.",
            )
            return
        if provider == "openrouter" and has_images and not model.get("image"):
            messagebox.showwarning(
                "Model Does Not Accept Images",
                f"{model['label']} can process parsed PDF text but not image files. "
                "Choose a vision model for this queue.",
            )
            return
        self.run_settings = {
            "provider": provider,
            "local_model": self.local_model_var.get(),
            "lmstudio_base": self.lmstudio_base_var.get().strip() or LMSTUDIO_BASE,
            "model_id": model["id"],
            "pdf_engine": PDF_ENGINES.get(self.pdf_engine_var.get()),
            "api_key": self.openrouter_key,
            "translate": self.translate_var.get(),
            "geocoding_key": self.geocoding_key if self.geocode_var.get() else "",
        }
        self.running   = True
        self.stop_flag = False
        self.run_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self.progressbar["maximum"] = len(self.input_paths)
        self.progressbar["value"]   = 0
        threading.Thread(target=self._worker, daemon=True).start()

    def _stop(self):
        self.stop_flag = True
        self.stop_btn.config(state="disabled")
        self.prog_label.config(text="Stopping…", fg=WARN)

    def _process_pdf(self, path, settings):
        pdf_info = get_pdf_info(path)
        raw, in_tok, out_tok, annotations = call_openrouter_pdf(
            path, settings["model_id"], settings["api_key"], settings["pdf_engine"])
        document = parse_document_response(raw)
        plain_text = document_plain_text(document)
        title = document.get("title") or Path(path).stem
        page_count = pdf_info["pages"] or len(document.get("pages") or [])
        metadata = {"pdf": {"pages": page_count,
                            "properties": pdf_info.get("metadata") or {}}}
        return {
            "Filename": os.path.basename(path),
            "Path": path,
            "Dimensions": f"{page_count} page{'s' if page_count != 1 else ''}",
            "Color Mode": "PDF",
            "File Size": human_size(os.path.getsize(path)),
            "Description": f"PDF text and structure analysis ready: {title}",
            "Text Content": plain_text[:2400],
            "Translation": "",
            "Photographer": pdf_info.get("metadata", {}).get("Author", ""),
            "Location": "",
            "GPS Coordinates": "",
            "Date Taken": pdf_info.get("metadata", {}).get("CreationDate", ""),
            "Subject Tags": "PDF, extracted text, document structure",
            "Quality Notes": "Text and reading order extracted; compare complex layouts with the source PDF.",
            "Metadata": _clean_metadata_json(metadata),
            "In Tokens": in_tok,
            "Out Tokens": out_tok,
            "_source_type": "pdf",
            "_document_structure": document,
            "_file_annotations": annotations,
        }

    def _process_image(self, path, settings, prompt):
        with Image.open(path) as img:
            dims = f"{img.width} × {img.height} px"
            mode = img.mode
            exif = get_exif_data(path, img)
        size = human_size(os.path.getsize(path))

        def analyze_once():
            if settings["provider"] == "local":
                return call_lmstudio(path, settings["lmstudio_base"], prompt)
            return call_openrouter(
                path, settings["model_id"], settings["api_key"], prompt)

        raw, in_tok, out_tok = analyze_once()
        ai = parse_ai_response(raw)
        if _weak_ai_result(ai) and not self.stop_flag:
            retry_raw, retry_in, retry_out = analyze_once()
            retry_ai = parse_ai_response(retry_raw)
            if not _weak_ai_result(retry_ai):
                ai = retry_ai
            in_tok += retry_in
            out_tok += retry_out

        gps_coordinates = exif["gps"] or ai.get("gps_coordinates", "")
        address = reverse_geocode(
            gps_coordinates, settings["geocoding_key"], self.geocode_cache)
        location = " | ".join(dict.fromkeys(value for value in (ai["location"], address) if value))
        return {
            "Filename": os.path.basename(path),
            "Path": path,
            "Dimensions": dims,
            "Color Mode": mode,
            "File Size": size,
            "Description": ai["description"],
            "Text Content": ai["text_content"],
            "Translation": ai["translation"],
            "Photographer": exif["photographer"] or ai["photographer"],
            "Location": location,
            "GPS Coordinates": gps_coordinates,
            "Date Taken": exif["date_taken"] or ai["date_taken"],
            "Subject Tags": ai["subject_tags"],
            "Quality Notes": ai["quality_notes"],
            "Metadata": build_metadata_column(exif, ai),
            "In Tokens": in_tok,
            "Out Tokens": out_tok,
            "_source_type": "image",
        }

    def _worker(self):
        queued = list(self.input_paths)
        settings = dict(self.run_settings)
        total = len(queued)
        errors = 0
        processed = 0
        prompt = build_analysis_prompt(settings["translate"])
        if settings["provider"] == "local":
            try:
                ensure_local_server(settings["lmstudio_base"], settings["local_model"])
            except Exception as exc:
                self.after(0, lambda message=str(exc): (
                    self.prog_label.config(text="Local server failed", fg=ERR),
                    self.status_bar.config(text=message),
                    self.run_btn.config(state="normal"), self.stop_btn.config(state="disabled"),
                    messagebox.showerror("Local server failed", message)))
                self.running = False
                return
        for i, path in enumerate(queued):
            if self.stop_flag:
                break
            fname = os.path.basename(path)
            idx   = i + 1
            self.after(0, lambda f=fname, x=idx, t=total: (
                self.prog_label.config(text=f"[{x}/{t}]  {f}"),
                self.prog_pct.config(text=f"{int(x / t * 100)}%"),
                self.progressbar.configure(value=x),
            ))
            try:
                if Path(path).suffix.lower() == ".pdf":
                    row = self._process_pdf(path, settings)
                else:
                    row = self._process_image(path, settings, prompt)
                self.results.append(row)
                self.after(0, lambda r=row: self._add_result_row(r))

            except Exception as ex:
                errors += 1
                row = {
                    "Filename": fname, "Path": path,
                    "Dimensions": "", "Color Mode": "",
                    "File Size": human_size(os.path.getsize(path)),
                    "Description": f"ERROR: {str(ex)[:200]}",
                    "Text Content": "", "Translation": "",
                    "Photographer": "", "Location": "", "GPS Coordinates": "",
                    "Date Taken": "", "Subject Tags": "",
                    "Quality Notes": "", "Metadata": "",
                    "In Tokens": 0, "Out Tokens": 0,
                    "_source_type": "pdf" if Path(path).suffix.lower() == ".pdf" else "image",
                }
                self.results.append(row)
                self.after(0, lambda r=row: self._add_result_row(r, error=True))
            processed += 1

        ts   = datetime.datetime.now().strftime("%H:%M:%S")
        self.after(0, lambda: (
            self.prog_label.config(
                text=(f"{'Stopped' if self.stop_flag else 'Done'}  ·  {processed} processed, "
                      f"{errors} error{'s' if errors != 1 else ''}"),
                fg=SUCCESS if errors == 0 else WARN),
            self.prog_pct.config(text=f"{int(processed / total * 100) if total else 0}%"),
            self.run_btn.config(state="normal"),
            self.stop_btn.config(state="disabled"),
            self.status_bar.config(text=f"{len(self.results)} rows ready  ·  {ts}"),
        ))
        self.running = False

    # ── Results row ───────────────────────────────────────────────────────────

    def _add_result_row(self, row, error=False):
        idx = len(self.results_frame.winfo_children())
        bg  = ERR_BG if error else (BG if idx % 2 == 0 else BG2)

        frame = tk.Frame(self.results_frame, bg=bg, width=RESULTS_WIDTH,
                         height=RESULT_ROW_HEIGHT)
        frame.pack(fill="x", pady=(0, 1))
        frame.pack_propagate(False)
        x_pos = 0
        for key, _label, width, scrollable in RESULT_COLUMNS:
            value = str(row.get(key, "") or "—")
            if scrollable:
                cell = tk.Text(
                    frame, bg=bg, fg=ERR if error else TEXT,
                    selectbackground=ACCENT, selectforeground=BG,
                    relief="flat", bd=0, highlightthickness=0, wrap="word",
                    font=("Segoe UI", 9), padx=7, pady=4, cursor="xterm",
                    takefocus=False,
                )
                cell.insert("1.0", value)
                cell.configure(state="disabled")
                cell.bind("<MouseWheel>",
                          lambda event, widget=cell: self._on_cell_wheel(event, widget))
            else:
                cell = tk.Label(
                    frame, text=value, bg=bg, fg=ERR if error else TEXT,
                    font=("Segoe UI", 9), anchor="nw", justify="left",
                    padx=8, pady=6,
                )
            cell.place(x=x_pos, y=0, width=width - 1,
                       height=RESULT_ROW_HEIGHT - 1)
            if key == "Filename":
                cell.configure(cursor="hand2")
                cell.bind("<Button-1>",
                          lambda event, result=row: self._copy_row(event, result))
            cell.bind("<Double-Button-1>",
                      lambda _event, result=row: self._show_row_details(result))
            tk.Frame(frame, bg=BORDER).place(
                x=x_pos + width - 1, y=0, width=1, height=RESULT_ROW_HEIGHT)
            x_pos += width

    # ── Export ────────────────────────────────────────────────────────────────

    EXPORT_COLS = EXPORT_COLS

    def _copy_row(self, event, row):
        buf = io.StringIO()
        writer = csv.writer(buf, delimiter="\t", lineterminator="\r\n",
                            quoting=csv.QUOTE_MINIMAL)
        writer.writerow([_export_value(row.get(c, "")) for c in self.EXPORT_COLS])
        try:
            self.clipboard_clear()
            self.clipboard_append(buf.getvalue())
            self.update_idletasks()
        except Exception:
            messagebox.showerror("Copy failed", "Could not copy this row.")
            return
        tip = tk.Toplevel(self)
        tip.overrideredirect(True)
        tip.configure(bg=BG3)
        tk.Label(tip, text="Copied to clipboard", bg=BG3, fg=TEXT,
                 font=("Segoe UI", 9), padx=8, pady=4).pack()
        tip.geometry(f"+{event.x_root + 12}+{event.y_root + 12}")
        tip.after(1300, tip.destroy)

    def _copy_table(self):
        if not self.results:
            messagebox.showinfo("Nothing to copy", "Run analysis first.")
            return
        buf = io.StringIO()
        writer = csv.writer(buf, delimiter="\t", lineterminator="\r\n",
                            quoting=csv.QUOTE_MINIMAL)
        writer.writerow(self.EXPORT_COLS)
        for r in self.results:
            writer.writerow([_export_value(r.get(c, "")) for c in self.EXPORT_COLS])
        text = buf.getvalue()
        # The Windows clipboard and Excel both choke on enormous blobs. If the
        # whole table is too big to share/paste reliably, grow the per-cell
        # limit back (Excel's hard ceiling is 32767) and warn the user instead
        # of silently failing.
        try:
            self.clipboard_clear()
            self.clipboard_append(text)
        except Exception:
            messagebox.showerror(
                "Copy failed",
                "The full table is too large for the Windows clipboard.\n"
                "Use the CSV or Excel button to save it to a file instead.")
            return
        self.update_idletasks()
        self.after(100, lambda: self._verify_clipboard(text))
        self.status_bar.config(text="Excel-compatible table copied to clipboard ✓")

    def _verify_clipboard(self, expected):
        """Tk's clipboard can silently drop very large payloads; detect that
        and, if so, retry the copy."""
        try:
            got = self.clipboard_get()
        except Exception:
            got = ""
        if not got:
            # Nothing landed (e.g. oversized payload was refused) — re-try once.
            try:
                self.clipboard_clear()
                self.clipboard_append(expected)
            except Exception:
                pass

    def _ask_save_path(self, default_name):
        """Try the native Save dialog; if it breaks (common when the app has run
        a very large job and Tk's dialog state is corrupted), fall back to saving
        straight into the working directory so the data is never lost."""
        try:
            return filedialog.asksaveasfilename(
                defaultextension=os.path.splitext(default_name)[1],
                filetypes=[("All Files", "*.*"),
                           (os.path.splitext(default_name)[1].lstrip(".").upper(),
                            "*" + os.path.splitext(default_name)[1])],
                initialfile=default_name), False
        except Exception:
            # Native dialog failed (TclError: Unspecified error). Save directly.
            fallback = os.path.join(os.getcwd(), default_name)
            return fallback, True

    def _download_csv(self):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Run analysis first.")
            return
        # Truncate oversized cells so the CSV opens cleanly everywhere.
        cleaned = [_export_dict(r) for r in self.results]
        p, used_fallback = self._ask_save_path("image_analysis.csv")
        if not p:
            return
        try:
            with open(p, "w", newline="", encoding="utf-8-sig") as f:
                w = csv.DictWriter(f, fieldnames=self.EXPORT_COLS,
                                   extrasaction="ignore")
                w.writeheader()
                w.writerows(cleaned)
        except Exception as ex:
            messagebox.showerror("Export Error", str(ex))
            return
        note = " (save dialog was unavailable)" if used_fallback else ""
        self.status_bar.config(
            text=f"CSV saved: {os.path.basename(p)} ✓{note}")
        messagebox.showinfo("CSV Saved",
                            f"CSV saved to:\n{p}\n\n"
                            f"Rows: {len(cleaned)} · Columns: {len(self.EXPORT_COLS)}")

    def _download_excel(self):
        if not self.results:
            messagebox.showinfo("Nothing to export", "Run analysis first.")
            return
        p, used_fallback = self._ask_save_path("image_analysis.xlsx")
        if not p:
            return
        try:
            self._build_excel(p)
        except Exception as ex:
            messagebox.showerror("Export Error", str(ex))
            return
        note = " (save dialog was unavailable)" if used_fallback else ""
        self.status_bar.config(
            text=f"Excel saved: {os.path.basename(p)} ✓{note}")
        messagebox.showinfo("Excel Saved",
                            f"Excel saved to:\n{p}\n\n"
                            f"Rows: {len(self.results)} · Columns: {len(self.EXPORT_COLS)}")

    def _build_excel(self, path):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Image Analysis"

        # ── Legible light-theme colour scheme ──
        # Header: dark navy / slate, white text
        hdr_fill  = PatternFill("solid", fgColor="1E293B")
        hdr_font  = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
        # Odd rows: clean white, dark text
        odd_fill  = PatternFill("solid", fgColor="FFFFFF")
        odd_font  = Font(name="Calibri", color="111827", size=9)
        # Even rows: very light warm grey, dark text
        even_fill = PatternFill("solid", fgColor="F1F5F9")
        even_font = Font(name="Calibri", color="1E293B", size=9)
        # Error rows: soft red tint, dark red text
        err_fill  = PatternFill("solid", fgColor="FEE2E2")
        err_font  = Font(name="Calibri", color="991B1B", size=9)
        # Borders
        thin   = Side(style="thin",   color="CBD5E1")
        medium = Side(style="medium", color="64748B")
        hdr_bdr = Border(bottom=medium)
        row_bdr = Border(bottom=thin)

        # Header
        for ci, name in enumerate(self.EXPORT_COLS, 1):
            c = ws.cell(row=1, column=ci, value=name)
            c.font      = hdr_font
            c.fill      = hdr_fill
            c.border    = hdr_bdr
            c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 24
        ws.freeze_panes = "A2"

        # Data rows
        for ri, row in enumerate(self.results, 2):
            is_err = str(row.get("Description", "")).startswith("ERROR:")
            if is_err:
                fill, font = err_fill, err_font
            elif ri % 2 == 0:
                fill, font = even_fill, even_font
            else:
                fill, font = odd_fill, odd_font

            for ci, col in enumerate(self.EXPORT_COLS, 1):
                raw_val = _export_value(row.get(col, ""))
                c = ws.cell(row=ri, column=ci, value=raw_val)
                c.font      = font
                c.fill      = fill
                c.border    = row_bdr
                c.alignment = Alignment(
                    horizontal="left", vertical="top",
                    wrap_text=(col in ("Description", "Subject Tags",
                                       "Quality Notes", "Text Content",
                                       "Translation",
                                       "Metadata", "Path")))
            ws.row_dimensions[ri].height = 48

        # Column widths
        widths = {"Filename": 28, "Path": 44, "Dimensions": 16,
                  "Color Mode": 12, "File Size": 12, "Description": 64,
                  "Text Content": 54, "Translation": 54, "Photographer": 22,
                  "Location": 26, "GPS Coordinates": 26, "Date Taken": 20,
                  "Subject Tags": 34, "Quality Notes": 42, "Metadata": 70}
        for ci, col in enumerate(self.EXPORT_COLS, 1):
            ws.column_dimensions[
                openpyxl.utils.get_column_letter(ci)
            ].width = widths.get(col, 18)

        ws.auto_filter.ref = ws.dimensions

        wb.save(path)


# ── Entry point ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = ImageAnalyzerApp()
    app.mainloop()

