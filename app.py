import hashlib
import html
import io
import json
import re
import secrets
import sqlite3
import time
from datetime import datetime
from pathlib import Path

import easyocr
import fitz
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from bs4 import BeautifulSoup
from docx import Document
from ebooklib import ITEM_DOCUMENT, epub
from fpdf import FPDF
from gtts import gTTS
from PIL import Image, ImageEnhance, ImageOps
from pptx import Presentation
from striprtf.striprtf import rtf_to_text

try:
    import speech_recognition as sr
    SPEECH_RECOGNITION_AVAILABLE = True
except ImportError:
    SPEECH_RECOGNITION_AVAILABLE = False


# ============================================================
# 1. APPLICATION SETTINGS
# ============================================================

st.set_page_config(
    page_title="NaijaSight Portal | Assistive Document, Triage & Voice Suite",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_NAME = "NaijaSight"
APP_VERSION = "2.3-DSN"
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "naijasight_portal.db"

MAX_FILE_SIZE_MB = 500
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

MAX_IMAGE_PIXELS = 25_000_000
MAX_PDF_PAGES = 100

Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS


# ============================================================
# 2. EDITORIAL PORTAL, LOW-VISION & MOBILE STYLING
# ============================================================

def apply_portal_styles(vision_theme: str = "Standard", reader_font_px: int = 18):
    if vision_theme == "High Contrast (Print Ledger)":
        bg_color = "#FFFFFF"
        card_bg = "#FFFFFF"
        text_color = "#000000"
        muted_color = "#262626"
        border_color = "#000000"
        banner_bg = "#FFFFFF"
        banner_text = "#000000"
        banner_border = "#8B1E1E"
        tag_bg = "#F4F4F5"
        tag_text = "#000000"
        tag_border = "#000000"
        accent_color = "#8B1E1E"
        accent_hover = "#691515"
        ochre_color = "#8B1E1E"
        input_bg = "#FFFFFF"
    elif vision_theme == "Low-Glare Night (Warm Charcoal)":
        bg_color = "#161615"
        card_bg = "#201F1D"
        text_color = "#F4F1EA"
        muted_color = "#C7C2B8"
        border_color = "#3A3732"
        banner_bg = "#201F1D"
        banner_text = "#F4F1EA"
        banner_border = "#D97706"
        tag_bg = "#2C2A27"
        tag_text = "#F4F1EA"
        tag_border = "#524E48"
        accent_color = "#2F6F57"
        accent_hover = "#245744"
        ochre_color = "#D97706"
        input_bg = "#1A1917"
    else:
        # Editorial Stone & Botanical Forest — calm, tactile, non-AI aesthetic
        bg_color = "#F7F5F0"
        card_bg = "#FFFDF9"
        text_color = "#1C1917"
        muted_color = "#57534E"
        border_color = "#E4DFD5"
        banner_bg = "#FFFDF9"
        banner_text = "#1C1917"
        banner_border = "#1E4D3B"
        tag_bg = "#F2EFE9"
        tag_text = "#1E4D3B"
        tag_border = "#D6D0C4"
        accent_color = "#1E4D3B"
        accent_hover = "#163B2D"
        ochre_color = "#B45309"
        input_bg = "#FFFDF9"

    st.markdown(
        f"""
        <style>
        :root {{
            --portal-forest: #1E4D3B;
            --portal-ochre: {ochre_color};
            --portal-accent: {accent_color};
            --portal-accent-hover: {accent_hover};
            --portal-bg: {bg_color};
            --portal-card: {card_bg};
            --portal-border: {border_color};
            --portal-text: {text_color};
            --portal-muted: {muted_color};
            --portal-input-bg: {input_bg};
            --banner-bg: {banner_bg};
            --banner-text: {banner_text};
            --banner-border: {banner_border};
            --tag-bg: {tag_bg};
            --tag-text: {tag_text};
            --tag-border: {tag_border};
        }}

        [data-testid="stAppViewContainer"] {{
            background-color: var(--portal-bg);
            color: var(--portal-text);
        }}

        [data-testid="stAppViewContainer"] p,
        [data-testid="stAppViewContainer"] li,
        [data-testid="stAppViewContainer"] label,
        [data-testid="stAppViewContainer"] .stMarkdown {{
            color: var(--portal-text);
        }}

        .block-container {{
            max-width: 1180px;
            padding-top: 1.5rem;
            padding-bottom: 3.5rem;
        }}

        .portal-banner {{
            background: var(--banner-bg);
            color: var(--banner-text);
            padding: 1.3rem 1.55rem;
            border-radius: 12px;
            margin-bottom: 1.35rem;
            border: 1px solid var(--portal-border);
            border-left: 5px solid var(--banner-border);
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            align-items: center;
            gap: 1rem;
        }}

        .portal-banner h1,
        .portal-banner h2,
        .portal-banner p,
        .portal-banner strong {{
            color: var(--banner-text) !important;
            margin: 0 !important;
        }}

        .portal-banner p {{
            color: var(--portal-muted) !important;
            margin-top: 0.2rem !important;
            font-size: 0.95rem;
        }}

        .portal-tag {{
            background: var(--tag-bg);
            border: 1px solid var(--tag-border);
            color: var(--tag-text);
            padding: 0.32rem 0.72rem;
            border-radius: 6px;
            font-size: 0.8rem;
            font-weight: 600;
            letter-spacing: 0.01em;
        }}

        .kpi-card {{
            background: var(--portal-card);
            border: 1px solid var(--portal-border);
            border-top: 3px solid var(--portal-accent);
            border-radius: 10px;
            padding: 0.95rem 1.1rem;
            margin-bottom: 0.5rem;
        }}

        .kpi-label {{
            font-size: 0.75rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            color: var(--portal-muted);
        }}

        .kpi-value {{
            font-size: 1.45rem;
            font-weight: 700;
            color: var(--portal-text);
            margin-top: 0.2rem;
        }}

        .triage-box {{
            background: var(--portal-card);
            border: 1px solid var(--portal-border);
            border-left: 4px solid var(--portal-ochre);
            border-radius: 10px;
            padding: 1rem 1.2rem;
            margin-bottom: 1rem;
            color: var(--portal-text);
        }}

        div.stButton > button[kind="primary"] {{
            background-color: var(--portal-accent);
            border-color: var(--portal-accent);
            color: #FFFDF9;
        }}

        div.stButton > button[kind="primary"]:hover {{
            background-color: var(--portal-accent-hover);
            border-color: var(--portal-accent-hover);
            color: #FFFDF9;
        }}

        div.stButton > button,
        div.stDownloadButton > button {{
            border-radius: 8px;
            min-height: 46px;
            font-weight: 600;
        }}

        textarea,
        [data-testid="stTextArea"] textarea {{
            font-size: {reader_font_px}px !important;
            line-height: 1.65 !important;
            background-color: var(--portal-input-bg) !important;
            color: var(--portal-text) !important;
            border-color: var(--portal-border) !important;
        }}

        button:focus-visible,
        input:focus-visible,
        textarea:focus-visible,
        select:focus-visible {{
            outline: 3px solid var(--portal-ochre) !important;
            outline-offset: 2px !important;
        }}

        /* Mobile & Phone Responsiveness */
        @media (max-width: 768px) {{
            .block-container {{
                padding-left: 0.85rem;
                padding-right: 0.85rem;
                padding-top: 1rem;
            }}

            .portal-banner {{
                padding: 1.1rem;
                flex-direction: column;
                align-items: flex-start;
            }}

            div.stButton > button,
            div.stDownloadButton > button {{
                min-height: 52px;
                width: 100%;
                white-space: normal;
            }}

            [data-testid="stFileUploader"] {{
                width: 100%;
            }}
        }}
        </style>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# 3. DATABASE & DATA RETENTION ENGINE
# ============================================================

def get_connection():
    connection = sqlite3.connect(DB_PATH, check_same_thread=False)
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with get_connection() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                password_salt TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS activity (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                filename TEXT NOT NULL,
                file_type TEXT NOT NULL,
                word_count INTEGER NOT NULL,
                character_count INTEGER NOT NULL,
                extracted_text TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )

        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS accessibility_feedback (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL,
                tester_group TEXT NOT NULL,
                issue_found TEXT NOT NULL,
                fix_applied TEXT NOT NULL,
                severity TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
        )

        columns = [
            row["name"]
            for row in conn.execute("PRAGMA table_info(activity)").fetchall()
        ]
        if "extracted_text" not in columns:
            conn.execute(
                "ALTER TABLE activity ADD COLUMN extracted_text TEXT NOT NULL DEFAULT ''"
            )


init_db()


# ============================================================
# 4. PASSWORD SECURITY
# ============================================================

def hash_password(password, salt=None):
    if salt is None:
        salt = secrets.token_hex(16)

    password_hash = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        bytes.fromhex(salt),
        300_000
    )

    return password_hash.hex(), salt


def verify_password(password, stored_hash, salt):
    calculated_hash, _ = hash_password(password, salt)
    return secrets.compare_digest(calculated_hash, stored_hash)


# ============================================================
# 5. ACCOUNT FUNCTIONS
# ============================================================

def create_account(full_name, username, password):
    full_name = full_name.strip()
    username = username.strip().lower()

    if not full_name or not username or not password:
        return False, "Please complete all fields."

    if len(username) < 3:
        return False, "Username must contain at least 3 characters."

    if not re.fullmatch(r"[a-zA-Z0-9_.-]+", username):
        return False, "Use letters, numbers, dots, underscores or hyphens."

    if len(password) < 8:
        return False, "Password must contain at least 8 characters."

    password_hash, salt = hash_password(password)

    try:
        with get_connection() as conn:
            conn.execute(
                """
                INSERT INTO users (
                    full_name,
                    username,
                    password_hash,
                    password_salt,
                    created_at
                )
                VALUES (?, ?, ?, ?, ?)
                """,
                (
                    full_name,
                    username,
                    password_hash,
                    salt,
                    datetime.now().isoformat(timespec="seconds")
                )
            )
        return True, "Your account has been created."

    except sqlite3.IntegrityError:
        return False, "That username is already in use."


def authenticate_user(username, password):
    username = username.strip().lower()

    with get_connection() as conn:
        user = conn.execute(
            "SELECT * FROM users WHERE username = ?",
            (username,)
        ).fetchone()

    if user is None:
        return None

    if verify_password(password, user["password_hash"], user["password_salt"]):
        return {
            "id": user["id"],
            "full_name": user["full_name"],
            "username": user["username"]
        }

    return None


# ============================================================
# 6. PERSISTENT ACTIVITY, FEEDBACK, EXPORTS & DATA RETENTION
# ============================================================

def log_activity(filename, file_type, text):
    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")
    if not username:
        return

    clean_text = text or ""
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO activity (
                username,
                filename,
                file_type,
                word_count,
                character_count,
                extracted_text,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                filename,
                file_type,
                len(clean_text.split()),
                len(clean_text),
                clean_text,
                "Completed",
                datetime.now().isoformat(timespec="seconds")
            )
        )


def get_activity_records():
    if st.session_state.get("demo_mode"):
        return []

    username = st.session_state.get("username")
    if not username:
        return []

    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id, filename, file_type, word_count,
                   character_count, extracted_text, status, created_at
            FROM activity
            WHERE username = ?
            ORDER BY id DESC
            """,
            (username,)
        ).fetchall()

    return [dict(row) for row in rows]


def get_activity():
    records = get_activity_records()
    if not records:
        return pd.DataFrame()

    summary_rows = [
        {
            "ID": r["id"],
            "Filename": r["filename"],
            "Type": r["file_type"],
            "Words": r["word_count"],
            "Characters": r["character_count"],
            "Status": r["status"],
            "Timestamp": r["created_at"],
        }
        for r in records
    ]
    return pd.DataFrame(summary_rows)


def delete_activity_item(record_id):
    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")
    if not username:
        return

    with get_connection() as conn:
        conn.execute(
            "DELETE FROM activity WHERE id = ? AND username = ?",
            (int(record_id), username)
        )


def delete_activity():
    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")
    if not username:
        return

    with get_connection() as conn:
        conn.execute(
            "DELETE FROM activity WHERE username = ?",
            (username,)
        )


def log_accessibility_feedback(tester_group: str, issue_found: str, fix_applied: str, severity: str):
    username = st.session_state.get("username") or "guest"
    with get_connection() as conn:
        conn.execute(
            """
            INSERT INTO accessibility_feedback (
                username, tester_group, issue_found, fix_applied, severity, created_at
            )
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                tester_group.strip(),
                issue_found.strip(),
                fix_applied.strip(),
                severity,
                datetime.now().isoformat(timespec="seconds")
            )
        )


def get_accessibility_feedback_df() -> pd.DataFrame:
    with get_connection() as conn:
        rows = conn.execute(
            """
            SELECT id AS ID, username AS Logged_By, tester_group AS Tester_Group,
                   severity AS Severity, issue_found AS Issue_Observed,
                   fix_applied AS Fix_Implemented, created_at AS Timestamp
            FROM accessibility_feedback
            ORDER BY id DESC
            """
        ).fetchall()
    return pd.DataFrame([dict(row) for row in rows])


def build_excel_export_bytes(records_override=None):
    records = records_override if records_override is not None else get_activity_records()
    summary_df = pd.DataFrame([
        {
            "Record_ID": r["id"],
            "Account_Username": st.session_state.get("username"),
            "Account_Name": st.session_state.get("full_name"),
            "Filename": r["filename"],
            "File_Type": r["file_type"],
            "Word_Count": r["word_count"],
            "Character_Count": r["character_count"],
            "Status": r["status"],
            "Created_At": r["created_at"],
        }
        for r in records
    ])

    texts_df = pd.DataFrame([
        {
            "Record_ID": r["id"],
            "Filename": r["filename"],
            "Created_At": r["created_at"],
            "Extracted_Text": r["extracted_text"],
        }
        for r in records
    ])

    feedback_df = get_accessibility_feedback_df()

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Activity_Log")
        texts_df.to_excel(writer, index=False, sheet_name="Saved_Texts")
        feedback_df.to_excel(writer, index=False, sheet_name="Accessibility_Testing")
    return output.getvalue()


def build_audit_json_bytes(records: list[dict]) -> bytes:
    feedback_records = get_accessibility_feedback_df().to_dict(orient="records")
    payload = {
        "application": APP_NAME,
        "version": APP_VERSION,
        "exported_at": datetime.now().isoformat(timespec="seconds"),
        "account": {
            "username": st.session_state.get("username"),
            "full_name": st.session_state.get("full_name"),
        },
        "limits": {
            "max_file_size_mb": MAX_FILE_SIZE_MB,
            "max_image_pixels": MAX_IMAGE_PIXELS,
            "max_pdf_pages": MAX_PDF_PAGES,
        },
        "activity_records": records,
        "accessibility_feedback": feedback_records,
    }
    return json.dumps(payload, indent=2, ensure_ascii=False).encode("utf-8")


def _sanitize_pdf_text(value: str) -> str:
    text = str(value or "")
    replacements = {
        "\u2014": "-",
        "\u2013": "-",
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2022": "*",
        "\u2026": "...",
        "\u20a6": "NGN ",
    }
    for src, dst in replacements.items():
        text = text.replace(src, dst)
    return text.encode("latin-1", errors="replace").decode("latin-1")


def build_pdf_report_bytes(records: list[dict]) -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_fill_color(30, 77, 59)
    pdf.set_text_color(255, 253, 249)
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 12, _sanitize_pdf_text("NaijaSight Portal - Extraction & Activity Report"), ln=True, fill=True)

    pdf.set_text_color(28, 25, 23)
    pdf.set_font("Helvetica", "", 10)
    pdf.ln(3)
    user_label = f"{st.session_state.get('full_name', '')} (@{st.session_state.get('username', '')})"
    generated_stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pdf.cell(0, 6, _sanitize_pdf_text(f"Account: {user_label}   |   Generated: {generated_stamp}"), ln=True)
    pdf.cell(0, 6, _sanitize_pdf_text(f"Records Included: {len(records)}"), ln=True)
    pdf.ln(4)

    for idx, rec in enumerate(records, start=1):
        pdf.set_draw_color(214, 208, 196)
        pdf.set_fill_color(242, 239, 233)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(30, 77, 59)
        title_line = (
            f"{idx}. {rec['filename']}  [{rec['file_type'] or 'file'}]  "
            f"| {rec['word_count']} words | {rec['created_at']}"
        )
        pdf.multi_cell(0, 7, _sanitize_pdf_text(title_line), border=1, fill=True)

        body = (rec.get("extracted_text") or "").strip()
        if not body:
            body = "(No text content stored for this record.)"
        elif len(body) > 2500:
            body = body[:2500] + "\n... [Truncated in PDF preview; full text included in TXT/Excel export]"

        pdf.set_font("Helvetica", "", 10)
        pdf.set_text_color(28, 25, 23)
        pdf.multi_cell(0, 5.5, _sanitize_pdf_text(body), border="LRB")
        pdf.ln(4)

    raw_out = pdf.output()
    return bytes(raw_out) if isinstance(raw_out, (bytes, bytearray)) else raw_out.encode("latin-1")


# ============================================================
# 7. SESSION STATE
# ============================================================

defaults = {
    "authenticated": False,
    "username": None,
    "full_name": None,
    "demo_mode": False,
    "nav_page": "Home",

    "extracted_text": "",
    "current_filename": "",
    "current_file_type": "",

    "last_ocr_confidence": None,
    "last_uncertain_words": [],
    "last_process_ms": None,

    "editable_result": "",
    "vision_theme": "Standard",
    "reader_font_px": 18,
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


def clear_current_result():
    st.session_state.extracted_text = ""
    st.session_state.current_filename = ""
    st.session_state.current_file_type = ""
    st.session_state.last_ocr_confidence = None
    st.session_state.last_uncertain_words = []
    st.session_state.last_process_ms = None
    st.session_state.editable_result = ""


def logout():
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.full_name = None
    st.session_state.demo_mode = False
    st.session_state.nav_page = "Home"
    clear_current_result()
    st.rerun()


# ============================================================
# 8. LOAD OCR ENGINE
# ============================================================

@st.cache_resource(show_spinner="Loading text recognition engine...")
def load_ocr_reader():
    return easyocr.Reader(["en"], gpu=False)


# ============================================================
# 9. FILE SIZE CHECK
# ============================================================

def check_file_size(uploaded_file):
    if uploaded_file.size > MAX_FILE_SIZE_BYTES:
        raise ValueError(
            f"This file exceeds the {MAX_FILE_SIZE_MB} MB upload limit."
        )
    return uploaded_file.getvalue()


# ============================================================
# 10. IMAGE ENHANCEMENT & LAYOUT GROUPING
# ============================================================

def prepare_image(image):
    image = ImageOps.exif_transpose(image)
    if image.width * image.height > MAX_IMAGE_PIXELS:
        raise ValueError("Image is too large to process safely.")
    return image.convert("RGB")


def crop_main_document(image):
    image = prepare_image(image)
    width, height = image.size
    if width / max(height, 1) >= 1.25 and width >= 1000:
        left = int(width * 0.15)
        top = int(height * 0.22)
        right = int(width * 0.85)
        bottom = int(height * 0.96)
        if right - left > 100 and bottom - top > 100:
            return image.crop((left, top, right, bottom))
    return image


def enhance_for_ocr(image):
    image = prepare_image(image)
    width, height = image.size

    longest_side = max(width, height)
    if longest_side < 1800:
        scale = min(1.6, 1800 / max(longest_side, 1))
        image = image.resize(
            (max(1, int(width * scale)), max(1, int(height * scale))),
            Image.Resampling.LANCZOS
        )

    enhanced = ImageOps.grayscale(image)
    enhanced = ImageOps.autocontrast(enhanced, cutoff=1)
    enhanced = ImageEnhance.Contrast(enhanced).enhance(1.25)
    enhanced = ImageEnhance.Sharpness(enhanced).enhance(1.35)
    return enhanced.convert("RGB")


def group_ocr_results(results):
    items = []
    heights = []
    for box, detected_text, confidence in results:
        detected_text = detected_text.strip()
        if not detected_text:
            continue
        xs = [point[0] for point in box]
        ys = [point[1] for point in box]
        x0, y0 = min(xs), min(ys)
        x1, y1 = max(xs), max(ys)
        height = max(1, y1 - y0)
        heights.append(height)
        items.append({
            "text": detected_text,
            "confidence": float(confidence),
            "x": x0,
            "cy": (y0 + y1) / 2,
            "height": height,
        })

    if not items:
        return [], []

    tolerance = max(8, float(np.median(heights)) * 0.65)
    items.sort(key=lambda item: (item["cy"], item["x"]))
    lines = []
    for item in items:
        target = None
        for line in reversed(lines[-3:]):
            line_cy = sum(part["cy"] for part in line) / len(line)
            if abs(item["cy"] - line_cy) <= tolerance:
                target = line
                break
        if target is None:
            lines.append([item])
        else:
            target.append(item)

    lines.sort(key=lambda line: min(part["cy"] for part in line))
    ordered = []
    for line in lines:
        line.sort(key=lambda item: item["x"])
        ordered.append(" ".join(item["text"] for item in line))

    flat_items = [item for line in lines for item in line]
    return ordered, flat_items


def recognize_image(image, mode="Read Entire Image", quality_mode="Standard"):
    prepared = prepare_image(image)

    if mode == "Read Main Document":
        prepared = crop_main_document(prepared)

    if quality_mode == "Enhanced":
        prepared = enhance_for_ocr(prepared)

    reader = load_ocr_reader()
    results = reader.readtext(
        np.array(prepared),
        detail=1,
        paragraph=False,
        decoder="greedy",
        text_threshold=0.65,
        low_text=0.35,
        link_threshold=0.45,
        batch_size=1,
        workers=0,
        canvas_size=1600,
        mag_ratio=1.0,
    )

    lines, items = group_ocr_results(results)
    if not items:
        return {"text": "", "confidence": 0, "uncertain_words": []}

    confidences = [item["confidence"] for item in items]
    uncertain_words = [
        {"text": item["text"], "confidence": round(item["confidence"] * 100, 1)}
        for item in items if item["confidence"] < 0.60
    ]
    return {
        "text": "\n".join(lines),
        "confidence": round(sum(confidences) / len(confidences) * 100, 1),
        "uncertain_words": uncertain_words,
    }


def image_to_text(image, mode="Read Entire Image", quality_mode="Standard"):
    result = recognize_image(image, mode=mode, quality_mode=quality_mode)
    st.session_state.last_ocr_confidence = result["confidence"]
    st.session_state.last_uncertain_words = result["uncertain_words"]
    return result["text"]


# ============================================================
# 12 - 20. DOCUMENT FORMAT PROCESSORS
# ============================================================

def process_image(uploaded_file, mode="Read Entire Image", quality_mode="Standard"):
    file_bytes = check_file_size(uploaded_file)
    image = Image.open(io.BytesIO(file_bytes))
    return image_to_text(image, mode=mode, quality_mode=quality_mode)


def process_pdf(uploaded_file, quality_mode="Standard"):
    file_bytes = check_file_size(uploaded_file)
    document = fitz.open(stream=file_bytes, filetype="pdf")

    if document.page_count > MAX_PDF_PAGES:
        document.close()
        raise ValueError(f"This PDF contains more than {MAX_PDF_PAGES} pages.")

    extracted_pages = []
    ocr_confidences = []
    all_uncertain_words = []
    progress = st.progress(0)

    try:
        for page_number, page in enumerate(document):
            page_text = page.get_text().strip()

            if len(page_text) < 15:
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                image = Image.open(io.BytesIO(pix.tobytes("png")))
                ocr_result = recognize_image(image, quality_mode=quality_mode)
                page_text = ocr_result["text"]
                ocr_confidences.append(ocr_result["confidence"])
                all_uncertain_words.extend(ocr_result["uncertain_words"])

            extracted_pages.append(f"PAGE {page_number + 1}\n{page_text}")
            progress.progress((page_number + 1) / document.page_count)
    finally:
        document.close()
        progress.empty()

    if ocr_confidences:
        st.session_state.last_ocr_confidence = round(
            sum(ocr_confidences) / len(ocr_confidences), 1
        )
        st.session_state.last_uncertain_words = all_uncertain_words

    return "\n\n".join(extracted_pages).strip()


def process_docx(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    document = Document(io.BytesIO(file_bytes))
    paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]

    for table in document.tables:
        for row in table.rows:
            row_text = [cell.text.strip() for cell in row.cells if cell.text.strip()]
            if row_text:
                paragraphs.append(" | ".join(row_text))

    return "\n".join(paragraphs).strip()


def process_pptx(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    presentation = Presentation(io.BytesIO(file_bytes))
    slides_text = []

    for index, slide in enumerate(presentation.slides):
        slide_parts = []
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_parts.append(shape.text.strip())
            if shape.has_table:
                for row in shape.table.rows:
                    row_vals = [c.text.strip() for c in row.cells if c.text.strip()]
                    if row_vals:
                        slide_parts.append(" | ".join(row_vals))
        slides_text.append(f"SLIDE {index + 1}\n" + "\n".join(slide_parts))

    return "\n\n".join(slides_text).strip()


def process_xlsx(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    output = []

    for sheet_name in workbook.sheet_names:
        dataframe = pd.read_excel(workbook, sheet_name=sheet_name)
        output.append(
            f"SHEET: {sheet_name}\n" + dataframe.fillna("").to_string(index=False)
        )

    return "\n\n".join(output).strip()


def process_csv(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    dataframe = pd.read_csv(io.BytesIO(file_bytes))
    return dataframe.fillna("").to_string(index=False)


def process_txt(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    try:
        return file_bytes.decode("utf-8").strip()
    except UnicodeDecodeError:
        return file_bytes.decode("latin-1", errors="replace").strip()


def process_rtf(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    raw_text = file_bytes.decode("utf-8", errors="replace")
    return rtf_to_text(raw_text).strip()


def process_epub(uploaded_file):
    file_bytes = check_file_size(uploaded_file)
    book = epub.read_epub(io.BytesIO(file_bytes))
    chapters = []

    for item in book.get_items():
        if item.get_type() == ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_content(), "html.parser")
            chapter_text = soup.get_text(separator="\n").strip()
            if chapter_text:
                chapters.append(chapter_text)

    return "\n\n".join(chapters).strip()


# ============================================================
# 21. DOCUMENT PROCESSING ROUTER & SMART TRIAGE ENGINE
# ============================================================

SAMPLE_DOCUMENTS = {
    "⚡ EKEDC Electricity Utility Bill": (
        "EKO ELECTRICITY DISTRIBUTION COMPANY (EKEDC)\n"
        "Account Number: 0412894510-01 | Meter No: 45091827364\n"
        "Customer Name: Adebayo Ogunleye | Tariff: Band A (Non-MD)\n"
        "Billing Period: October 2026 | Previous Balance: NGN 4,200.00\n"
        "Current Energy Charge: NGN 28,450.00 | VAT (7.5%): NGN 2,133.75\n"
        "Total Amount Due: NGN 34,783.75\n"
        "Due Date: 28/11/2026\n"
        "WARNING: Disconnection notice applies 7 days after due date if unpaid. "
        "Pay only via official EKEDC channels or USSD *737#."
    ),
    "💊 Pharmacy Prescription & Dosage Label": (
        "LAGOS UNIVERSITY TEACHING HOSPITAL (LUTH) PHARMACY\n"
        "Patient: Chidinma Okafor | Date: 21/11/2026\n"
        "Medication 1: Amoxicillin 500mg Capsules - Take 1 capsule every 8 hours for 7 days after meals.\n"
        "Medication 2: Paracetamol 500mg Tablets - Take 2 tablets every 6 hours if fever exceeds 38C.\n"
        "CAUTION: Complete the full antibiotic course. Do not exceed 4,000mg Paracetamol in 24 hours. "
        "Keep out of reach of children. Contact +234-800-555-0199 if rash or dizziness occurs."
    ),
    "🏦 Bank Transfer & POS Dispute Notice": (
        "ACCESS BANK NIGERIA - CUSTOMER TRANSACTION ADVICE\n"
        "Reference ID: TXN-994810234 | Date: 20/11/2026\n"
        "Transaction Type: POS / NIP Instant Transfer Reversal Claim\n"
        "Disputed Amount: NGN 45,000.00 | Merchant Terminal: 2044LA09\n"
        "Resolution Deadline: 25/11/2026\n"
        "IMPORTANT NOTICE: Never share your card PIN, BVN, or OTP with anyone. "
        "Your claim will be auto-resolved within 48 working hours."
    ),
}


def process_uploaded_document(
    uploaded_file,
    image_mode="Read Entire Image",
    quality_mode="Standard"
):
    extension = Path(uploaded_file.name).suffix.lower()

    processors = {
        ".png": process_image,
        ".jpg": process_image,
        ".jpeg": process_image,
        ".webp": process_image,
        ".bmp": process_image,
        ".tif": process_image,
        ".tiff": process_image,
        ".pdf": process_pdf,
        ".docx": process_docx,
        ".pptx": process_pptx,
        ".xlsx": process_xlsx,
        ".csv": process_csv,
        ".txt": process_txt,
        ".rtf": process_rtf,
        ".epub": process_epub
    }

    processor = processors.get(extension)
    if processor is None:
        raise ValueError(f"The file format {extension} is not supported.")

    if extension in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        return processor(uploaded_file, mode=image_mode, quality_mode=quality_mode)
    if extension == ".pdf":
        return processor(uploaded_file, quality_mode=quality_mode)
    return processor(uploaded_file)


def analyze_document_triage(text: str) -> dict:
    clean = (text or "").strip()
    if not clean:
        return {}

    amounts = re.findall(
        r"(?:NGN|₦|USD|\$)\s?\d[\d,]*(?:\.\d{1,2})?|\b\d[\d,]*(?:\.\d{1,2})?\s?(?:NGN|naira|USD)\b",
        clean,
        flags=re.IGNORECASE
    )

    dates = re.findall(
        r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}-\d{2}-\d{2}|"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},?\s+\d{2,4})\b",
        clean,
        flags=re.IGNORECASE
    )

    dosages = re.findall(
        r"\b\d+\s?(?:mg|ml|mcg|g|tablet[s]?|capsule[s]?|drop[s]?)\b[^.\n]*",
        clean,
        flags=re.IGNORECASE
    )

    identifiers = re.findall(
        r"(?:Account|Meter|Ref(?:erence)?|ID|Policy|Phone|Tel)[\s#:No.]*[A-Z0-9+-]{6,}",
        clean,
        flags=re.IGNORECASE
    )

    warning_keywords = (
        "warning", "caution", "danger", "due date", "deadline", "disconnection",
        "do not", "never", "important", "penalty", "expire", "overdose", "emergency"
    )
    alerts = []
    for line in clean.splitlines():
        low = line.lower()
        if any(kw in low for kw in warning_keywords):
            alerts.append(line.strip())

    def dedup(seq):
        seen = set()
        out = []
        for item in seq:
            k = item.strip().lower()
            if k and k not in seen:
                seen.add(k)
                out.append(item.strip())
        return out

    amounts = dedup(amounts)[:8]
    dates = dedup(dates)[:6]
    dosages = dedup(dosages)[:6]
    identifiers = dedup(identifiers)[:6]
    alerts = dedup(alerts)[:5]

    spoken_parts = ["Key facts summary."]
    if amounts:
        spoken_parts.append(f"Amounts detected: {', '.join(amounts[:4])}.")
    if dates:
        spoken_parts.append(f"Important dates: {', '.join(dates[:3])}.")
    if dosages:
        spoken_parts.append(f"Dosage instructions: {'; '.join(dosages[:2])}.")
    if alerts:
        spoken_parts.append(f"Critical alert: {alerts[0]}")
    if len(spoken_parts) == 1:
        first_sentences = " ".join(re.split(r"(?<=[.!?])\s+", clean)[:2])
        spoken_parts.append(first_sentences[:350])

    return {
        "amounts": amounts,
        "dates": dates,
        "dosages": dosages,
        "identifiers": identifiers,
        "alerts": alerts,
        "spoken_brief": " ".join(spoken_parts),
    }


def render_instant_browser_tts(text: str, button_label: str = "🔊 Read Aloud Offline (Zero Data)"):
    """Client-side Web Speech API fallback so blind users can listen offline with zero latency."""
    payload = json.dumps(str(text or "")[:3500], ensure_ascii=False)
    payload = payload.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    safe_label = html.escape(button_label)

    components.html(
        f"""
        <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; font-family:Arial,sans-serif;">
            <button id="speakBtn" style="background:#1E4D3B;color:#FFFDF9;border:none;border-radius:8px;padding:10px 14px;font-weight:600;cursor:pointer;">
                {safe_label}
            </button>
            <button id="stopBtn" style="background:#F2EFE9;color:#1C1917;border:1px solid #C8C1B4;border-radius:8px;padding:10px 14px;font-weight:600;cursor:pointer;">
                ⏹ Stop
            </button>
            <span id="ttsStatus" style="font-size:13px;color:#57534E;"></span>
        </div>
        <script>
        const txt = {payload};
        const statusEl = document.getElementById('ttsStatus');
        document.getElementById('speakBtn').onclick = () => {{
            if (!('speechSynthesis' in window)) {{
                statusEl.textContent = 'Browser speech unavailable.';
                return;
            }}
            window.speechSynthesis.cancel();
            const u = new SpeechSynthesisUtterance(txt);
            u.onstart = () => statusEl.textContent = 'Reading offline...';
            u.onend = () => statusEl.textContent = 'Done.';
            window.speechSynthesis.speak(u);
        }};
        document.getElementById('stopBtn').onclick = () => {{
            if ('speechSynthesis' in window) window.speechSynthesis.cancel();
            statusEl.textContent = 'Stopped.';
        }};
        </script>
        """,
        height=56,
    )


def render_triage_panel(text: str, key_prefix: str):
    triage = analyze_document_triage(text)
    if not triage:
        return

    st.markdown(
        """
        <div class="triage-box">
            <strong>⚡ Key Facts First (Bill, Prescription & Notice Triage)</strong><br>
            <span>Critical figures, dates, dosage flags, and warnings surfaced before full-text reading.</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    t1, t2, t3 = st.columns(3)
    with t1:
        st.markdown("**💰 Amounts & Figures**")
        if triage["amounts"]:
            for amt in triage["amounts"]:
                st.write(f"- {amt}")
        else:
            st.caption("None detected")

    with t2:
        st.markdown("**📅 Dates & Deadlines**")
        if triage["dates"]:
            for dt in triage["dates"]:
                st.write(f"- {dt}")
        else:
            st.caption("None detected")

    with t3:
        st.markdown("**🔖 IDs, Phones & Dosage**")
        combined_ids = triage["identifiers"] + triage["dosages"]
        if combined_ids:
            for item in combined_ids[:5]:
                st.write(f"- {item}")
        else:
            st.caption("None detected")

    if triage["alerts"]:
        for alert in triage["alerts"]:
            st.warning(f"⚠️ **Alert / Clause:** {alert}")

    render_instant_browser_tts(
        triage["spoken_brief"],
        button_label="🔊 Speak Key Facts Offline (Instant / Zero Data)"
    )

    with st.expander("🎧 Download Key Facts MP3 Brief (Cloud TTS)"):
        st.write(triage["spoken_brief"])
        brief_consent = st.checkbox(
            "I agree to send this short key-facts summary to Google TTS.",
            key=f"{key_prefix}_brief_consent"
        )
        if st.button(
            "Generate Key Facts MP3",
            key=f"{key_prefix}_brief_btn",
            disabled=not brief_consent
        ):
            try:
                with st.spinner("Generating key facts audio..."):
                    audio = create_speech(triage["spoken_brief"], slow=False)
                st.audio(audio, format="audio/mp3")
            except Exception as err:
                st.error(f"Could not generate summary audio: {err}")


# ============================================================
# 22. TEXT TO SPEECH & OFFLINE READER
# ============================================================

def create_speech(text, slow=False):
    if not text.strip():
        raise ValueError("There is no text to read.")
    if len(text) > 4500:
        raise ValueError("This reading part is too long. Choose a smaller paragraph or part.")
    audio_buffer = io.BytesIO()
    speech = gTTS(text=text, lang="en", slow=slow)
    speech.write_to_fp(audio_buffer)
    audio_buffer.seek(0)
    return audio_buffer


def split_text_for_speech(text, max_chars=4000):
    text = (text or "").strip()
    if not text:
        return []
    paragraphs = [part.strip() for part in re.split(r"\n\s*\n", text) if part.strip()]
    sections = []
    for paragraph in paragraphs:
        if len(paragraph) <= max_chars:
            sections.append(paragraph)
            continue
        sentences = re.split(r"(?<=[.!?])\s+", paragraph)
        current = ""
        for sentence in sentences:
            sentence = sentence.strip()
            if not sentence:
                continue
            if len(sentence) > max_chars:
                if current:
                    sections.append(current)
                    current = ""
                words = sentence.split()
                word_part = ""
                for word in words:
                    candidate = f"{word_part} {word}".strip()
                    if len(candidate) > max_chars and word_part:
                        sections.append(word_part)
                        word_part = word
                    else:
                        word_part = candidate
                if word_part:
                    sections.append(word_part)
                continue
            candidate = f"{current} {sentence}".strip()
            if len(candidate) > max_chars and current:
                sections.append(current)
                current = sentence
            else:
                current = candidate
        if current:
            sections.append(current)
    return sections


def render_speech_controls(text, key_prefix):
    sections = split_text_for_speech(text)
    if not sections:
        st.info("There is no text available to read aloud yet.")
        return

    st.caption(
        f"This text is divided into {len(sections)} reading part(s). "
        "Choose one part at a time instead of generating the entire document."
    )
    selected_part = st.selectbox(
        "Choose a paragraph or reading part",
        options=list(range(len(sections))),
        format_func=lambda index: f"Part {index + 1} of {len(sections)}",
        key=f"{key_prefix}_speech_part"
    )
    speed = st.radio(
        "Speech speed",
        ["Normal", "Slower"],
        horizontal=True,
        key=f"{key_prefix}_speech_speed"
    )
    with st.expander("Preview selected reading part"):
        st.write(sections[selected_part])

    render_instant_browser_tts(
        sections[selected_part],
        button_label=f"🔊 Read Part {selected_part + 1} Offline (Browser Voice)"
    )

    st.warning(
        "Generating downloadable MP3 speech sends the selected text to Google's text-to-speech service. "
        "Do not submit private or sensitive information."
    )
    consent = st.checkbox(
        "I agree to send this text to Google to create MP3 speech.",
        key=f"{key_prefix}_tts_consent"
    )
    if st.button(
        "Generate MP3 speech for this part",
        key=f"{key_prefix}_generate_speech",
        type="primary",
        disabled=not consent
    ):
        try:
            with st.spinner("Preparing this reading part..."):
                audio = create_speech(sections[selected_part], slow=(speed == "Slower"))
            st.audio(audio, format="audio/mp3")
            st.download_button(
                "Download this audio part",
                data=audio.getvalue(),
                file_name=f"naijasight_audio_part_{selected_part + 1}.mp3",
                mime="audio/mpeg",
                key=f"{key_prefix}_download_audio"
            )
        except Exception as error:
            st.error(f"Could not create speech: {error}")


def build_offline_reader_html(text, document_title="NaijaSight Offline Reader"):
    safe_title = html.escape(str(document_title or "NaijaSight Offline Reader"))
    embedded_text = json.dumps(str(text or ""), ensure_ascii=False)
    embedded_text = (
        embedded_text.replace("<", "\\u003c")
        .replace(">", "\\u003e")
        .replace("&", "\\u0026")
    )

    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#1E4D3B">
<title>{safe_title}</title>
<style>
:root {{ color-scheme: light; --bg:#f7f5f0; --panel:#fffdf9; --ink:#1c1917; --accent:#1e4d3b; --muted:#57534e; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family:Arial, sans-serif; line-height:1.65; }}
header {{ background:#1e4d3b; color:#fffdf9; padding:20px 16px; }}
header h1 {{ margin:0 auto; max-width:900px; font-size:1.45rem; }}
main {{ max-width:900px; margin:20px auto; padding:0 14px 40px; }}
.panel {{ background:var(--panel); border:1px solid #e4dfd5; border-radius:12px; padding:16px; margin-bottom:16px; }}
.controls {{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; }}
button, select {{ min-height:44px; border-radius:8px; border:1px solid #c8c1b4; padding:9px 13px; font-size:1rem; background:#fffdf9; color:#1c1917; }}
button.primary {{ background:var(--accent); color:#fffdf9; border-color:var(--accent); }}
button:focus, select:focus, input:focus {{ outline:3px solid #b45309; outline-offset:2px; }}
label {{ display:block; font-weight:600; margin:8px 0 4px; }}
#readerText {{ white-space:pre-wrap; overflow-wrap:anywhere; font-size:1.2rem; line-height:1.8; }}
#status {{ color:var(--muted); min-height:1.5em; }}
.small {{ color:var(--muted); font-size:.92rem; }}
body.high-contrast {{ --bg:#ffffff; --panel:#ffffff; --ink:#000000; --accent:#8b1e1e; --muted:#262626; color:#000000; }}
body.high-contrast header {{ background:#ffffff; color:#000000; border-bottom:2px solid #000000; }}
body.high-contrast button, body.high-contrast select {{ background:#ffffff; color:#000000; border:2px solid #000000; }}
body.high-contrast button.primary {{ background:#8b1e1e; color:#ffffff; }}
@media (max-width:600px) {{
  header h1 {{ font-size:1.2rem; }}
  main {{ margin:12px auto; padding:0 10px 28px; }}
  .panel {{ padding:13px; }}
  .controls button, .controls select {{ flex:1 1 42%; }}
  #readerText {{ font-size:1.12rem; }}
}}
</style>
</head>
<body>
<header><h1>NaijaSight — Offline Reader</h1></header>
<main>
  <section class="panel" aria-label="Reading controls">
    <p class="small">This file contains your saved text. After downloading it, you can reopen it without an internet connection.</p>
    <div class="controls">
      <button type="button" id="smaller">A− Smaller text</button>
      <button type="button" id="larger">A+ Larger text</button>
      <button type="button" id="contrast">High contrast</button>
    </div>
    <label for="partSelect">Choose a reading section</label>
    <select id="partSelect" style="width:100%"></select>
    <div class="controls" style="margin-top:12px">
      <button class="primary" type="button" id="read">Read section</button>
      <button type="button" id="pause">Pause</button>
      <button type="button" id="resume">Resume</button>
      <button type="button" id="stop">Stop</button>
    </div>
    <p id="status" role="status" aria-live="polite">Your saved document is ready.</p>
  </section>
  <section class="panel">
    <h2 id="sectionHeading">Document text</h2>
    <article id="readerText" tabindex="0"></article>
  </section>
</main>
<script>
const originalText = {embedded_text};
const partSelect = document.getElementById('partSelect');
const readerText = document.getElementById('readerText');
const statusLine = document.getElementById('status');
const sectionHeading = document.getElementById('sectionHeading');
let parts = [];
let fontSize = 1.2;

function splitIntoParts(text) {{
  const paragraphs = text.split(/\\n\\s*\\n/).map(p => p.trim()).filter(Boolean);
  const output = [];
  let current = '';
  for (const paragraph of paragraphs) {{
    if (paragraph.length <= 3000) {{
      if (current && (current.length + paragraph.length + 2 > 3000)) {{
        output.push(current); current = '';
      }}
      current = current ? current + '\\n\\n' + paragraph : paragraph;
      continue;
    }}
    if (current) {{ output.push(current); current = ''; }}
    const sentences = paragraph.match(/[^.!?]+[.!?]*/g) || [paragraph];
    for (const sentenceRaw of sentences) {{
      const sentence = sentenceRaw.trim();
      if (!sentence) continue;
      if (current && current.length + sentence.length + 1 > 3000) {{
        output.push(current); current = '';
      }}
      current = current ? current + ' ' + sentence : sentence;
    }}
    if (current) {{ output.push(current); current = ''; }}
  }}
  if (current) output.push(current);
  return output.length ? output : ['No text was saved in this reader.'];
}}

function showPart() {{
  const index = Number(partSelect.value || 0);
  readerText.textContent = parts[index] || '';
  sectionHeading.textContent = 'Reading section ' + (index + 1) + ' of ' + parts.length;
  stopSpeech();
}}
function stopSpeech() {{
  if ('speechSynthesis' in window) window.speechSynthesis.cancel();
}}
parts = splitIntoParts(originalText);
parts.forEach((part, index) => {{
  const option = document.createElement('option');
  option.value = String(index);
  option.textContent = 'Section ' + (index + 1) + ' of ' + parts.length;
  partSelect.appendChild(option);
}});
partSelect.addEventListener('change', showPart);
showPart();

document.getElementById('smaller').addEventListener('click', () => {{
  fontSize = Math.max(0.85, fontSize - 0.1);
  readerText.style.fontSize = fontSize + 'rem';
}});
document.getElementById('larger').addEventListener('click', () => {{
  fontSize = Math.min(2.2, fontSize + 0.1);
  readerText.style.fontSize = fontSize + 'rem';
}});
document.getElementById('contrast').addEventListener('click', () => {{
  document.body.classList.toggle('high-contrast');
}});
document.getElementById('read').addEventListener('click', () => {{
  if (!('speechSynthesis' in window) || !('SpeechSynthesisUtterance' in window)) {{
    statusLine.textContent = 'Read-aloud is not supported by this browser.';
    return;
  }}
  stopSpeech();
  const utterance = new SpeechSynthesisUtterance(parts[Number(partSelect.value || 0)]);
  utterance.onstart = () => statusLine.textContent = 'Reading section aloud.';
  utterance.onend = () => statusLine.textContent = 'Finished reading this section.';
  window.speechSynthesis.speak(utterance);
}});
document.getElementById('pause').addEventListener('click', () => {{
  if ('speechSynthesis' in window) window.speechSynthesis.pause();
}});
document.getElementById('resume').addEventListener('click', () => {{
  if ('speechSynthesis' in window) window.speechSynthesis.resume();
}});
document.getElementById('stop').addEventListener('click', () => {{
  stopSpeech(); statusLine.textContent = 'Reading stopped.';
}});
</script>
</body>
</html>"""


# ============================================================
# 23. SPEECH TO TEXT
# ============================================================

def transcribe_audio(audio_file):
    if not SPEECH_RECOGNITION_AVAILABLE:
        raise RuntimeError("Speech recognition is not installed.")

    recognizer = sr.Recognizer()
    audio_bytes = audio_file.getvalue()

    with sr.AudioFile(io.BytesIO(audio_bytes)) as source:
        audio_data = recognizer.record(source)

    return recognizer.recognize_google(audio_data, language="en-NG")


# ============================================================
# 24. AUTHENTICATION PORTAL
# ============================================================

def render_auth_portal():
    st.markdown(
        """
        <div class="portal-banner">
            <div>
                <h1>👁️ NaijaSight Portal</h1>
                <p>Accessible Document Intelligence, Bill/Prescription Triage & Voice Suite</p>
            </div>
            <div style="display:flex; gap:8px; flex-wrap:wrap;">
                <span class="portal-tag">SQLite + Excel + PDF</span>
                <span class="portal-tag">Low-Vision Adaptive</span>
                <span class="portal-tag">Up to 500 MB</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_left, col_right = st.columns([1.15, 1], gap="large")

    with col_left:
        st.subheader("Built for Clarity, Speed & Accessibility")
        st.write(
            "NaijaSight empowers blind, low-vision, and busy users to read documents, "
            "triage utility bills and prescriptions first, scan printed pages, generate speech, "
            "and retain a searchable database record of their work."
        )
        st.markdown(
            """
            - **Read a document:** Extract text from PDF, Word, Excel, CSV, PowerPoint, EPUB, RTF, TXT, and images with key-fact triage
            - **Scan with camera:** Capture printed sheets or notices with layout-aware OCR and enhancement modes
            - **Listen to text:** Use zero-data offline browser speech, chunked MP3 audio, or downloadable offline HTML readers
            - **Voice tools:** Transcribe spoken voice recordings and store them in your account history
            - **Accessibility & AI ethics:** Document PWD user-testing observations and review privacy/consent guardrails
            """
        )

        st.divider()
        if st.button("⚡ Quick Judge / Guest Demo Access", type="primary", use_container_width=True):
            st.session_state.authenticated = True
            st.session_state.username = "demo"
            st.session_state.full_name = "Demo User"
            st.session_state.demo_mode = True
            st.session_state.nav_page = "Home"
            st.rerun()

    with col_right:
        with st.container(border=True):
            login_tab, register_tab = st.tabs(["Log in", "Create account"])

            with login_tab:
                with st.form("login_form"):
                    st.subheader("Portal Sign In")
                    username = st.text_input("Username")
                    password = st.text_input("Password", type="password")
                    submitted = st.form_submit_button("Log in", type="primary", use_container_width=True)

                    if submitted:
                        user = authenticate_user(username, password)
                        if user:
                            st.session_state.authenticated = True
                            st.session_state.username = user["username"]
                            st.session_state.full_name = user["full_name"]
                            st.session_state.demo_mode = False
                            st.session_state.nav_page = "Home"
                            st.rerun()
                        else:
                            st.error("The username or password is incorrect.")

            with register_tab:
                with st.form("register_form"):
                    st.subheader("Register Portal Account")
                    full_name = st.text_input("Full name")
                    new_username = st.text_input("Choose a username")
                    new_password = st.text_input("Create a password (min 8 chars)", type="password")
                    confirm_password = st.text_input("Confirm password", type="password")
                    submitted = st.form_submit_button("Create account", use_container_width=True)

                    if submitted:
                        if new_password != confirm_password:
                            st.error("Your passwords do not match.")
                        else:
                            success, message = create_account(full_name, new_username, new_password)
                            if success:
                                st.success(message)
                                st.info("You can now log in.")
                            else:
                                st.error(message)


# ============================================================
# 25. SIDEBAR, LOW-VISION SETTINGS & PORTAL HEADER
# ============================================================

def render_sidebar():
    with st.sidebar:
        st.title("👁️ NaijaSight")
        st.caption(f"Portal Navigation · v{APP_VERSION}")
        st.divider()

        st.write(f"Signed in as **{st.session_state.full_name}**")
        if st.session_state.demo_mode:
            st.info("Demonstration mode (unsaved session)")

        pages = [
            "Home",
            "Read a document",
            "Scan with camera",
            "Listen to text",
            "Voice tools",
            "Activity history",
            "Accessibility & AI ethics"
        ]
        page = st.radio("Select Module", pages, key="nav_page")

        st.divider()
        st.markdown("**Low-Vision Display**")
        st.selectbox(
            "Contrast theme",
            ["Standard", "High Contrast (Print Ledger)", "Low-Glare Night (Warm Charcoal)"],
            key="vision_theme"
        )
        st.slider(
            "Editor text size (px)",
            min_value=16,
            max_value=28,
            step=2,
            key="reader_font_px"
        )

        st.divider()
        st.caption(f"Max upload size: {MAX_FILE_SIZE_MB} MB")

        if st.button("Log out", use_container_width=True):
            logout()

    return page


def render_portal_header(active_page):
    records = get_activity_records() if not st.session_state.demo_mode else []
    total_words = sum(r["word_count"] for r in records)

    st.markdown(
        f"""
        <div class="portal-banner">
            <div>
                <h2>👁️ NaijaSight Portal — {html.escape(active_page)}</h2>
                <p>Welcome, <strong>{html.escape(st.session_state.full_name or '')}</strong> (@{html.escape(st.session_state.username or '')})</p>
            </div>
            <div style="display:flex; gap:8px; flex-wrap:wrap;">
                <span class="portal-tag">Saved Records: {len(records)}</span>
                <span class="portal-tag">Words Processed: {total_words:,}</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# 26. HOME PAGE
# ============================================================

def go_to_page(page):
    st.session_state["nav_page"] = page


def render_home():
    st.subheader("Portal Category Hub")
    st.caption("Select a category below to open its dedicated workspace.")

    cards = [
        ("📄", "Read a document", "Extract text from PDF, Word, Excel, CSV, PowerPoint, EPUB, RTF, TXT, and images with key-fact triage.", "Read a document"),
        ("📷", "Scan with camera", "Capture printed pages, bills, or notices with your camera using Standard or Enhanced OCR.", "Scan with camera"),
        ("🔊", "Listen to text", "Read extracted or pasted text aloud offline or in clean MP3 parts and download an offline reader.", "Listen to text"),
        ("🎙️", "Voice tools", "Record spoken instructions or notes and transcribe them into editable text.", "Voice tools"),
    ]
    columns = st.columns(2, gap="large")
    for index, (symbol, title, description, destination) in enumerate(cards):
        with columns[index % 2]:
            with st.container(border=True):
                st.markdown(f"### {symbol}  {title}")
                st.write(description)
                st.button(
                    "Open " + title,
                    key="home_" + destination,
                    use_container_width=True,
                    type="primary",
                    on_click=go_to_page,
                    args=(destination,)
                )

    st.divider()
    st.subheader("Quick Workflow")
    st.write("1. Choose a category.  2. Upload, capture, or load a sample.  3. Check key facts & edit text.  4. Listen, export Excel/PDF/JSON, or open offline.")
    st.info("For clearest recognition, keep text upright and in focus. Always verify important names, dates, dosages, and figures.")


# ============================================================
# 27. RESULT WORKSPACE
# ============================================================

def render_result_workspace():
    text = st.session_state.extracted_text

    if not text:
        st.info("Your extracted text will appear here after processing.")
        return

    st.subheader("Extracted Text & Triage Workspace")
    st.caption("OCR confidence is an estimate from detected text regions, not a guarantee that every word is correct.")

    elapsed_ms = st.session_state.get("last_process_ms")
    confidence = st.session_state.get("last_ocr_confidence")
    if confidence is not None:
        m1, m2 = st.columns(2)
        m1.metric("OCR confidence", f"{confidence}%")
        if elapsed_ms is not None:
            m2.metric("Processing time", f"{elapsed_ms} ms")

        if confidence < 60:
            st.warning("The OCR engine is uncertain about some text. Please review carefully.")
        elif confidence < 80:
            st.info("Some parts of the text may need checking.")
        else:
            st.success("The OCR engine recognised the text with high confidence.")

        uncertain = st.session_state.get("last_uncertain_words", [])
        if uncertain:
            with st.expander("Review uncertain words"):
                for item in uncertain:
                    st.write(f"**{item['text']}** — {item['confidence']}% confidence")
    elif elapsed_ms is not None:
        st.caption(f"{len(text.split())} words · {len(text)} characters · Processed in {elapsed_ms} ms")
    else:
        st.caption(f"{len(text.split())} words · {len(text)} characters")

    if confidence is not None:
        st.caption(f"{len(text.split())} words · {len(text)} characters")

    if st.session_state.get("editable_result") != text:
        st.session_state.editable_result = text

    edited_text = st.text_area(
        "Review or edit the extracted text",
        height=280,
        key="editable_result"
    )
    st.session_state.extracted_text = edited_text

    render_triage_panel(edited_text, "workspace_triage")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "Download text (.TXT)",
            data=edited_text,
            file_name="naijasight_extracted_text.txt",
            mime="text/plain",
            use_container_width=True
        )

    with col2:
        offline_html = build_offline_reader_html(
            edited_text,
            st.session_state.get("current_filename") or "NaijaSight Offline Reader"
        )
        st.download_button(
            "Offline reader (.HTML)",
            data=offline_html.encode("utf-8"),
            file_name="naijasight_offline_reader.html",
            mime="text/html",
            use_container_width=True
        )

    with col3:
        if st.button("Clear result", use_container_width=True):
            clear_current_result()
            st.rerun()

    st.divider()
    st.subheader("Listen to the text")
    render_speech_controls(edited_text, "result")


# ============================================================
# 28. DOCUMENT READER PAGE
# ============================================================

def render_document_reader():
    st.subheader("Read a Document or Image")
    st.write("Upload a document or image file, or load a 1-click Nigerian sample bill/prescription for testing.")

    with st.expander("🧪 1-Click Sample Nigerian Documents (For Instant Testing & Demo)"):
        sample_choice = st.selectbox("Choose a sample document", list(SAMPLE_DOCUMENTS.keys()))
        if st.button("Load Sample into Workspace", key="load_sample_doc"):
            t0 = time.perf_counter()
            sample_text = SAMPLE_DOCUMENTS[sample_choice]
            st.session_state.extracted_text = sample_text
            st.session_state.editable_result = sample_text
            st.session_state.current_filename = f"{sample_choice}.txt"
            st.session_state.current_file_type = ".txt"
            st.session_state.last_ocr_confidence = None
            st.session_state.last_uncertain_words = []
            st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)
            log_activity(f"{sample_choice}.txt", ".txt", sample_text)
            st.success("Sample document loaded and logged.")
            st.rerun()

    uploaded_file = st.file_uploader(
        "Choose a file",
        type=[
            "png", "jpg", "jpeg", "webp", "bmp",
            "tif", "tiff", "pdf", "docx", "pptx",
            "xlsx", "csv", "txt", "rtf", "epub"
        ],
        max_upload_size=MAX_FILE_SIZE_MB
    )

    image_extensions = {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}
    image_mode = "Read Entire Image"
    quality_mode = "Standard"

    if uploaded_file:
        st.write(f"Selected file: **{uploaded_file.name}**")
        selected_extension = Path(uploaded_file.name).suffix.lower()

        if selected_extension in image_extensions:
            image_mode = st.radio(
                "Reading scope",
                ["Read Entire Image", "Read Main Document"],
                horizontal=True
            )

        if selected_extension in image_extensions or selected_extension == ".pdf":
            quality_mode = st.radio(
                "OCR image preparation",
                ["Standard", "Enhanced"],
                horizontal=True
            )

        st.caption(f"File size: {uploaded_file.size / (1024 * 1024):.2f} MB")

        if st.button("Extract text", type="primary"):
            clear_current_result()
            try:
                t0 = time.perf_counter()
                with st.spinner("Processing your document..."):
                    extracted_text = process_uploaded_document(
                        uploaded_file,
                        image_mode=image_mode,
                        quality_mode=quality_mode
                    )
                st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)

                if extracted_text:
                    st.session_state.extracted_text = extracted_text
                    st.session_state.current_filename = uploaded_file.name
                    st.session_state.current_file_type = selected_extension
                    st.session_state.editable_result = extracted_text

                    log_activity(uploaded_file.name, selected_extension, extracted_text)
                    st.success("Text extraction completed and saved.")
                else:
                    st.warning("No readable text was found in this file.")

            except Exception as error:
                st.error(f"Could not read the file: {error}")

    render_result_workspace()


# ============================================================
# 29. CAMERA SCANNER
# ============================================================

def render_camera_scanner():
    st.subheader("Scan with Camera")
    st.write("Take a clear picture of a printed page, textbook, utility bill, prescription, or notice.")

    camera_mode = st.radio(
        "Reading scope",
        ["Read Entire Image", "Read Main Document"],
        horizontal=True
    )
    camera_quality = st.radio(
        "OCR image preparation",
        ["Standard", "Enhanced"],
        horizontal=True
    )
    camera_image = st.camera_input("Take a picture")

    if camera_image:
        if st.button("Read this picture", type="primary"):
            clear_current_result()
            try:
                t0 = time.perf_counter()
                with st.spinner("Reading camera image..."):
                    image = Image.open(camera_image)
                    extracted_text = image_to_text(
                        image,
                        mode=camera_mode,
                        quality_mode=camera_quality
                    )
                st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)

                if extracted_text:
                    st.session_state.extracted_text = extracted_text
                    st.session_state.current_filename = "Camera scan"
                    st.session_state.current_file_type = "camera"
                    st.session_state.editable_result = extracted_text

                    log_activity("Camera scan", "camera", extracted_text)
                    st.success("Text extracted and saved.")
                else:
                    st.warning("No text was detected. Try taking another picture.")

            except Exception as error:
                st.error(f"Could not process the picture: {error}")

    render_result_workspace()


# ============================================================
# 30. LISTEN TO TEXT
# ============================================================

def render_listen_page():
    st.subheader("Listen to Text")
    st.write("Paste any text, load a sample, or use extracted text to generate speech in manageable parts.")

    if st.button("Load Sample Bill for Audio Test", key="listen_sample_btn"):
        st.session_state["standalone_listen_text"] = SAMPLE_DOCUMENTS["⚡ EKEDC Electricity Utility Bill"]
        st.rerun()

    initial_text = st.session_state.get("extracted_text", "")
    text_to_read = st.text_area(
        "Text to read aloud",
        value=initial_text,
        height=260,
        key="standalone_listen_text"
    )
    if text_to_read.strip():
        render_triage_panel(text_to_read, "standalone_triage")
    render_speech_controls(text_to_read, "standalone")


# ============================================================
# 31. VOICE TOOLS
# ============================================================

def render_voice_tools():
    st.subheader("Voice Transcription Tools")
    st.write("Record a short spoken message and convert it into text.")

    if not SPEECH_RECOGNITION_AVAILABLE:
        st.error("The speech recognition package is not installed.")
        return

    st.warning(
        "Voice recognition sends your recording to Google's speech recognition service. "
        "Avoid recording sensitive or private information."
    )
    consent = st.checkbox("I agree to send my recording to Google for transcription.")
    audio_file = st.audio_input("Record your voice")

    if audio_file:
        st.audio(audio_file)
        if st.button("Convert speech to text", disabled=not consent, type="primary"):
            try:
                t0 = time.perf_counter()
                with st.spinner("Transcribing recording..."):
                    transcript = transcribe_audio(audio_file)
                st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)

                st.session_state.extracted_text = transcript
                st.session_state.current_filename = "Voice recording"
                st.session_state.current_file_type = "audio"
                st.session_state.editable_result = transcript

                log_activity("Voice recording", "audio", transcript)
                st.success("Transcription completed and saved.")
            except Exception as error:
                st.error(f"Could not transcribe recording: {error}")

    render_result_workspace()


# ============================================================
# 32. ACTIVITY HISTORY, SEARCH, EXPORTS & RETENTION
# ============================================================

def render_history():
    st.subheader("Activity History, Stored Records & Exports")

    if st.session_state.demo_mode:
        st.info("Sign in with a registered account to store and retain history across sessions.")
        return

    records = get_activity_records()
    if not records:
        st.info("You have no saved activity records yet.")
        return

    total_words = sum(r["word_count"] for r in records)
    total_chars = sum(r["character_count"] for r in records)

    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Saved Extractions</div><div class="kpi-value">{len(records)}</div></div>',
            unsafe_allow_html=True
        )
    with k2:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Total Words</div><div class="kpi-value">{total_words:,}</div></div>',
            unsafe_allow_html=True
        )
    with k3:
        st.markdown(
            f'<div class="kpi-card"><div class="kpi-label">Total Characters</div><div class="kpi-value">{total_chars:,}</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("#### Search & Filter Retained Records")
    f_col1, f_col2 = st.columns([2, 1])
    with f_col1:
        search_query = st.text_input(
            "Search by filename or extracted text",
            placeholder="e.g., bill, prescription, camera, chapter 1..."
        ).strip().lower()

    available_types = sorted({(r["file_type"] or "unknown") for r in records})
    with f_col2:
        type_filter = st.selectbox(
            "Filter by type",
            options=["All Types"] + available_types
        )

    filtered_records = []
    for r in records:
        matches_type = (type_filter == "All Types") or ((r["file_type"] or "unknown") == type_filter)
        haystack = f"{r['filename']} {r['file_type']} {r.get('extracted_text', '')}".lower()
        matches_query = (not search_query) or (search_query in haystack)
        if matches_type and matches_query:
            filtered_records.append(r)

    st.caption(f"Showing {len(filtered_records)} of {len(records)} saved record(s).")

    if filtered_records:
        filtered_df = pd.DataFrame([
            {
                "ID": r["id"],
                "Filename": r["filename"],
                "Type": r["file_type"],
                "Words": r["word_count"],
                "Characters": r["character_count"],
                "Status": r["status"],
                "Timestamp": r["created_at"],
            }
            for r in filtered_records
        ])
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

        exp_col1, exp_col2, exp_col3 = st.columns(3)
        with exp_col1:
            st.download_button(
                "📊 Export Excel (.XLSX)",
                data=build_excel_export_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with exp_col2:
            st.download_button(
                "📄 Export Report (.PDF)",
                data=build_pdf_report_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with exp_col3:
            st.download_button(
                "🧾 Audit Snapshot (.JSON)",
                data=build_audit_json_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_audit.json",
                mime="application/json",
                use_container_width=True
            )
    else:
        st.warning("No records match your current search or filter.")

    st.divider()
    st.subheader("Reload or Manage a Saved Record")

    pool_for_selection = filtered_records if filtered_records else records
    record_map = {
        f"#{r['id']} — {r['filename']} ({r['word_count']} words, {r['created_at']})": r
        for r in pool_for_selection
    }
    selected_label = st.selectbox("Choose a saved record", list(record_map.keys()))
    chosen_record = record_map[selected_label]

    if chosen_record.get("extracted_text"):
        with st.expander("Preview retained text"):
            st.write(chosen_record["extracted_text"])

    col_reload, col_del_one = st.columns(2)
    with col_reload:
        if st.button("📂 Load Text into Workspace", use_container_width=True):
            st.session_state.extracted_text = chosen_record.get("extracted_text", "")
            st.session_state.editable_result = chosen_record.get("extracted_text", "")
            st.session_state.current_filename = chosen_record["filename"]
            st.session_state.current_file_type = chosen_record["file_type"]
            st.session_state.nav_page = "Read a document"
            st.rerun()

    with col_del_one:
        if st.button("🗑️ Delete This Record", use_container_width=True):
            delete_activity_item(chosen_record["id"])
            st.success("Record deleted.")
            st.rerun()

    st.divider()
    st.subheader("Clear All Activity History")
    st.warning("This permanently deletes all activity records stored under your account.")
    confirm_delete = st.checkbox("I understand that this cannot be undone.")
    if st.button("Delete all my history", disabled=not confirm_delete):
        delete_activity()
        st.success("Your activity history has been deleted.")
        st.rerun()


# ============================================================
# 32B. ACCESSIBILITY TESTING LOG & RESPONSIBLE AI NOTE
# ============================================================

def render_accessibility_and_ethics():
    st.subheader("Accessibility Testing Log & Responsible AI Note")
    st.write(
        "Document real user-testing observations (especially with blind and low-vision testers), "
        "track product fixes, and maintain transparent privacy and reliability guardrails."
    )

    st.markdown(
        """
        - **Informed Consent & Data Minimization:** Speech synthesis and voice transcription only run after explicit user checkbox consent.
        - **Low-Vision & Screen-Reader Support:** High-contrast and low-glare themes, adjustable text sizes, keyboard focus outlines, zero-data offline browser speech, and standalone offline HTML readers.
        - **Accuracy & Safety Guardrails:** OCR confidence estimates and uncertain-word flags highlight potential errors on financial amounts, dates, and medication dosages.
        """
    )

    st.divider()
    st.markdown("#### Log an Accessibility Testing Observation")
    with st.form("a11y_feedback_form"):
        tester_group = st.text_input(
            "Tester Profile / Context",
            placeholder="e.g., Low-vision student on mobile / Blind tester with screen reader"
        )
        severity = st.selectbox("Priority / Severity", ["High", "Medium", "Low"])
        issue_found = st.text_area("What broke, slowed down, or confused the tester?", height=90)
        fix_applied = st.text_area("What change did the team make in response?", height=90)
        submitted = st.form_submit_button("Save Testing Note to SQLite", type="primary")

        if submitted:
            if not tester_group.strip() or not issue_found.strip() or not fix_applied.strip():
                st.warning("Please complete tester profile, issue observed, and fix implemented.")
            else:
                log_accessibility_feedback(tester_group, issue_found, fix_applied, severity)
                st.success("Accessibility testing note saved to database (and included in Excel/JSON exports).")
                st.rerun()

    feedback_df = get_accessibility_feedback_df()
    st.markdown("#### Recorded Accessibility Testing Notes")
    if feedback_df.empty:
        st.info("No testing notes logged yet. Add your first observation above.")
    else:
        st.dataframe(feedback_df, use_container_width=True, hide_index=True)


# ============================================================
# 33. LOGGED-IN PORTAL ROUTER
# ============================================================

def render_logged_in_portal():
    page = render_sidebar()
    apply_portal_styles(
        vision_theme=st.session_state.get("vision_theme", "Standard"),
        reader_font_px=st.session_state.get("reader_font_px", 18)
    )
    render_portal_header(page)

    if page == "Home":
        render_home()
    elif page == "Read a document":
        render_document_reader()
    elif page == "Scan with camera":
        render_camera_scanner()
    elif page == "Listen to text":
        render_listen_page()
    elif page == "Voice tools":
        render_voice_tools()
    elif page == "Activity history":
        render_history()
    elif page == "Accessibility & AI ethics":
        render_accessibility_and_ethics()

    st.divider()
    st.caption(
        "NaijaSight Portal | Assistive reading, triage & voice suite. "
        "Always verify important names, dates, dosages, and figures."
    )


# ============================================================
# 34. APPLICATION ENTRY POINT
# ============================================================

if not st.session_state.authenticated:
    apply_portal_styles(
        vision_theme=st.session_state.get("vision_theme", "Standard"),
        reader_font_px=st.session_state.get("reader_font_px", 18)
    )
    render_auth_portal()
else:
    render_logged_in_portal()