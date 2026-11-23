import hashlib
import html
import io
import json
import os
import re
import secrets
import shutil
import sqlite3
import subprocess
import tempfile
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

try:
    import imageio_ffmpeg
    IMAGEIO_FFMPEG_AVAILABLE = True
except ImportError:
    IMAGEIO_FFMPEG_AVAILABLE = False


# ============================================================
# 1. APPLICATION SETTINGS
# ============================================================

st.set_page_config(
    page_title="NaijaSight Portal | Assistive Document, Naira & Voice Suite",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_NAME = "NaijaSight"
APP_VERSION = "2.7-DSN"
BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "naijasight_portal.db"

MAX_FILE_SIZE_MB = 500
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

MAX_IMAGE_PIXELS = 25_000_000
MAX_PDF_PAGES = 100

Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS


# ============================================================
# 2. CLEAN PRODUCT STYLING, SCOPED CONTROLS & LOW-VISION MODES
# ============================================================

def apply_portal_styles(vision_theme: str = "Standard (Crisp Light)", reader_font_px: int = 18):
    if vision_theme == "High Contrast (Pure White / Black)":
        bg_color = "#FFFFFF"
        sidebar_bg = "#FFFFFF"
        card_bg = "#FFFFFF"
        text_color = "#000000"
        muted_color = "#1F2937"
        border_color = "#000000"
        tag_bg = "#F3F4F6"
        tag_text = "#000000"
        tag_border = "#000000"
        accent_color = "#0F172A"
        accent_hover = "#000000"
        focus_color = "#1D4ED8"
        input_bg = "#FFFFFF"
        secondary_btn_bg = "#FFFFFF"
        secondary_btn_text = "#000000"
    elif vision_theme == "Low-Glare Slate":
        bg_color = "#111317"
        sidebar_bg = "#16191F"
        card_bg = "#1B1F27"
        text_color = "#F3F4F6"
        muted_color = "#9CA3AF"
        border_color = "#2D3340"
        tag_bg = "#232834"
        tag_text = "#E5E7EB"
        tag_border = "#374151"
        accent_color = "#2563EB"
        accent_hover = "#1D4ED8"
        focus_color = "#60A5FA"
        input_bg = "#14171D"
        secondary_btn_bg = "#232834"
        secondary_btn_text = "#F3F4F6"
    else:
        bg_color = "#F9FAFB"
        sidebar_bg = "#FFFFFF"
        card_bg = "#FFFFFF"
        text_color = "#0F172A"
        muted_color = "#4B5563"
        border_color = "#E5E7EB"
        tag_bg = "#F3F4F6"
        tag_text = "#1F2937"
        tag_border = "#D1D5DB"
        accent_color = "#1D4ED8"
        accent_hover = "#1E40AF"
        focus_color = "#1D4ED8"
        input_bg = "#FFFFFF"
        secondary_btn_bg = "#F3F4F6"
        secondary_btn_text = "#111827"

    st.markdown(
        f"""
        <style>
        :root {{
            --ns-bg: {bg_color};
            --ns-sidebar: {sidebar_bg};
            --ns-card: {card_bg};
            --ns-text: {text_color};
            --ns-muted: {muted_color};
            --ns-border: {border_color};
            --ns-tag-bg: {tag_bg};
            --ns-tag-text: {tag_text};
            --ns-tag-border: {tag_border};
            --ns-accent: {accent_color};
            --ns-accent-hover: {accent_hover};
            --ns-focus: {focus_color};
            --ns-input-bg: {input_bg};
            --ns-sec-btn-bg: {secondary_btn_bg};
            --ns-sec-btn-text: {secondary_btn_text};
        }}

        [data-testid="stAppViewContainer"] {{
            background-color: var(--ns-bg) !important;
            color: var(--ns-text) !important;
        }}

        [data-testid="stHeader"] {{
            background-color: transparent !important;
        }}

        [data-testid="stSidebar"] {{
            background-color: var(--ns-sidebar) !important;
            color: var(--ns-text) !important;
            border-right: 1px solid var(--ns-border) !important;
        }}

        [data-testid="stSidebar"] * {{
            color: var(--ns-text);
        }}

        [data-testid="stSidebar"] .stCaption,
        [data-testid="stSidebar"] small {{
            color: var(--ns-muted) !important;
        }}

        [data-testid="stAppViewContainer"] .stMarkdown,
        [data-testid="stAppViewContainer"] .stMarkdown p,
        [data-testid="stAppViewContainer"] .stMarkdown li,
        [data-testid="stAppViewContainer"] label {{
            color: var(--ns-text);
        }}

        .block-container {{
            max-width: 1160px;
            padding-top: 1.5rem;
            padding-bottom: 3.5rem;
        }}

        .ns-header-bar {{
            background: var(--ns-card);
            border: 1px solid var(--ns-border);
            border-radius: 12px;
            padding: 1.15rem 1.4rem;
            margin-bottom: 1.25rem;
            display: flex;
            flex-wrap: wrap;
            justify-content: space-between;
            align-items: center;
            gap: 0.9rem;
        }}

        .ns-header-bar h1,
        .ns-header-bar h2,
        .ns-header-bar p,
        .ns-header-bar strong {{
            color: var(--ns-text) !important;
            margin: 0 !important;
        }}

        .ns-header-bar p {{
            color: var(--ns-muted) !important;
            margin-top: 0.2rem !important;
            font-size: 0.92rem;
        }}

        .ns-chip-row {{
            display: flex;
            gap: 8px;
            flex-wrap: wrap;
        }}

        .ns-chip {{
            background: var(--ns-tag-bg);
            border: 1px solid var(--ns-tag-border);
            color: var(--ns-tag-text);
            padding: 0.3rem 0.7rem;
            border-radius: 999px;
            font-size: 0.78rem;
            font-weight: 600;
        }}

        .ns-kpi-card {{
            background: var(--ns-card);
            border: 1px solid var(--ns-border);
            border-radius: 12px;
            padding: 0.95rem 1.1rem;
            margin-bottom: 0.5rem;
        }}

        .ns-kpi-label {{
            font-size: 0.74rem;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.05em;
            color: var(--ns-muted);
        }}

        .ns-kpi-value {{
            font-size: 1.4rem;
            font-weight: 700;
            color: var(--ns-text);
            margin-top: 0.2rem;
        }}

        .ns-triage-card {{
            background: var(--ns-card);
            border: 1px solid var(--ns-border);
            border-left: 4px solid var(--ns-accent);
            border-radius: 12px;
            padding: 1rem 1.2rem;
            margin: 0.85rem 0 1rem 0;
            color: var(--ns-text);
        }}

        div.stButton > button[kind="primary"] {{
            background-color: var(--ns-accent) !important;
            border: 1px solid var(--ns-accent) !important;
            color: #FFFFFF !important;
            border-radius: 8px;
            min-height: 42px;
            font-weight: 600;
        }}

        div.stButton > button[kind="primary"]:hover {{
            background-color: var(--ns-accent-hover) !important;
            border-color: var(--ns-accent-hover) !important;
            color: #FFFFFF !important;
        }}

        div.stButton > button[kind="secondary"],
        div.stDownloadButton > button {{
            background-color: var(--ns-sec-btn-bg) !important;
            border: 1px solid var(--ns-border) !important;
            color: var(--ns-sec-btn-text) !important;
            border-radius: 8px;
            min-height: 42px;
            font-weight: 600;
        }}

        [data-testid="stTextArea"] textarea {{
            font-size: {reader_font_px}px !important;
            line-height: 1.65 !important;
            background-color: var(--ns-input-bg) !important;
            color: var(--ns-text) !important;
            border: 1px solid var(--ns-border) !important;
            border-radius: 10px !important;
        }}

        [data-testid="stTextInput"] input,
        [data-testid="stNumberInput"] input {{
            background-color: var(--ns-input-bg) !important;
            color: var(--ns-text) !important;
            border: 1px solid var(--ns-border) !important;
            border-radius: 8px !important;
        }}

        button:focus-visible,
        input:focus-visible,
        textarea:focus-visible,
        select:focus-visible {{
            outline: 2px solid var(--ns-focus) !important;
            outline-offset: 2px !important;
        }}

        @media (max-width: 768px) {{
            .block-container {{
                padding-left: 0.85rem;
                padding-right: 0.85rem;
                padding-top: 1rem;
            }}

            .ns-header-bar {{
                padding: 1rem;
                flex-direction: column;
                align-items: flex-start;
            }}

            div.stButton > button,
            div.stDownloadButton > button {{
                min-height: 48px;
                width: 100%;
                white-space: normal;
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

    pdf.set_fill_color(29, 78, 216)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 15)
    pdf.cell(0, 12, _sanitize_pdf_text("NaijaSight Portal - Extraction & Activity Report"), ln=True, fill=True)

    pdf.set_text_color(31, 41, 55)
    pdf.set_font("Helvetica", "", 10)
    pdf.ln(3)
    user_label = f"{st.session_state.get('full_name', '')} (@{st.session_state.get('username', '')})"
    generated_stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    pdf.cell(0, 6, _sanitize_pdf_text(f"Account: {user_label}   |   Generated: {generated_stamp}"), ln=True)
    pdf.cell(0, 6, _sanitize_pdf_text(f"Records Included: {len(records)}"), ln=True)
    pdf.ln(4)

    for idx, rec in enumerate(records, start=1):
        pdf.set_draw_color(209, 213, 219)
        pdf.set_fill_color(243, 244, 246)
        pdf.set_font("Helvetica", "B", 11)
        pdf.set_text_color(15, 23, 42)
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
        pdf.set_text_color(17, 24, 39)
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
    "last_image_color_hint": None,

    "editable_result": "",
    "vision_theme": "Standard (Crisp Light)",
    "reader_font_px": 18,
    "auto_speak_triage": False,
    "target_payment_ngn": 0.0,
    "cash_basket_notes": [],
    "last_voice_amount_transcript": "",
    "rendered_video_bytes": None,
    "rendered_video_name": "",
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
    st.session_state.last_image_color_hint = None
    st.session_state.editable_result = ""


def logout():
    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.full_name = None
    st.session_state.demo_mode = False
    st.session_state.nav_page = "Home"
    st.session_state.cash_basket_notes = []
    st.session_state.rendered_video_bytes = None
    st.session_state.rendered_video_name = ""
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
# 10. IMAGE ENHANCEMENT, COLOR HINT & LAYOUT GROUPING
# ============================================================

def prepare_image(image):
    image = ImageOps.exif_transpose(image)
    if image.width * image.height > MAX_IMAGE_PIXELS:
        raise ValueError("Image is too large to process safely.")
    return image.convert("RGB")


def estimate_naira_color_hint(image: Image.Image) -> str | None:
    thumb = image.resize((64, 64))
    arr = np.asarray(thumb, dtype=np.float32)
    center = arr[16:48, 16:48]
    r, g, b = center[:, :, 0].mean(), center[:, :, 1].mean(), center[:, :, 2].mean()
    sat = max(r, g, b) - min(r, g, b)
    if sat < 14:
        return None
    if b > r + 18 and b > g + 10:
        return "Blue/Cyan tint (consistent with ₦1,000 or ₦100 note)"
    if r > g + 22 and r > b + 22:
        return "Red/Terracotta tint (consistent with ₦500 or ₦10 note)"
    if g > r + 15 and g > b + 12:
        return "Green tint (consistent with ₦200 or ₦20 note)"
    return None


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
    color_hint = estimate_naira_color_hint(prepared)

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
        return {"text": "", "confidence": 0, "uncertain_words": [], "color_hint": color_hint}

    confidences = [item["confidence"] for item in items]
    uncertain_words = [
        {"text": item["text"], "confidence": round(item["confidence"] * 100, 1)}
        for item in items if item["confidence"] < 0.60
    ]
    return {
        "text": "\n".join(lines),
        "confidence": round(sum(confidences) / len(confidences) * 100, 1),
        "uncertain_words": uncertain_words,
        "color_hint": color_hint,
    }


def image_to_text(image, mode="Read Entire Image", quality_mode="Standard"):
    result = recognize_image(image, mode=mode, quality_mode=quality_mode)
    st.session_state.last_ocr_confidence = result["confidence"]
    st.session_state.last_uncertain_words = result["uncertain_words"]
    st.session_state.last_image_color_hint = result.get("color_hint")
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
# 21. DOCUMENT ROUTER, NAIRA CASH BASKET, VOICE AMOUNT & TRIAGE
# ============================================================

SAMPLE_DOCUMENTS = {
    "💵 Single ₦1,000 Note (Snap & Add to Basket)": (
        "CENTRAL BANK OF NIGERIA\n"
        "ONE THOUSAND NAIRA - 1000\n"
        "Legal tender for payment of any amount."
    ),
    "💵 Single ₦500 Note (Snap & Add to Basket)": (
        "CENTRAL BANK OF NIGERIA\n"
        "FIVE HUNDRED NAIRA - 500\n"
        "Legal tender for payment of any amount."
    ),
    "📱 Mobile Transfer / POS Screen Receipt": (
        "OPay Digital Services Limited - Transfer Receipt\n"
        "Transaction Status: Successful\n"
        "Amount: NGN 3,500.00\n"
        "Recipient: Mama Nkechi Provision Store\n"
        "Bank: Moniepoint MFB | Account: 8012349876\n"
        "Session ID: 09040526112201449281\n"
        "Date: 22/11/2026 14:12"
    ),
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
}


def parse_spoken_amount_ngn(transcript: str) -> float | None:
    if not transcript:
        return None
    s = transcript.lower().replace(",", " ").replace("₦", " ").replace("ngn", " ")

    num_matches = re.findall(r"\b\d+(?:\.\d{1,2})?\b", s)
    if num_matches:
        try:
            val = float(num_matches[0])
            if val > 0:
                return val
        except ValueError:
            pass

    units = {
        "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5,
        "six": 6, "seven": 7, "eight": 8, "nine": 9, "ten": 10,
        "eleven": 11, "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
        "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19,
    }
    tens = {
        "twenty": 20, "thirty": 30, "forty": 40, "fifty": 50,
        "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
    }
    tokens = re.findall(r"[a-z]+", s)
    total = 0
    current = 0
    found_any = False

    for tok in tokens:
        if tok in {"naira", "kobo", "and"}:
            continue
        if tok in units:
            current += units[tok]
            found_any = True
        elif tok in tens:
            current += tens[tok]
            found_any = True
        elif tok == "hundred":
            current = (current if current > 0 else 1) * 100
            found_any = True
        elif tok == "thousand":
            total += (current if current > 0 else 1) * 1000
            current = 0
            found_any = True
        elif tok == "million":
            total += (current if current > 0 else 1) * 1_000_000
            current = 0
            found_any = True

    total += current
    return float(total) if (found_any and total > 0) else None


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


def detect_naira_banknotes(text: str) -> dict:
    clean = (text or "").strip()
    if not clean:
        return {"notes": [], "total_ngn": 0, "cbn_verified": False}

    low = clean.lower()
    cbn_verified = any(
        phrase in low for phrase in ("central bank of nigeria", "legal tender", "naira")
    )

    denom_patterns = [
        (1000, r"(?:one\s+thousand\s+naira|(?:₦|ngn)?\s?\b1,?000\b)"),
        (500, r"(?:five\s+hundred\s+naira|(?:₦|ngn)?\s?\b500\b)"),
        (200, r"(?:two\s+hundred\s+naira|(?:₦|ngn)?\s?\b200\b)"),
        (100, r"(?:one\s+hundred\s+naira|(?:₦|ngn)?\s?\b100\b)"),
        (50, r"(?:fifty\s+naira|(?:₦|ngn)?\s?\b50\b)"),
        (20, r"(?:twenty\s+naira|(?:₦|ngn)?\s?\b20\b)"),
        (10, r"(?:ten\s+naira|(?:₦|ngn)?\s?\b10\b)"),
        (5, r"(?:five\s+naira|(?:₦|ngn)?\s?\b5\b)"),
    ]

    is_currency_context = cbn_verified or bool(re.search(r"(?:₦|ngn|naira)", low))
    if not is_currency_context:
        return {"notes": [], "total_ngn": 0, "cbn_verified": False}

    detected_notes = []
    for line in clean.splitlines():
        line_low = line.lower()
        for value, pat in denom_patterns:
            if re.search(pat, line_low):
                detected_notes.append(value)
                break

    total_ngn = sum(detected_notes)
    return {
        "notes": detected_notes,
        "total_ngn": total_ngn,
        "cbn_verified": cbn_verified,
    }


def analyze_screen_or_receipt(text: str) -> dict:
    clean = (text or "").strip()
    if not clean:
        return {}

    low = clean.lower()
    screen_markers = (
        "successful", "transaction", "transfer", "recipient", "beneficiary",
        "session id", "opay", "palmpay", "moniepoint", "gtbank", "access bank",
        "firstbank", "zenith", "uba", "kuda", "pos", "ussd", "debited", "credited"
    )
    if not any(m in low for m in screen_markers):
        return {}

    status = "Unknown"
    if "successful" in low or "approved" in low or "completed" in low:
        status = "Successful"
    elif "failed" in low or "declined" in low:
        status = "Failed / Declined"
    elif "pending" in low or "processing" in low:
        status = "Pending"

    recipient_match = re.search(
        r"(?:Recipient|Beneficiary|Merchant|Account Name|To)\s*[:\-]\s*([^\n|]+)",
        clean,
        flags=re.IGNORECASE
    )
    bank_match = re.search(
        r"(?:Bank|Institution)\s*[:\-]\s*([^\n|]+)",
        clean,
        flags=re.IGNORECASE
    )
    amount_match = re.search(
        r"(?:Amount|Total|Paid|Sent|Received)?\s*[:\-]?\s*((?:NGN|₦)\s?\d[\d,]*(?:\.\d{1,2})?)",
        clean,
        flags=re.IGNORECASE
    )

    recipient = recipient_match.group(1).strip() if recipient_match else None
    bank = bank_match.group(1).strip() if bank_match else None
    amount = amount_match.group(1).strip() if amount_match else None

    summary_bits = [f"Screen receipt status: {status}."]
    if amount:
        summary_bits.append(f"Amount: {amount}.")
    if recipient:
        summary_bits.append(f"Recipient: {recipient}.")
    if bank:
        summary_bits.append(f"Bank: {bank}.")

    return {
        "status": status,
        "amount": amount,
        "recipient": recipient,
        "bank": bank,
        "spoken_receipt": " ".join(summary_bits),
    }


def format_payment_comparison(detected_numeric: float, target_payment_ngn: float, label_prefix: str = "Detected") -> str | None:
    if not target_payment_ngn or target_payment_ngn <= 0 or detected_numeric <= 0:
        return None
    diff = round(detected_numeric - float(target_payment_ngn), 2)
    if abs(diff) < 0.01:
        return f"Exact payment match: ₦{detected_numeric:,.2f} matches target ₦{target_payment_ngn:,.2f}."
    if diff > 0:
        return (
            f"Over target by ₦{diff:,.2f} ({label_prefix} ₦{detected_numeric:,.2f} vs Target ₦{target_payment_ngn:,.2f}). "
            f"Expected change: ₦{diff:,.2f}."
        )
    return (
        f"Short by ₦{abs(diff):,.2f} ({label_prefix} ₦{detected_numeric:,.2f} vs Target ₦{target_payment_ngn:,.2f}). "
        f"Add ₦{abs(diff):,.2f}."
    )


def analyze_document_triage(text: str, target_payment_ngn: float = 0.0, basket_total_ngn: float = 0.0) -> dict:
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
        r"(?:Account|Meter|Ref(?:erence)?|Session|ID|Policy|Phone|Tel)[\s#:No.]*[A-Z0-9+-]{6,}",
        clean,
        flags=re.IGNORECASE
    )

    warning_keywords = (
        "warning", "caution", "danger", "due date", "deadline", "disconnection",
        "do not", "never", "important", "penalty", "expire", "overdose", "emergency", "failed", "declined"
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

    naira_info = detect_naira_banknotes(clean)
    receipt_info = analyze_screen_or_receipt(clean)

    detected_numeric = float(naira_info["total_ngn"])
    if detected_numeric == 0 and amounts:
        num_match = re.search(r"(\d[\d,]*(?:\.\d{1,2})?)", amounts[0])
        if num_match:
            try:
                detected_numeric = float(num_match.group(1).replace(",", ""))
            except ValueError:
                detected_numeric = 0.0

    effective_cash_total = float(basket_total_ngn) if basket_total_ngn > 0 else detected_numeric
    payment_status_line = format_payment_comparison(
        effective_cash_total,
        target_payment_ngn,
        label_prefix="Running cash total" if basket_total_ngn > 0 else "Detected"
    )

    spoken_parts = []
    if receipt_info:
        spoken_parts.append(receipt_info["spoken_receipt"])
    elif naira_info["notes"]:
        notes_str = ", ".join(f"{n} Naira" for n in naira_info["notes"])
        spoken_parts.append(f"This scan shows {notes_str}, subtotal {naira_info['total_ngn']} Naira.")
        if basket_total_ngn > 0:
            spoken_parts.append(f"Running cash basket total is {basket_total_ngn:,.0f} Naira.")
    else:
        spoken_parts.append("Key facts summary.")
        if amounts:
            spoken_parts.append(f"Amounts detected: {', '.join(amounts[:4])}.")
        if dates:
            spoken_parts.append(f"Important dates: {', '.join(dates[:3])}.")
        if dosages:
            spoken_parts.append(f"Dosage instructions: {'; '.join(dosages[:2])}.")

    if payment_status_line:
        spoken_parts.append(payment_status_line)
    if alerts:
        spoken_parts.append(f"Alert: {alerts[0]}")
    if len(spoken_parts) == 1 and spoken_parts[0] == "Key facts summary.":
        first_sentences = " ".join(re.split(r"(?<=[.!?])\s+", clean)[:2])
        spoken_parts.append(first_sentences[:350])

    return {
        "amounts": amounts,
        "dates": dates,
        "dosages": dosages,
        "identifiers": identifiers,
        "alerts": alerts,
        "naira_info": naira_info,
        "receipt_info": receipt_info,
        "payment_status_line": payment_status_line,
        "spoken_brief": " ".join(spoken_parts),
    }


def render_payment_target_and_basket(text: str, key_prefix: str):
    st.markdown("#### 🧮 Payment Target & Running Cash Basket")

    with st.expander("🎙️ Say Amount to Pay by Voice"):
        if not SPEECH_RECOGNITION_AVAILABLE:
            st.caption("Voice recognition package unavailable; use the number field below.")
        else:
            st.caption("Speak your payment target (e.g., 'three thousand five hundred' or '2000').")
            voice_consent = st.checkbox(
                "I agree to send this short voice clip to Google for number transcription.",
                key=f"{key_prefix}_amt_voice_consent"
            )
            amt_audio = st.audio_input("Record payment amount", key=f"{key_prefix}_amt_audio")
            if amt_audio and st.button(
                "Set amount from voice",
                key=f"{key_prefix}_set_amt_voice_btn",
                disabled=not voice_consent
            ):
                try:
                    spoken = transcribe_audio(amt_audio)
                    st.session_state.last_voice_amount_transcript = spoken
                    parsed_amt = parse_spoken_amount_ngn(spoken)
                    if parsed_amt:
                        st.session_state.target_payment_ngn = parsed_amt
                        st.success(f"Heard: '{spoken}' → Target set to ₦{parsed_amt:,.2f}")
                        st.rerun()
                    else:
                        st.warning(f"Heard: '{spoken}', but could not read a clear Naira amount.")
                except Exception as err:
                    st.error(f"Voice amount error: {err}")

    target_amt = st.number_input(
        "Amount you want to pay (₦ NGN)",
        min_value=0.0,
        step=50.0,
        value=float(st.session_state.get("target_payment_ngn", 0.0)),
        key=f"{key_prefix}_target_pay"
    )
    st.session_state.target_payment_ngn = target_amt

    naira_info = detect_naira_banknotes(text)
    basket_notes: list[int] = list(st.session_state.get("cash_basket_notes", []))
    basket_total = float(sum(basket_notes))

    b_col1, b_col2, b_col3, b_col4 = st.columns([2, 1, 1, 1])
    with b_col1:
        if naira_info.get("notes"):
            detected_str = ", ".join(f"₦{n:,}" for n in naira_info["notes"])
            st.write(f"**Current scan notes:** {detected_str} (₦{naira_info['total_ngn']:,})")
        else:
            st.caption("Snap a Naira banknote or load a Naira sample to add to your running basket.")

    with b_col2:
        if st.button(
            "➕ Add scan to basket",
            key=f"{key_prefix}_add_basket",
            disabled=not bool(naira_info.get("notes")),
            type="primary",
            use_container_width=True
        ):
            st.session_state.cash_basket_notes = basket_notes + list(naira_info["notes"])
            st.rerun()

    with b_col3:
        if st.button(
            "↩️ Undo last",
            key=f"{key_prefix}_undo_basket",
            disabled=not bool(basket_notes),
            use_container_width=True
        ):
            st.session_state.cash_basket_notes = basket_notes[:-1]
            st.rerun()

    with b_col4:
        if st.button(
            "🗑️ Clear basket",
            key=f"{key_prefix}_clear_basket",
            disabled=not bool(basket_notes),
            use_container_width=True
        ):
            st.session_state.cash_basket_notes = []
            st.rerun()

    basket_notes = list(st.session_state.get("cash_basket_notes", []))
    basket_total = float(sum(basket_notes))
    if basket_notes:
        notes_breakdown = ", ".join(f"₦{n:,}" for n in basket_notes)
        st.info(
            f"🧺 **Cash Basket ({len(basket_notes)} note(s)):** {notes_breakdown}  |  "
            f"**Running Total:** **₦{basket_total:,.2f}**"
        )

    return target_amt, basket_total


def render_instant_browser_tts(
    text: str,
    button_label: str = "🔊 Read Aloud Offline (Zero Data)",
    auto_play: bool = False
):
    payload = json.dumps(str(text or "")[:3500], ensure_ascii=False)
    payload = payload.replace("<", "\\u003c").replace(">", "\\u003e").replace("&", "\\u0026")
    safe_label = html.escape(button_label)
    auto_js = "true" if auto_play else "false"

    components.html(
        f"""
        <div style="display:flex; gap:8px; flex-wrap:wrap; align-items:center; font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
            <button id="speakBtn" style="background:#1D4ED8;color:#FFFFFF;border:none;border-radius:8px;padding:9px 14px;font-weight:600;cursor:pointer;">
                {safe_label}
            </button>
            <button id="stopBtn" style="background:#F3F4F6;color:#111827;border:1px solid #D1D5DB;border-radius:8px;padding:9px 14px;font-weight:600;cursor:pointer;">
                ⏹ Stop
            </button>
            <span id="ttsStatus" style="font-size:13px;color:#4B5563;"></span>
        </div>
        <script>
        const txt = {payload};
        const autoPlay = {auto_js};
        const statusEl = document.getElementById('ttsStatus');
        function speakNow() {{
            if (!('speechSynthesis' in window)) {{
                statusEl.textContent = 'Browser speech unavailable.';
                return;
            }}
            window.speechSynthesis.cancel();
            const u = new SpeechSynthesisUtterance(txt);
            u.onstart = () => statusEl.textContent = 'Reading offline...';
            u.onend = () => statusEl.textContent = 'Done.';
            window.speechSynthesis.speak(u);
        }}
        document.getElementById('speakBtn').onclick = speakNow;
        document.getElementById('stopBtn').onclick = () => {{
            if ('speechSynthesis' in window) window.speechSynthesis.cancel();
            statusEl.textContent = 'Stopped.';
        }};
        if (autoPlay && txt.trim().length > 0) {{
            setTimeout(speakNow, 250);
        }}
        </script>
        """,
        height=52,
    )


def render_triage_panel(text: str, key_prefix: str):
    target_amt, basket_total = render_payment_target_and_basket(text, key_prefix)

    triage = analyze_document_triage(
        text,
        target_payment_ngn=target_amt,
        basket_total_ngn=basket_total
    )
    if not triage:
        return

    st.markdown(
        """
        <div class="ns-triage-card">
            <strong>⚡ Key Facts, Naira Cash & Screen Triage</strong><br>
            <span>Banknote counts, running cash basket, payment screen status, amounts, dates, and warnings surfaced first.</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    naira_info = triage.get("naira_info", {})
    if naira_info.get("notes"):
        notes_list = ", ".join(f"₦{n:,}" for n in naira_info["notes"])
        st.success(
            f"💵 **Scan Notes:** {notes_list}  |  "
            f"**Scan Subtotal:** ₦{naira_info['total_ngn']:,}  |  "
            f"**CBN Text:** {'Verified' if naira_info['cbn_verified'] else 'Partial'}"
        )

    color_hint = st.session_state.get("last_image_color_hint")
    if color_hint and naira_info.get("notes"):
        st.caption(f"Visual note tint check: {color_hint}")

    receipt_info = triage.get("receipt_info", {})
    if receipt_info:
        r_status = receipt_info.get("status", "Unknown")
        r_amt = receipt_info.get("amount") or "Not shown"
        r_rec = receipt_info.get("recipient") or "Not shown"
        st.info(f"📱 **Screen / Receipt Readout:** Status **{r_status}** · Amount **{r_amt}** · Recipient **{r_rec}**")

    if triage.get("payment_status_line"):
        st.warning(f"🧮 **Payment Check:** {triage['payment_status_line']}")

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
        button_label="🔊 Speak Key Facts Offline (Instant / Zero Data)",
        auto_play=bool(st.session_state.get("auto_speak_triage", False))
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
# 22. TEXT TO SPEECH, OFFLINE READER & VIDEO VOICEOVER ENGINE
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


def resolve_ffmpeg_binary() -> str | None:
    system_ffmpeg = shutil.which("ffmpeg")
    if system_ffmpeg:
        return system_ffmpeg
    if IMAGEIO_FFMPEG_AVAILABLE:
        try:
            exe = imageio_ffmpeg.get_ffmpeg_exe()
            if exe and os.path.exists(exe):
                return exe
        except Exception:
            pass
    return None


def probe_video_has_audio(ffmpeg_bin: str, video_path: Path) -> bool:
    """Return True if the input video contains at least one audio stream."""
    proc = subprocess.run(
        [ffmpeg_bin, "-i", str(video_path)],
        capture_output=True,
        text=True
    )
    stderr = proc.stderr or ""
    return "Audio:" in stderr


def probe_video_duration_sec(ffmpeg_bin: str, video_path: Path) -> float | None:
    proc = subprocess.run(
        [ffmpeg_bin, "-i", str(video_path)],
        capture_output=True,
        text=True
    )
    m = re.search(r"Duration:\s*(\d+):(\d+):(\d+(?:\.\d+)?)", proc.stderr or "")
    if not m:
        return None
    hours = int(m.group(1))
    minutes = int(m.group(2))
    seconds = float(m.group(3))
    return hours * 3600 + minutes * 60 + seconds


def render_embedded_voiceover_video(
    video_bytes: bytes,
    commentary_text: str,
    resolution_choice: str = "1080p (1K · 1920x1080)",
    existing_audio_mode: str = "Keep & Mix with Commentary",
    slow_voice: bool = False,
    enhance_visuals: bool = True,
    pad_video_to_voice: bool = True,
) -> tuple[bytes, str]:
    """
    Generate speech voiceover from commentary_text, scale/enhance video to 720p/1K/4K,
    keep or silence existing video audio, and embed the commentary track into an MP4.
    """
    clean_script = (commentary_text or "").strip()
    if not clean_script:
        raise ValueError("Please enter or generate commentary text for the voiceover.")

    ffmpeg_bin = resolve_ffmpeg_binary()
    if not ffmpeg_bin:
        raise RuntimeError(
            "FFmpeg was not found. Install system ffmpeg or run `pip install imageio-ffmpeg` to embed voiceovers into video."
        )

    res_map = {
        "720p (1280x720)": (1280, 720, "2500k"),
        "1080p (1K · 1920x1080)": (1920, 1080, "5000k"),
        "4K (3840x2160)": (3840, 2160, "14000k"),
    }
    target_w, target_h, video_bitrate = res_map.get(resolution_choice, (1920, 1080, "5000k"))

    # Create MP3 via gTTS (chunked if longer than 4000 chars)
    chunks = split_text_for_speech(clean_script, max_chars=3800)
    if not chunks:
        raise ValueError("Commentary script has no readable words.")

    with tempfile.TemporaryDirectory() as tmpdir:
        tmp_path = Path(tmpdir)
        in_video = tmp_path / "input_video.mp4"
        out_video = tmp_path / "output_voiceover.mp4"
        voice_wav = tmp_path / "voiceover_clean.wav"

        in_video.write_bytes(video_bytes)

        # Synthesize each chunk and convert/concat into a single WAV voice track
        chunk_wavs = []
        for idx, part in enumerate(chunks):
            mp3_buf = create_speech(part, slow=slow_voice)
            part_mp3 = tmp_path / f"part_{idx}.mp3"
            part_wav = tmp_path / f"part_{idx}.wav"
            part_mp3.write_bytes(mp3_buf.getvalue())

            conv = subprocess.run(
                [
                    ffmpeg_bin, "-y",
                    "-i", str(part_mp3),
                    "-ac", "1",
                    "-ar", "44100",
                    str(part_wav)
                ],
                capture_output=True,
                text=True
            )
            if conv.returncode != 0:
                raise RuntimeError(f"Audio conversion failed: {conv.stderr[-300:]}")
            chunk_wavs.append(part_wav)

        if len(chunk_wavs) == 1:
            shutil.copyfile(chunk_wavs[0], voice_wav)
        else:
            concat_list = tmp_path / "concat_list.txt"
            concat_list.write_text(
                "\n".join(f"file '{w.as_posix()}'" for w in chunk_wavs),
                encoding="utf-8"
            )
            cat_proc = subprocess.run(
                [
                    ffmpeg_bin, "-y",
                    "-f", "concat",
                    "-safe", "0",
                    "-i", str(concat_list),
                    "-ac", "1",
                    "-ar", "44100",
                    str(voice_wav)
                ],
                capture_output=True,
                text=True
            )
            if cat_proc.returncode != 0:
                raise RuntimeError(f"Voice track join failed: {cat_proc.stderr[-300:]}")

        # Build video filter chain: aspect-preserving scale + pad + optional clarity enhancement
        vf_parts = [
            f"scale={target_w}:{target_h}:force_original_aspect_ratio=decrease",
            f"pad={target_w}:{target_h}:(ow-iw)/2:(oh-ih)/2:black",
        ]
        if enhance_visuals:
            vf_parts.append("eq=contrast=1.06:brightness=0.01:saturation=1.08")
            vf_parts.append("unsharp=5:5:0.6:3:3:0.3")
        vf_parts.append("format=yuv420p")
        vf_chain = ",".join(vf_parts)

        has_orig_audio = probe_video_has_audio(ffmpeg_bin, in_video)
        cmd = [ffmpeg_bin, "-y", "-i", str(in_video), "-i", str(voice_wav)]

        filter_complex_nodes = []
        audio_map_label = "[aout]"

        if existing_audio_mode == "Silence / Mute Original Audio" or not has_orig_audio:
            filter_complex_nodes.append("[1:a]aresample=44100,volume=1.15[aout]")
        else:
            # Keep original audio at lower ducked volume so commentary stays crisp and intelligible
            filter_complex_nodes.append(
                "[0:a]aresample=44100,volume=0.22[orig_duck];"
                "[1:a]aresample=44100,volume=1.20[vo_loud];"
                "[orig_duck][vo_loud]amix=inputs=2:duration=longest:dropout_transition=2[aout]"
            )

        cmd.extend(["-filter_complex", ";".join(filter_complex_nodes)])
        cmd.extend([
            "-vf", vf_chain,
            "-c:v", "libx264",
            "-preset", "medium",
            "-b:v", video_bitrate,
            "-c:a", "aac",
            "-b:a", "192k",
            "-map", "0:v:0",
            "-map", audio_map_label,
            "-movflags", "+faststart",
        ])

        if pad_video_to_voice:
            # Keep full voiceover embedded even if the clip is shorter by extending last frame smoothly
            cmd.extend(["-vf", f"tpad=stop_mode=clone:stop_duration=30,{vf_chain}", "-shortest"])
        else:
            cmd.append("-shortest")

        cmd.append(str(out_video))

        # Simpler, rock-solid filter graph without duplicate -vf flags
        clean_cmd = [
            ffmpeg_bin, "-y",
            "-i", str(in_video),
            "-i", str(voice_wav),
        ]
        fc = list(filter_complex_nodes)
        if pad_video_to_voice:
            fc.append(f"[0:v]tpad=stop_mode=clone:stop_duration=60,{vf_chain}[vout]")
        else:
            fc.append(f"[0:v]{vf_chain}[vout]")

        clean_cmd = [
            ffmpeg_bin, "-y",
            "-i", str(in_video),
            "-i", str(voice_wav),
            "-filter_complex", ";".join(fc),
            "-map", "[vout]",
            "-map", audio_map_label,
            "-c:v", "libx264",
            "-preset", "medium",
            "-b:v", video_bitrate,
            "-c:a", "aac",
            "-b:a", "192k",
            "-movflags", "+faststart",
            "-shortest" if not pad_video_to_voice else "-longest",
            str(out_video),
        ]

        # Note: in ffmpeg, -shortest stops when shortest stream ends; if we want full voiceover, omit -shortest
        if pad_video_to_voice:
            clean_cmd = [arg for arg in clean_cmd if arg != "-longest"]

        proc = subprocess.run(clean_cmd, capture_output=True, text=True)
        if proc.returncode != 0 or not out_video.exists():
            raise RuntimeError(f"Video voiceover rendering failed: {proc.stderr[-400:]}")

        res_tag = "4K" if "4K" in resolution_choice else ("1K_1080p" if "1080p" in resolution_choice else "720p")
        out_name = f"naijasight_voiceover_{res_tag}.mp4"
        return out_video.read_bytes(), out_name


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
        button_label=f"🔊 Read Part {selected_part + 1} Offline (Browser Voice)",
        auto_play=False
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
<meta name="theme-color" content="#1d4ed8">
<title>{safe_title}</title>
<style>
:root {{ color-scheme: light; --bg:#f9fafb; --panel:#ffffff; --ink:#111827; --accent:#1d4ed8; --muted:#4b5563; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif; line-height:1.65; }}
header {{ background:#ffffff; color:#111827; border-bottom:1px solid #e5e7eb; padding:18px 16px; }}
header h1 {{ margin:0 auto; max-width:900px; font-size:1.35rem; }}
main {{ max-width:900px; margin:20px auto; padding:0 14px 40px; }}
.panel {{ background:var(--panel); border:1px solid #e5e7eb; border-radius:12px; padding:16px; margin-bottom:16px; }}
.controls {{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; }}
button, select {{ min-height:42px; border-radius:8px; border:1px solid #d1d5db; padding:8px 13px; font-size:0.98rem; background:#f9fafb; color:#111827; cursor:pointer; }}
button.primary {{ background:var(--accent); color:#ffffff; border-color:var(--accent); }}
button:focus, select:focus, input:focus {{ outline:2px solid #1d4ed8; outline-offset:2px; }}
label {{ display:block; font-weight:600; margin:10px 0 4px; }}
#readerText {{ white-space:pre-wrap; overflow-wrap:anywhere; font-size:1.15rem; line-height:1.8; }}
#status {{ color:var(--muted); min-height:1.5em; }}
.small {{ color:var(--muted); font-size:.92rem; }}
body.high-contrast {{ --bg:#ffffff; --panel:#ffffff; --ink:#000000; --accent:#000000; --muted:#1f2937; color:#000000; }}
body.high-contrast header {{ background:#ffffff; color:#000000; border-bottom:2px solid #000000; }}
body.high-contrast button, body.high-contrast select {{ background:#ffffff; color:#000000; border:2px solid #000000; }}
body.high-contrast button.primary {{ background:#000000; color:#ffffff; }}
@media (max-width:600px) {{
  header h1 {{ font-size:1.15rem; }}
  main {{ margin:12px auto; padding:0 10px 28px; }}
  .panel {{ padding:13px; }}
  .controls button, .controls select {{ flex:1 1 42%; }}
  #readerText {{ font-size:1.1rem; }}
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
let fontSize = 1.15;

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
        <div class="ns-header-bar">
            <div>
                <h1>👁️ NaijaSight</h1>
                <p>Assistive Naira Cash Verification, Video Voiceover Studio, Screen/Bill Reader & Voice Suite</p>
            </div>
            <div class="ns-chip-row">
                <span class="ns-chip">720p / 1K / 4K Voiceover</span>
                <span class="ns-chip">Running Cash Basket</span>
                <span class="ns-chip">Offline Voice Ready</span>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_left, col_right = st.columns([1.15, 1], gap="large")

    with col_left:
        st.subheader("Everyday Independence for Blind & Low-Vision Users")
        st.write(
            "Snap Naira notes one by one into a running cash basket, speak payment targets, "
            "generate and embed voiceovers into 720p/1K/4K videos, triage bills, and listen offline."
        )
        st.markdown(
            """
            - **Video Voiceover & Commentary Studio:** Generate and embed voiceovers at 720p, 1080p (1K), or 4K with clarity enhancement and original-audio mute/mix control
            - **Multi-Note Cash Basket & Change Check:** Add scanned notes one by one, undo mistakes, and hear totals or expected change
            - **Screen & Document Triage:** Upload bank/POS screenshots, PDFs, Word, Excel, or images with instant key facts
            - **Offline Voice & HTML Reader:** Instant browser speech, optional hands-free auto-read, and standalone offline HTML readers
            - **Accessibility & AI Ethics:** Log real PWD testing observations and maintain transparent safety guardrails
            """
        )

        st.divider()
        if st.button("Continue with Guest Demo", type="primary", use_container_width=True):
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
                    st.subheader("Sign in")
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
                    st.subheader("Create an account")
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
        st.markdown("### 👁️ NaijaSight")
        st.caption(f"Workspace · v{APP_VERSION}")
        st.divider()

        st.markdown(f"**{html.escape(st.session_state.full_name or 'User')}**")
        if st.session_state.demo_mode:
            st.caption("Guest demo session (history not saved)")

        pages = [
            "Home",
            "Scan with camera",
            "Read a document",
            "Video voiceover",
            "Listen to text",
            "Voice tools",
            "Activity history",
            "Accessibility & AI ethics"
        ]
        page = st.radio("Modules", pages, key="nav_page")

        st.divider()
        st.markdown("**Voice & Display**")
        st.toggle(
            "Hands-free auto-speak key facts",
            key="auto_speak_triage"
        )
        st.selectbox(
            "Theme",
            ["Standard (Crisp Light)", "High Contrast (Pure White / Black)", "Low-Glare Slate"],
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
    basket_total = sum(st.session_state.get("cash_basket_notes", []))

    st.markdown(
        f"""
        <div class="ns-header-bar">
            <div>
                <h2>{html.escape(active_page)}</h2>
                <p>Signed in as <strong>{html.escape(st.session_state.full_name or '')}</strong> (@{html.escape(st.session_state.username or '')})</p>
            </div>
            <div class="ns-chip-row">
                <span class="ns-chip">Cash basket: ₦{basket_total:,}</span>
                <span class="ns-chip">Saved records: {len(records)}</span>
                <span class="ns-chip">Words processed: {total_words:,}</span>
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
    st.subheader("Workspace Modules")
    st.caption("Choose a tool below to get started.")

    cards = [
        ("📷", "Scan with camera", "Snap Naira notes into a running basket, set payment amount by voice or number, and hear totals or change aloud.", "Scan with camera"),
        ("📄", "Read a document", "Upload screenshots of bank/POS apps, PDFs, Word, Excel, or images with instant key-fact triage.", "Read a document"),
        ("🎬", "Video voiceover", "Generate and embed voiceover commentary into videos at 720p, 1K, or 4K with clarity boost and original-audio mute/mix.", "Video voiceover"),
        ("🔊", "Listen to text", "Read extracted or pasted text aloud offline or in clean MP3 parts and download an offline reader.", "Listen to text"),
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
    st.subheader("How It Works")
    st.write("1. Pick a module.  2. Set payment targets, snap cash/screens, or upload a video.  3. Generate commentary & embed voiceover.  4. Listen offline or export.")
    st.info("Keep banknotes flat and well-lit when snapping, and preview your commentary script before embedding into 720p/1K/4K video.")


# ============================================================
# 26B. VIDEO VOICEOVER & EMBEDDED COMMENTARY STUDIO
# ============================================================

def render_video_voiceover_page():
    st.subheader("Video Voiceover & Embedded Commentary Studio")
    st.write(
        "Upload a video, write or load commentary text, choose whether to keep or silence the existing audio, "
        "and render a clean 720p, 1K (1080p), or 4K video with the voiceover permanently embedded."
    )

    ffmpeg_bin = resolve_ffmpeg_binary()
    if not ffmpeg_bin:
        st.warning(
            "FFmpeg is not installed yet. Install `imageio-ffmpeg` (`pip install imageio-ffmpeg`) "
            "or system `ffmpeg` to render and embed voiceover tracks into video files."
        )

    uploaded_video = st.file_uploader(
        "Upload a video file",
        type=["mp4", "mov", "mkv", "avi", "webm"],
        max_upload_size=MAX_FILE_SIZE_MB,
        key="vo_video_uploader"
    )

    if uploaded_video:
        st.caption(f"Selected video: **{uploaded_video.name}** ({uploaded_video.size / (1024 * 1024):.2f} MB)")
        st.video(uploaded_video)

    seed_text = st.session_state.get("extracted_text", "")
    if st.button("Use current extracted text as commentary script", key="vo_use_extracted"):
        if seed_text.strip():
            st.session_state["vo_script_input"] = seed_text.strip()[:3500]
            st.rerun()
        else:
            st.info("No extracted text in the workspace yet.")

    default_script = st.session_state.get("vo_script_input", "")
    commentary_script = st.text_area(
        "Voiceover / commentary script (must be clear and accurate before rendering)",
        value=default_script,
        height=170,
        placeholder="Type the exact narration or commentary that will be spoken and embedded into the video...",
        key="vo_script_input"
    )

    if commentary_script.strip():
        render_instant_browser_tts(
            commentary_script,
            button_label="🔊 Preview Commentary Offline Before Rendering"
        )

    c1, c2 = st.columns(2)
    with c1:
        resolution_choice = st.selectbox(
            "Output resolution quality",
            ["720p (1280x720)", "1080p (1K · 1920x1080)", "4K (3840x2160)"],
            index=1,
            key="vo_resolution"
        )
        enhance_visuals = st.checkbox(
            "Apply visual clarity, contrast & color enhancement",
            value=True,
            key="vo_enhance_visuals"
        )

    with c2:
        existing_audio_mode = st.radio(
            "Existing video voiceover / audio track",
            ["Silence / Mute Original Audio", "Keep & Mix with Commentary"],
            key="vo_existing_audio_mode"
        )
        slow_voice = st.checkbox(
            "Speak commentary at a slower, clearer pace",
            value=False,
            key="vo_slow_voice"
        )

    pad_video_to_voice = st.checkbox(
        "Hold final frame if voiceover runs slightly longer than the video clip",
        value=True,
        key="vo_pad_video"
    )

    st.warning(
        "Generating voiceover audio sends the commentary script to Google's text-to-speech service. "
        "Confirm your commentary text is accurate and free of private information."
    )
    vo_consent = st.checkbox(
        "I confirm the commentary script is accurate and I agree to generate the voiceover audio.",
        key="vo_consent_chk"
    )

    can_render = bool(uploaded_video and commentary_script.strip() and vo_consent and ffmpeg_bin)
    if st.button(
        "Generate & Embed Voiceover into Video",
        type="primary",
        disabled=not can_render,
        key="vo_render_btn"
    ):
        try:
            t0 = time.perf_counter()
            raw_video_bytes = check_file_size(uploaded_video)
            with st.spinner(f"Generating voiceover and rendering {resolution_choice} video..."):
                out_bytes, out_filename = render_embedded_voiceover_video(
                    video_bytes=raw_video_bytes,
                    commentary_text=commentary_script,
                    resolution_choice=resolution_choice,
                    existing_audio_mode=existing_audio_mode,
                    slow_voice=slow_voice,
                    enhance_visuals=enhance_visuals,
                    pad_video_to_voice=pad_video_to_voice,
                )
            elapsed_ms = round((time.perf_counter() - t0) * 1000, 1)
            st.session_state.rendered_video_bytes = out_bytes
            st.session_state.rendered_video_name = out_filename
            log_activity(
                f"{uploaded_video.name} [{resolution_choice}]",
                "video_voiceover",
                commentary_script
            )
            st.success(f"Embedded voiceover video ready in {elapsed_ms} ms.")
        except Exception as err:
            st.error(f"Could not generate voiceover video: {err}")

    if st.session_state.get("rendered_video_bytes"):
        st.divider()
        st.markdown("#### Rendered Video With Embedded Voiceover")
        st.video(st.session_state.rendered_video_bytes)
        st.download_button(
            "Download Video With Embedded Voiceover (.MP4)",
            data=st.session_state.rendered_video_bytes,
            file_name=st.session_state.get("rendered_video_name") or "naijasight_voiceover.mp4",
            mime="video/mp4",
            type="primary",
            use_container_width=True,
            key="vo_download_btn"
        )


# ============================================================
# 27. RESULT WORKSPACE
# ============================================================

def render_result_workspace():
    text = st.session_state.extracted_text

    if not text:
        st.info("Extracted or transcribed text will appear here.")
        return

    st.subheader("Extracted Text & Triage")
    st.caption("OCR confidence estimates region clarity rather than guaranteeing every character.")

    elapsed_ms = st.session_state.get("last_process_ms")
    confidence = st.session_state.get("last_ocr_confidence")
    if confidence is not None:
        m1, m2 = st.columns(2)
        m1.metric("OCR confidence", f"{confidence}%")
        if elapsed_ms is not None:
            m2.metric("Processing time", f"{elapsed_ms} ms")

        if confidence < 60:
            st.warning("Some regions had low OCR confidence. Review carefully before relying on numbers or names.")
        elif confidence < 80:
            st.info("Moderate OCR confidence. Double-check key details.")
        else:
            st.success("High OCR confidence across detected text regions.")

        uncertain = st.session_state.get("last_uncertain_words", [])
        if uncertain:
            with st.expander("Review uncertain words"):
                for item in uncertain:
                    st.write(f"**{item['text']}** — {item['confidence']}% confidence")
    elif elapsed_ms is not None:
        st.caption(f"{len(text.split())} words · {len(text)} characters · {elapsed_ms} ms")
    else:
        st.caption(f"{len(text.split())} words · {len(text)} characters")

    if st.session_state.get("editable_result") != text:
        st.session_state.editable_result = text

    edited_text = st.text_area(
        "Review or edit text",
        height=240,
        key="editable_result"
    )
    st.session_state.extracted_text = edited_text

    render_triage_panel(edited_text, "workspace_triage")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.download_button(
            "Download (.TXT)",
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
    st.subheader("Listen to text")
    render_speech_controls(edited_text, "result")


# ============================================================
# 28. DOCUMENT READER PAGE
# ============================================================

def render_document_reader():
    st.subheader("Read a Document, Screenshot, or Image")
    st.write("Upload a screenshot of a bank/POS screen, bill, prescription, or document—or load a sample below.")

    with st.expander("Sample scenarios (Naira notes, bank screen, bill & prescription)"):
        sample_choice = st.selectbox("Choose a sample", list(SAMPLE_DOCUMENTS.keys()))
        if st.button("Load sample", key="load_sample_doc"):
            t0 = time.perf_counter()
            sample_text = SAMPLE_DOCUMENTS[sample_choice]
            st.session_state.extracted_text = sample_text
            st.session_state.editable_result = sample_text
            st.session_state.current_filename = f"{sample_choice}.txt"
            st.session_state.current_file_type = ".txt"
            st.session_state.last_ocr_confidence = None
            st.session_state.last_uncertain_words = []
            st.session_state.last_image_color_hint = None
            st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)
            log_activity(f"{sample_choice}.txt", ".txt", sample_text)
            st.success("Sample loaded.")
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
                "OCR preparation",
                ["Standard", "Enhanced"],
                horizontal=True
            )

        st.caption(f"Size: {uploaded_file.size / (1024 * 1024):.2f} MB")

        if st.button("Extract text", type="primary"):
            clear_current_result()
            try:
                t0 = time.perf_counter()
                with st.spinner("Extracting text..."):
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
                    st.success("Extraction complete.")
                else:
                    st.warning("No readable text was found in this file.")

            except Exception as error:
                st.error(f"Could not read the file: {error}")

    render_result_workspace()


# ============================================================
# 29. CAMERA SCANNER (MONEY, SCREENS & PRINTED PAGES)
# ============================================================

def render_camera_scanner():
    st.subheader("Scan Money, Screen, or Printed Page")
    st.write("Snap Naira banknotes one by one into your running cash basket, check payment/change, or read a screen/bill.")

    pay_amt = st.number_input(
        "Amount you want to pay (₦ NGN — set before or after snapping)",
        min_value=0.0,
        step=50.0,
        value=float(st.session_state.get("target_payment_ngn", 0.0)),
        key="camera_target_pay_input"
    )
    st.session_state.target_payment_ngn = pay_amt

    c1, c2 = st.columns(2)
    with c1:
        camera_mode = st.radio(
            "Reading scope",
            ["Read Entire Image", "Read Main Document"],
            horizontal=True
        )
    with c2:
        camera_quality = st.radio(
            "OCR preparation",
            ["Standard", "Enhanced"],
            horizontal=True
        )

    camera_image = st.camera_input("Take a picture")

    if camera_image:
        if st.button("Read & verify picture", type="primary"):
            clear_current_result()
            try:
                t0 = time.perf_counter()
                with st.spinner("Reading picture..."):
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
                    st.success("Picture read and verified.")
                else:
                    st.warning("No text or banknote figures detected. Try better lighting and hold steady.")

            except Exception as error:
                st.error(f"Could not process picture: {error}")

    render_result_workspace()


# ============================================================
# 30. LISTEN TO TEXT
# ============================================================

def render_listen_page():
    st.subheader("Listen to Text")
    st.write("Paste text, load a sample, or use your current workspace text.")

    if st.button("Load sample ₦1,000 note", key="listen_sample_btn"):
        st.session_state["standalone_listen_text"] = SAMPLE_DOCUMENTS["💵 Single ₦1,000 Note (Snap & Add to Basket)"]
        st.rerun()

    initial_text = st.session_state.get("extracted_text", "")
    text_to_read = st.text_area(
        "Text to read aloud",
        value=initial_text,
        height=240,
        key="standalone_listen_text"
    )
    if text_to_read.strip():
        render_triage_panel(text_to_read, "standalone_triage")
    render_speech_controls(text_to_read, "standalone")


# ============================================================
# 31. VOICE TOOLS
# ============================================================

def render_voice_tools():
    st.subheader("Voice Transcription")
    st.write("Record a short spoken note and convert it into editable text.")

    if not SPEECH_RECOGNITION_AVAILABLE:
        st.error("The speech recognition package is not installed.")
        return

    st.warning(
        "Voice transcription sends audio to Google's speech service. "
        "Do not record sensitive or private details."
    )
    consent = st.checkbox("I agree to send my recording to Google for transcription.")
    audio_file = st.audio_input("Record voice")

    if audio_file:
        st.audio(audio_file)
        if st.button("Transcribe recording", disabled=not consent, type="primary"):
            try:
                t0 = time.perf_counter()
                with st.spinner("Transcribing..."):
                    transcript = transcribe_audio(audio_file)
                st.session_state.last_process_ms = round((time.perf_counter() - t0) * 1000, 1)

                st.session_state.extracted_text = transcript
                st.session_state.current_filename = "Voice recording"
                st.session_state.current_file_type = "audio"
                st.session_state.editable_result = transcript

                log_activity("Voice recording", "audio", transcript)
                st.success("Transcription complete.")
            except Exception as error:
                st.error(f"Could not transcribe recording: {error}")

    render_result_workspace()


# ============================================================
# 32. ACTIVITY HISTORY, SEARCH, EXPORTS & RETENTION
# ============================================================

def render_history():
    st.subheader("Activity History & Exports")

    if st.session_state.demo_mode:
        st.info("Sign in with a registered account to store and manage history across sessions.")
        return

    records = get_activity_records()
    if not records:
        st.info("No saved activity records yet.")
        return

    total_words = sum(r["word_count"] for r in records)
    total_chars = sum(r["character_count"] for r in records)

    k1, k2, k3 = st.columns(3)
    with k1:
        st.markdown(
            f'<div class="ns-kpi-card"><div class="ns-kpi-label">Saved extractions</div><div class="ns-kpi-value">{len(records)}</div></div>',
            unsafe_allow_html=True
        )
    with k2:
        st.markdown(
            f'<div class="ns-kpi-card"><div class="ns-kpi-label">Total words</div><div class="ns-kpi-value">{total_words:,}</div></div>',
            unsafe_allow_html=True
        )
    with k3:
        st.markdown(
            f'<div class="ns-kpi-card"><div class="ns-kpi-label">Total characters</div><div class="ns-kpi-value">{total_chars:,}</div></div>',
            unsafe_allow_html=True
        )

    st.markdown("#### Search & Filter")
    f_col1, f_col2 = st.columns([2, 1])
    with f_col1:
        search_query = st.text_input(
            "Search filename or text",
            placeholder="e.g., naira, receipt, bill, prescription..."
        ).strip().lower()

    available_types = sorted({(r["file_type"] or "unknown") for r in records})
    with f_col2:
        type_filter = st.selectbox(
            "File type",
            options=["All Types"] + available_types
        )

    filtered_records = []
    for r in records:
        matches_type = (type_filter == "All Types") or ((r["file_type"] or "unknown") == type_filter)
        haystack = f"{r['filename']} {r['file_type']} {r.get('extracted_text', '')}".lower()
        matches_query = (not search_query) or (search_query in haystack)
        if matches_type and matches_query:
            filtered_records.append(r)

    st.caption(f"Showing {len(filtered_records)} of {len(records)} record(s).")

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
                "Export Excel (.XLSX)",
                data=build_excel_export_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_records.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )
        with exp_col2:
            st.download_button(
                "Export Report (.PDF)",
                data=build_pdf_report_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_report.pdf",
                mime="application/pdf",
                use_container_width=True
            )
        with exp_col3:
            st.download_button(
                "Audit Snapshot (.JSON)",
                data=build_audit_json_bytes(filtered_records),
                file_name=f"naijasight_{st.session_state.username}_audit.json",
                mime="application/json",
                use_container_width=True
            )
    else:
        st.warning("No records match your search or filter.")

    st.divider()
    st.subheader("Manage Saved Record")

    pool_for_selection = filtered_records if filtered_records else records
    record_map = {
        f"#{r['id']} — {r['filename']} ({r['word_count']} words, {r['created_at']})": r
        for r in pool_for_selection
    }
    selected_label = st.selectbox("Saved record", list(record_map.keys()))
    chosen_record = record_map[selected_label]

    if chosen_record.get("extracted_text"):
        with st.expander("Preview text"):
            st.write(chosen_record["extracted_text"])

    col_reload, col_del_one = st.columns(2)
    with col_reload:
        if st.button("Load into workspace", use_container_width=True):
            st.session_state.extracted_text = chosen_record.get("extracted_text", "")
            st.session_state.editable_result = chosen_record.get("extracted_text", "")
            st.session_state.current_filename = chosen_record["filename"]
            st.session_state.current_file_type = chosen_record["file_type"]
            st.session_state.nav_page = "Read a document"
            st.rerun()

    with col_del_one:
        if st.button("Delete record", use_container_width=True):
            delete_activity_item(chosen_record["id"])
            st.success("Record deleted.")
            st.rerun()

    st.divider()
    st.subheader("Clear Activity History")
    st.warning("This permanently deletes all activity records under your account.")
    confirm_delete = st.checkbox("I understand this cannot be undone.")
    if st.button("Delete all history", disabled=not confirm_delete):
        delete_activity()
        st.success("Activity history deleted.")
        st.rerun()


# ============================================================
# 32B. ACCESSIBILITY TESTING LOG & RESPONSIBLE AI NOTE
# ============================================================

def render_accessibility_and_ethics():
    st.subheader("Accessibility & AI Ethics")
    st.write(
        "Track usability observations from real testing sessions, document fixes, "
        "and maintain clear consent and accuracy guardrails."
    )

    st.markdown(
        """
        - **Explicit Consent:** Cloud speech, voice transcription, and video voiceover generation require explicit user confirmation before processing.
        - **Low-Vision & Offline Support:** Adjustable editor font size, contrast themes, hands-free auto-speak option, offline browser speech, and standalone offline HTML readers.
        - **Financial, Video & Accuracy Guardrails:** Multi-note cash baskets, payment/change checks, explicit script review before video embedding, and OCR confidence scores remind users to verify critical details.
        """
    )

    st.divider()
    st.markdown("#### Log a Testing Observation")
    with st.form("a11y_feedback_form"):
        tester_group = st.text_input(
            "Tester profile or context",
            placeholder="e.g., Blind user testing video commentary / Low-vision mobile tester"
        )
        severity = st.selectbox("Severity", ["High", "Medium", "Low"])
        issue_found = st.text_area("Issue observed", height=90)
        fix_applied = st.text_area("Fix or improvement made", height=90)
        submitted = st.form_submit_button("Save observation", type="primary")

        if submitted:
            if not tester_group.strip() or not issue_found.strip() or not fix_applied.strip():
                st.warning("Please complete tester profile, issue observed, and fix made.")
            else:
                log_accessibility_feedback(tester_group, issue_found, fix_applied, severity)
                st.success("Observation saved.")
                st.rerun()

    feedback_df = get_accessibility_feedback_df()
    st.markdown("#### Logged Observations")
    if feedback_df.empty:
        st.info("No testing observations logged yet.")
    else:
        st.dataframe(feedback_df, use_container_width=True, hide_index=True)


# ============================================================
# 33. LOGGED-IN PORTAL ROUTER
# ============================================================

def render_logged_in_portal():
    page = render_sidebar()
    apply_portal_styles(
        vision_theme=st.session_state.get("vision_theme", "Standard (Crisp Light)"),
        reader_font_px=st.session_state.get("reader_font_px", 18)
    )
    render_portal_header(page)

    if page == "Home":
        render_home()
    elif page == "Scan with camera":
        render_camera_scanner()
    elif page == "Read a document":
        render_document_reader()
    elif page == "Video voiceover":
        render_video_voiceover_page()
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
        "NaijaSight · Verify important cash amounts, bank transfers, commentary scripts, dates, and dosages."
    )


# ============================================================
# 34. APPLICATION ENTRY POINT
# ============================================================

if not st.session_state.authenticated:
    apply_portal_styles(
        vision_theme=st.session_state.get("vision_theme", "Standard (Crisp Light)"),
        reader_font_px=st.session_state.get("reader_font_px", 18)
    )
    render_auth_portal()
else:
    render_logged_in_portal()