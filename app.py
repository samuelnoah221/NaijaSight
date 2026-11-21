import streamlit as st
import easyocr
import fitz
import numpy as np
import pandas as pd
import io
import re
import json
import sqlite3
import hashlib
import secrets

from pathlib import Path
from datetime import datetime
from PIL import Image, ImageOps, ImageEnhance
from gtts import gTTS

from docx import Document
from pptx import Presentation
from striprtf.striprtf import rtf_to_text
from ebooklib import epub, ITEM_DOCUMENT
from bs4 import BeautifulSoup

try:
    import speech_recognition as sr
    SPEECH_RECOGNITION_AVAILABLE = True
except ImportError:
    SPEECH_RECOGNITION_AVAILABLE = False


# ============================================================
# 1. APPLICATION SETTINGS
# ============================================================

st.set_page_config(
    page_title="NaijaSight",
    page_icon="👁️",
    layout="wide",
    initial_sidebar_state="expanded"
)

APP_NAME = "NaijaSight"

BASE_DIR = Path(__file__).parent
DB_PATH = BASE_DIR / "naijasight_portal.db"

MAX_FILE_SIZE_MB = 500
MAX_FILE_SIZE_BYTES = MAX_FILE_SIZE_MB * 1024 * 1024

MAX_IMAGE_PIXELS = 25_000_000
MAX_PDF_PAGES = 100

Image.MAX_IMAGE_PIXELS = MAX_IMAGE_PIXELS


# ============================================================
# 2. PAGE STYLING
# ============================================================

st.markdown(
    """
    <style>
    .main {
        padding-top: 1.5rem;
    }

    .block-container {
        max-width: 1150px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    .app-subtitle {
        color: #777777;
        font-size: 16px;
    }

    /* Improve the Streamlit interface on narrow phone screens. */
    @media (max-width: 640px) {
        .block-container {
            padding-left: 0.8rem;
            padding-right: 0.8rem;
            padding-top: 1rem;
        }

        div.stButton > button,
        div.stDownloadButton > button {
            min-height: 48px;
            white-space: normal;
        }

        [data-testid="stFileUploader"] {
            width: 100%;
        }
    }

    div.stButton > button {
        border-radius: 8px;
        min-height: 42px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# 3. DATABASE
# ============================================================

def get_connection():

    connection = sqlite3.connect(DB_PATH)

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
                status TEXT NOT NULL,
                created_at TEXT NOT NULL
            )
            """
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

    calculated_hash, _ = hash_password(
        password,
        salt
    )

    return secrets.compare_digest(
        calculated_hash,
        stored_hash
    )


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

    if verify_password(
        password,
        user["password_hash"],
        user["password_salt"]
    ):

        return {
            "id": user["id"],
            "full_name": user["full_name"],
            "username": user["username"]
        }

    return None


# ============================================================
# 6. ACTIVITY HISTORY
# ============================================================

def log_activity(filename, file_type, text):

    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")

    if not username:
        return

    with get_connection() as conn:

        conn.execute(
            """
            INSERT INTO activity (
                username,
                filename,
                file_type,
                word_count,
                character_count,
                status,
                created_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (
                username,
                filename,
                file_type,
                len(text.split()),
                len(text),
                "Completed",
                datetime.now().isoformat(timespec="seconds")
            )
        )


def get_activity():

    if st.session_state.get("demo_mode"):
        return pd.DataFrame()

    username = st.session_state.get("username")

    with get_connection() as conn:

        rows = conn.execute(
            """
            SELECT filename, file_type, word_count,
                   character_count, status, created_at
            FROM activity
            WHERE username = ?
            ORDER BY id DESC
            """,
            (username,)
        ).fetchall()

    return pd.DataFrame([
        dict(row) for row in rows
    ])


def delete_activity():

    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")

    with get_connection() as conn:

        conn.execute(
            "DELETE FROM activity WHERE username = ?",
            (username,)
        )


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

    "editable_result": "",

    "auth_page": "login"
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

    st.session_state.editable_result = ""


def logout():

    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.full_name = None
    st.session_state.demo_mode = False

    clear_current_result()

    st.rerun()


# ============================================================
# 8. LOAD OCR ENGINE
# ============================================================

@st.cache_resource(show_spinner="Loading text recognition engine...")
def load_ocr_reader():

    return easyocr.Reader(
        ["en"],
        gpu=False
    )


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
# 10. IMAGE ENHANCEMENT
# ============================================================

def prepare_image(image):
    """Correct orientation and validate dimensions without damaging text."""
    image = ImageOps.exif_transpose(image)
    if image.width * image.height > MAX_IMAGE_PIXELS:
        raise ValueError("Image is too large to process safely.")
    return image.convert("RGB")


def crop_main_document(image):
    """Apply a conservative crop for wide desktop screenshots.

    This is only an approximation, not document-boundary detection. Users
    should compare it with Read Entire Image because a crop can omit content.
    """
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
    """Create a separate high-contrast copy while preserving the source image."""
    image = prepare_image(image)
    width, height = image.size

    # Enlarge small images so printed characters occupy more pixels.
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
    """Group detected text boxes into visual lines, then read left to right."""
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
    """Extract text using the selected reading area and image preparation."""
    prepared = prepare_image(image)

    if mode == "Read Main Document":
        prepared = crop_main_document(prepared)

    if quality_mode == "Enhanced":
        prepared = enhance_for_ocr(prepared)

    reader = load_ocr_reader()
    # Performance-focused OCR settings:
    # - greedy decoding is much faster than beamsearch on CPU.
    # - automatic rotation testing runs OCR several extra times and is
    #   expensive on Streamlit Cloud, so users should rotate sideways photos
    #   before uploading them.
    # - a smaller canvas reduces CPU work while retaining useful resolution
    #   for ordinary document photos and screenshots.
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
# 12. IMAGE PROCESSING
# ============================================================

def process_image(uploaded_file, mode="Read Entire Image", quality_mode="Standard"):

    file_bytes = check_file_size(uploaded_file)

    image = Image.open(
        io.BytesIO(file_bytes)
    )

    return image_to_text(image, mode=mode, quality_mode=quality_mode)


# ============================================================
# 13. PDF PROCESSING
# ============================================================

def process_pdf(uploaded_file, quality_mode="Standard"):

    file_bytes = check_file_size(uploaded_file)

    document = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    if document.page_count > MAX_PDF_PAGES:

        document.close()

        raise ValueError(
            f"This PDF contains more than {MAX_PDF_PAGES} pages."
        )

    extracted_pages = []

    ocr_confidences = []
    all_uncertain_words = []

    progress = st.progress(0)

    for page_number, page in enumerate(document):

        page_text = page.get_text().strip()

        # If the PDF page has little or no selectable text,
        # convert it into an image and use OCR.
        if len(page_text) < 15:

            pix = page.get_pixmap(
                matrix=fitz.Matrix(1.5, 1.5),
                alpha=False
            )

            image = Image.open(
                io.BytesIO(
                    pix.tobytes("png")
                )
            )

            ocr_result = recognize_image(image, quality_mode=quality_mode)

            page_text = ocr_result["text"]

            ocr_confidences.append(
                ocr_result["confidence"]
            )

            all_uncertain_words.extend(
                ocr_result["uncertain_words"]
            )

        extracted_pages.append(
            f"PAGE {page_number + 1}\n{page_text}"
        )

        progress.progress(
            (page_number + 1) / document.page_count
        )

    document.close()

    progress.empty()

    if ocr_confidences:

        st.session_state.last_ocr_confidence = round(
            sum(ocr_confidences) / len(ocr_confidences),
            1
        )

        st.session_state.last_uncertain_words = (
            all_uncertain_words
        )

    return "\n\n".join(
        extracted_pages
    ).strip()


# ============================================================
# 14. WORD DOCUMENT PROCESSING
# ============================================================

def process_docx(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    document = Document(
        io.BytesIO(file_bytes)
    )

    paragraphs = []

    for paragraph in document.paragraphs:

        if paragraph.text.strip():

            paragraphs.append(
                paragraph.text.strip()
            )

    for table in document.tables:

        for row in table.rows:

            row_text = [
                cell.text.strip()
                for cell in row.cells
                if cell.text.strip()
            ]

            if row_text:

                paragraphs.append(
                    " | ".join(row_text)
                )

    return "\n".join(paragraphs).strip()


# ============================================================
# 15. POWERPOINT PROCESSING
# ============================================================

def process_pptx(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    presentation = Presentation(
        io.BytesIO(file_bytes)
    )

    slides_text = []

    for index, slide in enumerate(presentation.slides):

        slide_parts = []

        for shape in slide.shapes:

            if hasattr(shape, "text") and shape.text.strip():

                slide_parts.append(
                    shape.text.strip()
                )

        slides_text.append(
            f"SLIDE {index + 1}\n"
            + "\n".join(slide_parts)
        )

    return "\n\n".join(slides_text).strip()


# ============================================================
# 16. EXCEL PROCESSING
# ============================================================

def process_xlsx(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    workbook = pd.ExcelFile(
        io.BytesIO(file_bytes)
    )

    output = []

    for sheet_name in workbook.sheet_names:

        dataframe = pd.read_excel(
            workbook,
            sheet_name=sheet_name
        )

        output.append(
            f"SHEET: {sheet_name}\n"
            + dataframe.fillna("").to_string(index=False)
        )

    return "\n\n".join(output).strip()


# ============================================================
# 17. CSV PROCESSING
# ============================================================

def process_csv(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    dataframe = pd.read_csv(
        io.BytesIO(file_bytes)
    )

    return dataframe.fillna("").to_string(index=False)


# ============================================================
# 18. TEXT FILE PROCESSING
# ============================================================

def process_txt(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    try:

        return file_bytes.decode(
            "utf-8"
        ).strip()

    except UnicodeDecodeError:

        return file_bytes.decode(
            "latin-1",
            errors="replace"
        ).strip()


# ============================================================
# 19. RTF PROCESSING
# ============================================================

def process_rtf(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    raw_text = file_bytes.decode(
        "utf-8",
        errors="replace"
    )

    return rtf_to_text(
        raw_text
    ).strip()


# ============================================================
# 20. EPUB PROCESSING
# ============================================================

def process_epub(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    book = epub.read_epub(
        io.BytesIO(file_bytes)
    )

    chapters = []

    for item in book.get_items():

        if item.get_type() == ITEM_DOCUMENT:

            soup = BeautifulSoup(
                item.get_content(),
                "html.parser"
            )

            chapter_text = soup.get_text(
                separator="\n"
            ).strip()

            if chapter_text:

                chapters.append(
                    chapter_text
                )

    return "\n\n".join(chapters).strip()


# ============================================================
# 21. DOCUMENT PROCESSING ROUTER
# ============================================================

def process_uploaded_document(
    uploaded_file,
    image_mode="Read Entire Image",
    quality_mode="Standard"
):

    extension = Path(
        uploaded_file.name
    ).suffix.lower()

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

        raise ValueError(
            f"The file format {extension} is not supported."
        )

    if extension in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff"}:
        return processor(uploaded_file, mode=image_mode, quality_mode=quality_mode)
    if extension == ".pdf":
        return processor(uploaded_file, quality_mode=quality_mode)
    return processor(uploaded_file)


# ============================================================
# 22. TEXT TO SPEECH
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
    """Split long text into paragraph/sentence-sized sections."""
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
    """Consent-based speech controls for short and long text."""
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
        key=f"{key_prefix}_speech_speed",
        help="Slower uses Google text-to-speech's slow speech setting."
    )
    with st.expander("Preview selected reading part"):
        st.write(sections[selected_part])

    st.warning(
        "Generating speech sends the selected text to Google's text-to-speech service. "
        "Do not submit private or sensitive information."
    )
    consent = st.checkbox(
        "I agree to send this text to Google to create speech.",
        key=f"{key_prefix}_tts_consent"
    )
    if st.button(
        "Generate speech for this part",
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


# ============================================================
# 22B. OFFLINE READER EXPORT
# ============================================================

def build_offline_reader_html(text, document_title="NaijaSight Offline Reader"):
    """Build a self-contained HTML reader that can be saved and opened offline."""
    safe_title = html.escape(str(document_title or "NaijaSight Offline Reader"))
    # Prevent user text from terminating the embedded script element.
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
<meta name="theme-color" content="#172554">
<title>{safe_title}</title>
<style>
:root {{ color-scheme: light; --bg:#f4f7fb; --panel:#fff; --ink:#172033; --accent:#174ea6; --muted:#526078; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:var(--bg); color:var(--ink); font-family:Arial, sans-serif; line-height:1.65; }}
header {{ background:#172554; color:white; padding:20px 16px; }}
header h1 {{ margin:0 auto; max-width:900px; font-size:1.45rem; }}
main {{ max-width:900px; margin:20px auto; padding:0 14px 40px; }}
.panel {{ background:var(--panel); border:1px solid #d8e0ec; border-radius:14px; padding:16px; margin-bottom:16px; }}
.controls {{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; }}
button, select {{ min-height:44px; border-radius:9px; border:1px solid #9aa9bd; padding:9px 13px; font-size:1rem; background:white; color:#172033; }}
button.primary {{ background:var(--accent); color:white; border-color:var(--accent); }}
button:focus, select:focus, input:focus {{ outline:3px solid #f0b429; outline-offset:2px; }}
label {{ display:block; font-weight:600; margin:8px 0 4px; }}
#readerText {{ white-space:pre-wrap; overflow-wrap:anywhere; font-size:1.2rem; line-height:1.8; }}
#status {{ color:var(--muted); min-height:1.5em; }}
.small {{ color:var(--muted); font-size:.92rem; }}
body.high-contrast {{ --bg:#000; --panel:#000; --ink:#fff; --accent:#ffdf00; --muted:#eee; color:#fff; }}
body.high-contrast header {{ background:#000; border-bottom:2px solid #ffdf00; }}
body.high-contrast button, body.high-contrast select {{ background:#000; color:#fff; border:2px solid #ffdf00; }}
body.high-contrast button.primary {{ background:#ffdf00; color:#000; }}
@media (max-width:600px) {{
  header h1 {{ font-size:1.2rem; }}
  main {{ margin:12px auto; padding:0 10px 28px; }}
  .panel {{ padding:13px; }}
  .controls button, .controls select {{ flex:1 1 42%; }}
  #readerText {{ font-size:1.12rem; }}
}}
@media (prefers-reduced-motion: reduce) {{ * {{ scroll-behavior:auto !important; }} }}
</style>
</head>
<body>
<header><h1>NaijaSight — Offline Reader</h1></header>
<main>
  <section class="panel" aria-label="Reading controls">
    <p class="small">This file contains your saved text. After downloading it, you can reopen it without an internet connection. Read-aloud uses voices available on this device and browser.</p>
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
  <p class="small">Privacy: this standalone file does not send its saved text to NaijaSight. Browser speech availability and offline operation depend on the device's installed voices and browser.</p>
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
      if (sentence.length > 3000) {{
        const words = sentence.split(/\\s+/);
        for (const word of words) {{
          if (current && current.length + word.length + 1 > 3000) {{
            output.push(current); current = '';
          }}
          current = current ? current + ' ' + word : word;
        }}
      }} else {{
        current = current ? current + ' ' + sentence : sentence;
      }}
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
    statusLine.textContent = 'Read-aloud is not supported by this browser. You can still read the saved text.';
    return;
  }}
  stopSpeech();
  const utterance = new SpeechSynthesisUtterance(parts[Number(partSelect.value || 0)]);
  utterance.onstart = () => statusLine.textContent = 'Reading section aloud.';
  utterance.onend = () => statusLine.textContent = 'Finished reading this section.';
  utterance.onerror = () => statusLine.textContent = 'The device could not read this section aloud. Try another browser or device voice.';
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

        raise RuntimeError(
            "Speech recognition is not installed."
        )

    recognizer = sr.Recognizer()

    audio_bytes = audio_file.getvalue()

    with sr.AudioFile(
        io.BytesIO(audio_bytes)
    ) as source:

        audio_data = recognizer.record(source)

    return recognizer.recognize_google(
        audio_data,
        language="en-NG"
    )


# ============================================================
# 24. AUTHENTICATION PAGE
# ============================================================

def render_auth_portal():

    st.title("👁️ NaijaSight")

    st.markdown(
        """
        <p class="app-subtitle">
        A reading assistant for people who are blind
        or have low vision.
        </p>
        """,
        unsafe_allow_html=True
    )

    login_tab, register_tab = st.tabs(
        ["Log in", "Create account"]
    )

    with login_tab:

        with st.form("login_form"):

            st.subheader("Welcome back")

            username = st.text_input(
                "Username"
            )

            password = st.text_input(
                "Password",
                type="password"
            )

            submitted = st.form_submit_button(
                "Log in",
                use_container_width=True
            )

            if submitted:

                user = authenticate_user(
                    username,
                    password
                )

                if user:

                    st.session_state.authenticated = True
                    st.session_state.username = user["username"]
                    st.session_state.full_name = user["full_name"]
                    st.session_state.demo_mode = False

                    st.rerun()

                else:

                    st.error(
                        "The username or password is incorrect."
                    )

    with register_tab:

        with st.form("register_form"):

            st.subheader("Create your account")

            full_name = st.text_input(
                "Full name"
            )

            new_username = st.text_input(
                "Choose a username"
            )

            new_password = st.text_input(
                "Create a password",
                type="password"
            )

            confirm_password = st.text_input(
                "Confirm password",
                type="password"
            )

            submitted = st.form_submit_button(
                "Create account",
                use_container_width=True
            )

            if submitted:

                if new_password != confirm_password:

                    st.error(
                        "Your passwords do not match."
                    )

                else:

                    success, message = create_account(
                        full_name,
                        new_username,
                        new_password
                    )

                    if success:

                        st.success(message)
                        st.info("You can now log in.")

                    else:

                        st.error(message)

    st.divider()

    st.caption(
        "You can explore the application using demonstration mode. "
        "Activity is not saved in this mode."
    )

    if st.button(
        "Continue with demo",
        use_container_width=True
    ):

        st.session_state.authenticated = True
        st.session_state.username = "demo"
        st.session_state.full_name = "Demo User"
        st.session_state.demo_mode = True

        st.rerun()


# ============================================================
# 25. SIDEBAR
# ============================================================

def render_sidebar():

    with st.sidebar:

        st.title("👁️ NaijaSight")

        st.caption("Reading assistant")

        st.divider()

        st.write(
            f"Signed in as **{st.session_state.full_name}**"
        )

        if st.session_state.demo_mode:

            st.info("Demonstration mode")

        pages = [
            "Home", "Read a document", "Scan with camera",
            "Listen to text", "Voice tools", "Activity history"
        ]
        page = st.radio(
            "Navigation",
            pages,
            key="nav_page"
        )

        st.divider()

        st.caption(
            f"Maximum upload size: {MAX_FILE_SIZE_MB} MB"
        )

        if st.button(
            "Log out",
            use_container_width=True
        ):

            logout()

    return page


# ============================================================
# 26. HOME PAGE
# ============================================================

def go_to_page(page):
    """Change the navigation radio value before Streamlit reruns the app."""
    st.session_state["nav_page"] = page


def render_home():
    st.title("Welcome to NaijaSight")
    st.markdown(
        "<p class='app-subtitle'>A simple reading assistant for people who are blind or have low vision.</p>",
        unsafe_allow_html=True,
    )
    st.subheader("What would you like to do?")
    st.caption("Choose an action to open its workspace.")

    cards = [
        ("📄", "Read a document", "Extract text from PDF, Word, PowerPoint, Excel, EPUB and more.", "Read a document"),
        ("📷", "Scan an image", "Upload an image or use your camera to recognise printed text.", "Scan with camera"),
        ("🔊", "Listen to text", "Paste text or listen to text you have already extracted.", "Listen to text"),
        ("🎙️", "Voice tools", "Record speech and convert it into text.", "Voice tools"),
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
    st.subheader("Getting started")
    st.write("1. Choose a document or image.  2. Extract the text.  3. Review the result.  4. Listen to it or download it.")
    st.info("For clearer recognition, keep the text in focus and avoid glare. Always check important names, dates and amounts.")


# ============================================================
# 27. RESULT WORKSPACE
# ============================================================

def render_result_workspace():

    text = st.session_state.extracted_text

    if not text:

        st.info(
            "Your extracted text will appear here after processing."
        )

        return

    st.subheader("Extracted text")
    st.caption("OCR confidence is an estimate from detected text regions, not a guarantee that every word is correct.")

    # OCR confidence display
    confidence = st.session_state.get(
        "last_ocr_confidence"
    )

    if confidence is not None:

        st.metric(
            "OCR confidence",
            f"{confidence}%"
        )

        if confidence < 60:

            st.warning(
                "The OCR engine is uncertain about some of the text. "
                "Please review the result carefully."
            )

        elif confidence < 80:

            st.info(
                "Some parts of the text may need checking."
            )

        else:

            st.success(
                "The OCR engine recognised the text with relatively "
                "high confidence."
            )

        uncertain = st.session_state.get(
            "last_uncertain_words",
            []
        )

        if uncertain:

            with st.expander(
                "Review uncertain text"
            ):

                for item in uncertain:

                    st.write(
                        f"**{item['text']}** — "
                        f"{item['confidence']}% confidence"
                    )

    st.caption(
        f"{len(text.split())} words · {len(text)} characters"
    )

    # Keep the editor synchronised with newly extracted text.
    if st.session_state.get("editable_result") != text:

        st.session_state.editable_result = text

    edited_text = st.text_area(
        "Review or edit the extracted text",
        height=300,
        key="editable_result"
    )

    st.session_state.extracted_text = edited_text

    col1, col2 = st.columns(2)

    with col1:
        st.download_button(
            "Download text",
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
            "Download offline reader",
            data=offline_html.encode("utf-8"),
            file_name="naijasight_offline_reader.html",
            mime="text/html",
            use_container_width=True,
            help="Save this self-contained reader to your phone or computer and open it later without internet."
        )

    if st.button(
        "Clear result",
        use_container_width=True
    ):
        clear_current_result()
        st.rerun()

    st.divider()
    st.subheader("Listen to the text")
    render_speech_controls(edited_text, "result")


# ============================================================
# 28. DOCUMENT READER PAGE
# ============================================================

def render_document_reader():

    st.title("Read a document")

    st.write(
        "Upload a document or image. NaijaSight will try to "
        "extract its text."
    )

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

        st.write(
            f"Selected file: **{uploaded_file.name}**"
        )
        selected_extension = Path(uploaded_file.name).suffix.lower()

        if selected_extension in image_extensions:
            image_mode = st.radio(
                "Choose how NaijaSight should read this image",
                ["Read Entire Image", "Read Main Document"],
                horizontal=True,
                help="Read Entire Image includes all visible text. Read Main Document applies an approximate crop to wide desktop screenshots; it can miss content."
            )

        if selected_extension in image_extensions or selected_extension == ".pdf":
            quality_mode = st.radio(
                "OCR image preparation",
                ["Standard", "Enhanced"],
                horizontal=True,
                help="Standard uses the original image. Enhanced creates a grayscale, higher-contrast and slightly enlarged copy. Enhanced may help faint text but can also make some images worse."
            )

        st.caption(
            f"File size: {uploaded_file.size / (1024 * 1024):.2f} MB"
        )

        if selected_extension in image_extensions:
            st.info(
                "For faster scanning, use a clear, upright image. "
                "Rotate sideways photos before uploading. The faster OCR "
                "settings avoid repeated rotation and beam-search passes."
            )

        if st.button(
            "Extract text",
            type="primary"
        ):

            clear_current_result()

            try:

                with st.spinner(
                    "Preparing the OCR engine and reading your image. "
                    "The first scan after a server restart can take longer..."
                ):

                    extracted_text = process_uploaded_document(
                        uploaded_file,
                        image_mode=image_mode,
                        quality_mode=quality_mode
                    )

                if extracted_text:

                    st.session_state.extracted_text = extracted_text

                    st.session_state.current_filename = (
                        uploaded_file.name
                    )

                    st.session_state.current_file_type = (
                        Path(uploaded_file.name).suffix.lower()
                    )

                    st.session_state.editable_result = extracted_text

                    log_activity(
                        uploaded_file.name,
                        Path(uploaded_file.name).suffix.lower(),
                        extracted_text
                    )

                    st.success(
                        "Text extraction completed."
                    )

                else:

                    st.warning(
                        "No readable text was found. Try a clearer image "
                        "or check whether the document contains text."
                    )

            except Exception as error:

                st.error(
                    f"Something went wrong while reading the file: {error}"
                )

    render_result_workspace()


# ============================================================
# 29. CAMERA SCANNER
# ============================================================

def render_camera_scanner():

    st.title("Scan with camera")

    st.write(
        "Take a picture of a page, notice, book or printed document."
    )

    camera_mode = st.radio(
        "Choose reading mode",
        ["Read Entire Image", "Read Main Document"],
        horizontal=True,
    )
    camera_quality = st.radio(
        "OCR image preparation",
        ["Standard", "Enhanced"],
        horizontal=True,
        help="Enhanced creates a grayscale, higher-contrast and slightly enlarged copy. Compare both methods if the result is unclear."
    )
    camera_image = st.camera_input(
        "Take a picture"
    )

    if camera_image:

        if st.button(
            "Read this picture",
            type="primary"
        ):

            clear_current_result()

            try:

                with st.spinner(
                    "Reading the picture..."
                ):

                    image = Image.open(
                        camera_image
                    )

                    extracted_text = image_to_text(
                        image,
                        mode=camera_mode,
                        quality_mode=camera_quality
                    )

                if extracted_text:

                    st.session_state.extracted_text = extracted_text

                    st.session_state.current_filename = "Camera scan"
                    st.session_state.current_file_type = "camera"

                    st.session_state.editable_result = extracted_text

                    log_activity(
                        "Camera scan",
                        "camera",
                        extracted_text
                    )

                    st.success(
                        "Text extracted."
                    )

                else:

                    st.warning(
                        "No text was detected. Try taking another picture."
                    )

            except Exception as error:

                st.error(
                    f"Could not process the picture: {error}"
                )

    render_result_workspace()


# ============================================================
# 30. LISTEN TO TEXT
# ============================================================

def render_listen_page():
    st.title("Listen to text")
    st.write(
        "Paste text below, or use text you have already extracted with NaijaSight. "
        "Long text is divided into smaller reading parts."
    )
    initial_text = st.session_state.get("extracted_text", "")
    text_to_read = st.text_area(
        "Text to read aloud",
        value=initial_text,
        height=260,
        key="standalone_listen_text"
    )
    render_speech_controls(text_to_read, "standalone")


# ============================================================
# 31. VOICE TOOLS
# ============================================================

def render_voice_tools():

    st.title("Voice tools")

    st.write(
        "Record a short voice message and convert it into text."
    )

    if not SPEECH_RECOGNITION_AVAILABLE:

        st.error(
            "The speech recognition package is not installed."
        )

        return

    st.warning(
        "Voice recognition sends your recording to Google's "
        "speech recognition service. Avoid recording private "
        "or sensitive information."
    )

    consent = st.checkbox(
        "I agree to send my recording to Google for transcription."
    )

    audio_file = st.audio_input(
        "Record your voice"
    )

    if audio_file:

        st.audio(
            audio_file
        )

        if st.button(
            "Convert speech to text",
            disabled=not consent,
            type="primary"
        ):

            try:

                with st.spinner(
                    "Transcribing your recording..."
                ):

                    transcript = transcribe_audio(
                        audio_file
                    )

                st.session_state.extracted_text = transcript

                st.session_state.current_filename = "Voice recording"
                st.session_state.current_file_type = "audio"

                st.session_state.editable_result = transcript

                log_activity(
                    "Voice recording",
                    "audio",
                    transcript
                )

                st.success(
                    "Transcription completed."
                )

            except Exception as error:

                st.error(
                    f"Could not transcribe the recording: {error}"
                )

    render_result_workspace()


# ============================================================
# 32. ACTIVITY HISTORY
# ============================================================

def render_history():

    st.title("Activity history")

    if st.session_state.demo_mode:

        st.info(
            "History is not saved in demonstration mode."
        )

        return

    history = get_activity()

    if history.empty:

        st.info(
            "You have no saved activity yet."
        )

        return

    st.write(
        "This page stores basic activity details. "
        "It does not store the full text extracted from your files."
    )

    st.dataframe(
        history,
        use_container_width=True,
        hide_index=True
    )

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        history.to_excel(
            writer,
            index=False,
            sheet_name="Activity"
        )

    st.download_button(
        "Export history to Excel",
        data=output.getvalue(),
        file_name="naijasight_activity.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.divider()

    st.subheader("Delete activity history")

    st.warning(
        "This will permanently delete the activity records "
        "stored for your account."
    )

    confirm_delete = st.checkbox(
        "I understand that this cannot be undone."
    )

    if st.button(
        "Delete my history",
        disabled=not confirm_delete
    ):

        delete_activity()

        st.success(
            "Your activity history has been deleted."
        )

        st.rerun()


# ============================================================
# 33. MAIN APPLICATION
# ============================================================

def render_logged_in_portal():

    page = render_sidebar()

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

    st.divider()

    st.caption(
        "NaijaSight is a prototype. Text recognition and speech "
        "results may contain errors. Check important information."
    )


# ============================================================
# 34. APPLICATION ENTRY POINT
# ============================================================

if not st.session_state.authenticated:

    render_auth_portal()

else:

    render_logged_in_portal()
