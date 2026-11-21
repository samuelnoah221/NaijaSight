
import streamlit as st
import easyocr
import fitz
import numpy as np
import pandas as pd
import io
import os
import re
import csv
import sqlite3
import hashlib
import secrets
import html
import time

from pathlib import Path
from datetime import datetime
from PIL import Image, ImageOps
from gtts import gTTS

# Optional document readers
from docx import Document
from pptx import Presentation
from striprtf.striprtf import rtf_to_text
from ebooklib import epub, ITEM_DOCUMENT
from bs4 import BeautifulSoup

# Optional voice recognition
try:
    import speech_recognition as sr
    SPEECH_RECOGNITION_AVAILABLE = True
except ImportError:
    SPEECH_RECOGNITION_AVAILABLE = False


# ============================================================
# PAGE SETTINGS
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
# SIMPLE STYLING
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

    .app-header {
        padding: 18px 0 12px 0;
    }

    .app-subtitle {
        color: #777777;
        font-size: 16px;
    }

    .small-note {
        color: #777777;
        font-size: 13px;
    }

    div.stButton > button {
        border-radius: 8px;
        min-height: 42px;
    }

    div[data-testid="stFileUploader"] {
        border-radius: 10px;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# DATABASE
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
# PASSWORD HANDLING
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

    return secrets.compare_digest(
        calculated_hash,
        stored_hash
    )


# ============================================================
# ACCOUNT FUNCTIONS
# ============================================================

def create_account(full_name, username, password):

    full_name = full_name.strip()
    username = username.strip().lower()

    if not full_name or not username or not password:
        return False, "Please complete all fields."

    if len(username) < 3:
        return False, "Your username must contain at least 3 characters."

    if not re.fullmatch(r"[a-zA-Z0-9_.-]+", username):
        return False, "Use only letters, numbers, dots, underscores or hyphens in your username."

    if len(password) < 8:
        return False, "Your password must contain at least 8 characters."

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
# ACTIVITY HISTORY
# ============================================================

def log_activity(filename, file_type, text, status="Completed"):

    if st.session_state.get("demo_mode"):
        return

    username = st.session_state.get("username")

    if not username:
        return

    word_count = len(text.split())
    character_count = len(text)

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
                word_count,
                character_count,
                status,
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

    return pd.DataFrame([dict(row) for row in rows])


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
# SESSION STATE
# ============================================================

defaults = {
    "authenticated": False,
    "username": None,
    "full_name": None,
    "demo_mode": False,
    "extracted_text": "",
    "current_filename": "",
    "current_file_type": "",
    "auth_page": "login"
}

for key, value in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = value


def logout():

    st.session_state.authenticated = False
    st.session_state.username = None
    st.session_state.full_name = None
    st.session_state.demo_mode = False
    st.session_state.extracted_text = ""
    st.session_state.current_filename = ""
    st.session_state.current_file_type = ""

    st.rerun()


# ============================================================
# EASY OCR
# ============================================================

@st.cache_resource(show_spinner="Loading text recognition...")
def load_ocr_reader():

    return easyocr.Reader(
        ["en"],
        gpu=False
    )


def check_file_size(uploaded_file):

    file_size = uploaded_file.size

    if file_size > MAX_FILE_SIZE_BYTES:
        raise ValueError(
            f"This file is larger than the {MAX_FILE_SIZE_MB} MB upload limit."
        )

    return uploaded_file.getvalue()


def image_to_text(image):

    if image.width * image.height > MAX_IMAGE_PIXELS:
        raise ValueError(
            "This image is too large to process safely."
        )

    image = ImageOps.exif_transpose(image)
    image = image.convert("RGB")

    reader = load_ocr_reader()

    image_array = np.array(image)

    results = reader.readtext(
        image_array,
        detail=0,
        paragraph=True
    )

    return "\n".join(results).strip()


# ============================================================
# IMAGE PROCESSING
# ============================================================

def process_image(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    image = Image.open(io.BytesIO(file_bytes))

    text = image_to_text(image)

    return text


# ============================================================
# PDF PROCESSING
# ============================================================

def process_pdf(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    document = fitz.open(
        stream=file_bytes,
        filetype="pdf"
    )

    if document.page_count > MAX_PDF_PAGES:
        raise ValueError(
            f"This PDF has {document.page_count} pages. "
            f"The current limit is {MAX_PDF_PAGES} pages."
        )

    extracted_pages = []

    progress = st.progress(0)

    for page_number, page in enumerate(document):

        page_text = page.get_text().strip()

        # Use OCR if the page has little or no selectable text.
        if len(page_text) < 15:

            pix = page.get_pixmap(
                matrix=fitz.Matrix(1.5, 1.5),
                alpha=False
            )

            image = Image.open(
                io.BytesIO(pix.tobytes("png"))
            )

            page_text = image_to_text(image)

        extracted_pages.append(
            f"PAGE {page_number + 1}\n{page_text}"
        )

        progress.progress(
            (page_number + 1) / document.page_count
        )

    document.close()
    progress.empty()

    return "\n\n".join(extracted_pages).strip()


# ============================================================
# WORD DOCUMENT
# ============================================================

def process_docx(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    document = Document(
        io.BytesIO(file_bytes)
    )

    paragraphs = [
        paragraph.text
        for paragraph in document.paragraphs
        if paragraph.text.strip()
    ]

    for table in document.tables:

        for row in table.rows:

            row_text = [
                cell.text.strip()
                for cell in row.cells
                if cell.text.strip()
            ]

            if row_text:
                paragraphs.append(" | ".join(row_text))

    return "\n".join(paragraphs).strip()


# ============================================================
# POWERPOINT
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
                slide_parts.append(shape.text.strip())

        slides_text.append(
            f"SLIDE {index + 1}\n" + "\n".join(slide_parts)
        )

    return "\n\n".join(slides_text).strip()


# ============================================================
# EXCEL
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
# CSV
# ============================================================

def process_csv(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    dataframe = pd.read_csv(
        io.BytesIO(file_bytes)
    )

    return dataframe.fillna("").to_string(index=False)


# ============================================================
# TEXT FILE
# ============================================================

def process_txt(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    try:
        return file_bytes.decode("utf-8").strip()

    except UnicodeDecodeError:
        return file_bytes.decode(
            "latin-1",
            errors="replace"
        ).strip()


# ============================================================
# RTF
# ============================================================

def process_rtf(uploaded_file):

    file_bytes = check_file_size(uploaded_file)

    raw_text = file_bytes.decode(
        "utf-8",
        errors="replace"
    )

    return rtf_to_text(raw_text).strip()


# ============================================================
# EPUB
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
                chapters.append(chapter_text)

    return "\n\n".join(chapters).strip()


# ============================================================
# DOCUMENT ROUTER
# ============================================================

def process_uploaded_document(uploaded_file):

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

    return processor(uploaded_file)


# ============================================================
# TEXT TO SPEECH
# ============================================================

def create_speech(text):

    if not text.strip():
        raise ValueError("There is no text to read.")

    audio_buffer = io.BytesIO()

    speech = gTTS(
        text=text,
        lang="en",
        slow=False
    )

    speech.write_to_fp(audio_buffer)

    audio_buffer.seek(0)

    return audio_buffer


# ============================================================
# SPEECH TO TEXT
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
# AUTHENTICATION PAGE
# ============================================================

def render_auth_portal():

    st.markdown(
        """
        <div class="app-header">
            <h1>👁️ NaijaSight</h1>
            <p class="app-subtitle">
                A reading assistant for people who are blind
                or have low vision.
            </p>
        </div>
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

                    st.success("You are now logged in.")
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

                    st.error("Your passwords do not match.")

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
        "You can also explore the app using the demonstration account. "
        "Demo activity is not saved."
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
# SIDEBAR
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

        page = st.radio(
            "Navigation",
            [
                "Home",
                "Read a document",
                "Scan with camera",
                "Voice tools",
                "Activity history"
            ]
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
# HOME PAGE
# ============================================================

def render_home():

    st.title("Welcome to NaijaSight")

    st.write(
        "Use NaijaSight to extract text from documents and images, "
        "edit the result, and listen to the text."
    )

    st.subheader("What would you like to do?")

    col1, col2 = st.columns(2)

    with col1:

        st.markdown("### 📄 Read a document")

        st.write(
            "Extract text from PDF, Word, PowerPoint, "
            "Excel, EPUB and other supported files."
        )

    with col2:

        st.markdown("### 📷 Scan an image")

        st.write(
            "Take a picture of printed text and convert it into "
            "editable text."
        )

    col3, col4 = st.columns(2)

    with col3:

        st.markdown("### 🔊 Listen to text")

        st.write(
            "Turn extracted text into spoken audio."
        )

    with col4:

        st.markdown("### 🎙️ Voice tools")

        st.write(
            "Record speech and convert it into text."
        )

    st.divider()

    st.info(
        "For best results, use clear images with good lighting "
        "and keep the text in focus."
    )


# ============================================================
# RESULT WORKSPACE
# ============================================================

def render_result_workspace():

    text = st.session_state.extracted_text

    if not text:
        st.info(
            "Your extracted text will appear here after processing."
        )
        return

    st.subheader("Extracted text")

    st.caption(
        f"{len(text.split())} words · {len(text)} characters"
    )

    edited_text = st.text_area(
        "Review or edit the extracted text",
        value=text,
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

        if st.button(
            "Clear result",
            use_container_width=True
        ):

            st.session_state.extracted_text = ""
            st.session_state.current_filename = ""
            st.rerun()

    st.divider()

    st.subheader("Listen to the text")

    st.warning(
        "Text-to-speech sends the text you choose to Google's "
        "text-to-speech service. Only continue if you agree."
    )

    allow_tts = st.checkbox(
        "I agree to send this text to Google to create speech.",
        key="tts_consent"
    )

    if st.button(
        "Generate audio",
        disabled=not allow_tts
    ):

        try:

            with st.spinner("Preparing audio..."):

                audio = create_speech(
                    edited_text
                )

            st.audio(
                audio,
                format="audio/mp3"
            )

            st.download_button(
                "Download audio",
                data=audio.getvalue(),
                file_name="naijasight_audio.mp3",
                mime="audio/mpeg"
            )

        except Exception as error:

            st.error(
                f"Could not create audio: {error}"
            )


# ============================================================
# DOCUMENT READER PAGE
# ============================================================

def render_document_reader():

    st.title("Read a document")

    st.write(
        "Upload a document or image. NaijaSight will try to "
        "extract the text for you."
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

    if uploaded_file:

        st.write(
            f"Selected file: **{uploaded_file.name}**"
        )

        st.caption(
            f"File size: {uploaded_file.size / (1024 * 1024):.2f} MB"
        )

        if st.button(
            "Extract text",
            type="primary"
        ):

            try:

                with st.spinner(
                    "Reading your file. This may take a while..."
                ):

                    extracted_text = process_uploaded_document(
                        uploaded_file
                    )

                if extracted_text:

                    st.session_state.extracted_text = extracted_text
                    st.session_state.current_filename = uploaded_file.name
                    st.session_state.current_file_type = Path(
                        uploaded_file.name
                    ).suffix.lower()

                    log_activity(
                        uploaded_file.name,
                        Path(uploaded_file.name).suffix.lower(),
                        extracted_text
                    )

                    st.success("Text extraction completed.")

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
# CAMERA SCANNER
# ============================================================

def render_camera_scanner():

    st.title("Scan with camera")

    st.write(
        "Use your device camera to take a picture of a page, "
        "notice, book or printed document."
    )

    camera_image = st.camera_input(
        "Take a picture"
    )

    if camera_image:

        if st.button(
            "Read this picture",
            type="primary"
        ):

            try:

                with st.spinner("Reading the picture..."):

                    image = Image.open(camera_image)

                    extracted_text = image_to_text(image)

                if extracted_text:

                    st.session_state.extracted_text = extracted_text
                    st.session_state.current_filename = "Camera scan"
                    st.session_state.current_file_type = "camera"

                    log_activity(
                        "Camera scan",
                        "camera",
                        extracted_text
                    )

                    st.success("Text extracted.")

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
# VOICE TOOLS
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
        "Voice recognition sends the recording to Google's "
        "speech recognition service. Do not record private or "
        "sensitive information."
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

                with st.spinner("Transcribing your recording..."):

                    transcript = transcribe_audio(
                        audio_file
                    )

                st.session_state.extracted_text = transcript
                st.session_state.current_filename = "Voice recording"
                st.session_state.current_file_type = "audio"

                log_activity(
                    "Voice recording",
                    "audio",
                    transcript
                )

                st.success("Transcription completed.")

            except Exception as error:

                st.error(
                    f"Could not transcribe the recording: {error}"
                )

    render_result_workspace()


# ============================================================
# ACTIVITY HISTORY
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

        st.success("Your activity history has been deleted.")
        st.rerun()


# ============================================================
# MAIN APPLICATION
# ============================================================

def render_logged_in_portal():

    page = render_sidebar()

    if page == "Home":

        render_home()

    elif page == "Read a document":

        render_document_reader()

    elif page == "Scan with camera":

        render_camera_scanner()

    elif page == "Voice tools":

        render_voice_tools()

    elif page == "Activity history":

        render_history()

    st.divider()

    st.caption(
        "NaijaSight is a prototype. Text recognition and speech "
        "results may contain errors, so check important information."
    )


# ============================================================
# APPLICATION ENTRY POINT
# ============================================================

if not st.session_state.authenticated:
    render_auth_portal()
else:
    render_logged_in_portal()
