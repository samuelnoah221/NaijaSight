import io
import zipfile
from pathlib import Path

import easyocr
import fitz
import numpy as np
import pandas as pd
import speech_recognition as sr
import streamlit as st
from bs4 import BeautifulSoup
from docx import Document
from ebooklib import ITEM_DOCUMENT, epub
from gtts import gTTS, gTTSError
from PIL import Image, UnidentifiedImageError
from pptx import Presentation
from striprtf.striprtf import rtf_to_text


# ==========================================
# 1. PAGE CONFIGURATION
# ==========================================

st.set_page_config(
    page_title="NaijaSight",
    page_icon="👁️",
    layout="centered"
)


# ==========================================
# 2. ACCESSIBILITY DESIGN
# ==========================================

st.markdown("""
<style>

html, body, [class*="st-"] {
    font-size: 18px;
}

h1, h2, h3 {
    color: #003366 !important;
    font-weight: 800 !important;
}

p, label, span {
    line-height: 1.6;
}

.stButton > button {
    min-height: 55px;
    font-size: 19px;
    font-weight: bold;
    border-radius: 10px;
}

[data-testid="stFileUploader"] {
    border: 2px solid #003366;
    border-radius: 12px;
    padding: 15px;
}

textarea {
    font-size: 18px !important;
    line-height: 1.6 !important;
}

*:focus-visible {
    outline: 4px solid #ff9900 !important;
    outline-offset: 3px !important;
}

</style>
""", unsafe_allow_html=True)


# ==========================================
# 3. APPLICATION INTRODUCTION
# ==========================================

st.title("👁️ NaijaSight")

st.subheader("Your AI-powered reading assistant")

st.write(
    "NaijaSight helps people with visual "
    "impairments access written information "
    "through text extraction and voice."
)

st.info(
    "Upload a document (up to 500 MB), take a picture, "
    "or use your microphone to give a command."
)


# ==========================================
# 4. SESSION STATE
# ==========================================

if "extracted_text" not in st.session_state:
    st.session_state["extracted_text"] = ""

if "editable_text" not in st.session_state:
    st.session_state["editable_text"] = ""

if "audio_bytes" not in st.session_state:
    st.session_state["audio_bytes"] = None

if "voice_message" not in st.session_state:
    st.session_state["voice_message"] = ""


def clear_audio():
    st.session_state["audio_bytes"] = None


# ==========================================
# 5. LOAD OCR ENGINE
# ==========================================

@st.cache_resource
def load_reader():
    return easyocr.Reader(["en"], gpu=False)


try:
    with st.spinner("Loading NaijaSight AI engine..."):
        reader = load_reader()
except Exception as engine_error:
    st.error(
        f"Failed to initialize the OCR engine. Please refresh the page or check server logs. Details: {engine_error}"
    )
    st.stop()


# ==========================================
# 6. SUPPORTED FILE FORMATS & LIMITS
# ==========================================

SUPPORTED_FILES = [
    "png", "jpg", "jpeg", "webp",
    "bmp", "tiff", "pdf",
    "docx", "txt", "rtf", "epub",
    "xlsx", "csv", "pptx"
]

MAX_FILE_SIZE_MB = 500
MAX_FILE_SIZE = MAX_FILE_SIZE_MB * 1024 * 1024


def validate_file_size(file_obj):
    if file_obj is None:
        raise ValueError("No file provided.")
    if file_obj.size > MAX_FILE_SIZE:
        size_mb = file_obj.size / (1024 * 1024)
        raise ValueError(
            f"File is too large ({size_mb:.1f} MB). Maximum allowed size is {MAX_FILE_SIZE_MB} MB."
        )


# ==========================================
# 7. IMAGE TEXT EXTRACTION
# ==========================================

def extract_image_text(image):
    try:
        image = image.convert("RGB")
        image_array = np.array(image)
        results = reader.readtext(image_array)
        return "\n".join(result[1] for result in results if len(result) > 1)
    except MemoryError:
        raise RuntimeError("The image resolution is too high to process in memory. Try a smaller image.")
    except Exception as e:
        raise RuntimeError(f"OCR image processing error: {e}")


# ==========================================
# 8. PDF EXTRACTION
# ==========================================

def extract_pdf(file_bytes):
    try:
        pdf = fitz.open(stream=file_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Unable to open PDF file. It may be corrupted. ({e})")

    if pdf.is_encrypted:
        pdf.close()
        raise ValueError("This PDF is password-protected. Please unlock it before uploading.")

    pages = []
    try:
        for number, page in enumerate(pdf, start=1):
            text = page.get_text().strip()

            # OCR scanned pages without selectable text
            if not text:
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                image = Image.open(io.BytesIO(pix.tobytes("png")))
                text = extract_image_text(image)

            pages.append(f"PAGE {number}\n{text}")
    except MemoryError:
        raise RuntimeError("Ran out of memory while reading this large PDF. Try splitting it into smaller parts.")
    finally:
        pdf.close()

    return "\n\n".join(pages)


# ==========================================
# 9. WORD DOCUMENT EXTRACTION
# ==========================================

def extract_word(file_bytes):
    try:
        document = Document(io.BytesIO(file_bytes))
    except (zipfile.BadZipFile, Exception) as e:
        raise ValueError(f"Could not read the Word document. Ensure it is a valid .docx file. ({e})")

    content = []
    for paragraph in document.paragraphs:
        if paragraph.text.strip():
            content.append(paragraph.text)

    for table in document.tables:
        for row in table.rows:
            row_text = " | ".join(cell.text.strip() for cell in row.cells)
            if row_text.strip(" |"):
                content.append(row_text)

    return "\n".join(content)


# ==========================================
# 10. EXCEL EXTRACTION
# ==========================================

def extract_excel(file_bytes):
    try:
        workbook = pd.ExcelFile(io.BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Could not open Excel file. It may be corrupted or password-protected. ({e})")

    content = []
    for sheet in workbook.sheet_names:
        try:
            dataframe = pd.read_excel(workbook, sheet_name=sheet)
            if not dataframe.empty:
                content.append(f"Sheet: {sheet}")
                content.append(dataframe.fillna("").to_string(index=False))
        except Exception as e:
            content.append(f"Sheet: {sheet} (Could not read sheet: {e})")

    return "\n\n".join(content)


# ==========================================
# 11. POWERPOINT EXTRACTION
# ==========================================

def extract_powerpoint(file_bytes):
    try:
        presentation = Presentation(io.BytesIO(file_bytes))
    except (zipfile.BadZipFile, Exception) as e:
        raise ValueError(f"Could not open PowerPoint file. Ensure it is a valid .pptx file. ({e})")

    slides = []
    for number, slide in enumerate(presentation.slides, start=1):
        slide_text = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text.strip():
                slide_text.append(shape.text.strip())
        if slide_text:
            slides.append(f"SLIDE {number}\n" + "\n".join(slide_text))

    return "\n\n".join(slides)


# ==========================================
# 12. EPUB EXTRACTION
# ==========================================

def extract_epub(file_bytes):
    try:
        book = epub.read_epub(io.BytesIO(file_bytes))
    except Exception as e:
        raise ValueError(f"Could not read EPUB e-book. It may be DRM-protected or corrupted. ({e})")

    chapters = []
    for item in book.get_items():
        if item.get_type() == ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_content(), "html.parser")
            text = soup.get_text(separator="\n", strip=True)
            if text:
                chapters.append(text)

    return "\n\n".join(chapters)


# ==========================================
# 13. UNIVERSAL FILE PROCESSOR
# ==========================================

def process_file(uploaded_file):
    validate_file_size(uploaded_file)

    extension = Path(uploaded_file.name).suffix.lower()
    file_bytes = uploaded_file.getvalue()

    if not file_bytes:
        raise ValueError("The uploaded file is empty (0 bytes).")

    if extension in [".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tiff"]:
        try:
            image = Image.open(io.BytesIO(file_bytes))
            image.verify()  # Check for image corruption
            image = Image.open(io.BytesIO(file_bytes))  # Re-open after verify()
        except (UnidentifiedImageError, Exception) as e:
            raise ValueError(f"Invalid or corrupted image file. ({e})")
        return extract_image_text(image)

    elif extension == ".pdf":
        return extract_pdf(file_bytes)

    elif extension == ".docx":
        return extract_word(file_bytes)

    elif extension == ".xlsx":
        return extract_excel(file_bytes)

    elif extension == ".pptx":
        return extract_powerpoint(file_bytes)

    elif extension == ".csv":
        try:
            dataframe = pd.read_csv(io.BytesIO(file_bytes))
            return dataframe.fillna("").to_string(index=False)
        except pd.errors.EmptyDataError:
            raise ValueError("The CSV file is empty.")
        except Exception as e:
            raise ValueError(f"Could not parse CSV file: {e}")

    elif extension == ".txt":
        return file_bytes.decode("utf-8-sig", errors="replace")

    elif extension == ".rtf":
        try:
            return rtf_to_text(file_bytes.decode("utf-8", errors="replace"))
        except Exception as e:
            raise ValueError(f"Could not parse RTF document: {e}")

    elif extension == ".epub":
        return extract_epub(file_bytes)

    else:
        raise ValueError(f"File format '{extension}' is not supported.")


# ==========================================
# 14. TEXT TO SPEECH ENGINE
# ==========================================

def generate_speech(text):
    clean_text = text.strip()
    if not clean_text:
        raise ValueError("There is no text to read.")

    try:
        speech = gTTS(text=clean_text, lang="en", slow=False)
        audio_buffer = io.BytesIO()
        speech.write_to_fp(audio_buffer)
        return audio_buffer.getvalue()
    except gTTSError as e:
        raise RuntimeError(
            f"Text-to-speech service error. Check your internet connection or try shortening very long text. ({e})"
        )
    except Exception as e:
        raise RuntimeError(f"Unexpected speech generation error: {e}")


# ==========================================
# 15. DOCUMENT UPLOAD
# ==========================================

st.divider()

st.header("1. Upload a document")

try:
    uploaded_file = st.file_uploader(
        "Choose a document from your device",
        type=SUPPORTED_FILES,
        max_upload_size=MAX_FILE_SIZE_MB,
        help=f"Maximum file size: {MAX_FILE_SIZE_MB} MB."
    )
except TypeError:
    # Fallback for older Streamlit versions (< 1.53) that rely on config.toml
    uploaded_file = st.file_uploader(
        "Choose a document from your device",
        type=SUPPORTED_FILES,
        help=f"Maximum file size: {MAX_FILE_SIZE_MB} MB (configure server.maxUploadSize=500 in .streamlit/config.toml)."
    )


# ==========================================
# 16. CAMERA CAPTURE
# ==========================================

st.header("2. Take a document picture")

camera_image = st.camera_input(
    "Use your device camera to capture a document"
)


# ==========================================
# 17. EXTRACT TEXT
# ==========================================

if st.button(
    "Extract document text",
    type="primary",
    use_container_width=True
):
    source_file = uploaded_file or camera_image

    if source_file is None:
        st.warning("Please upload a document or take a picture first.")
    else:
        with st.spinner("NaijaSight is reading your document..."):
            try:
                if camera_image is not None and uploaded_file is None:
                    validate_file_size(camera_image)
                    image = Image.open(camera_image)
                    full_text = extract_image_text(image)
                else:
                    full_text = process_file(uploaded_file)

                if full_text and full_text.strip():
                    st.session_state["extracted_text"] = full_text
                    st.session_state["editable_text"] = full_text
                    clear_audio()
                    st.success("Document text extracted successfully!")
                else:
                    st.warning("No readable text was found in this document.")

            except ValueError as val_err:
                st.warning(f"Document issue: {val_err}")
            except MemoryError:
                st.error("The file is too complex to process within available memory. Please try a smaller file.")
            except Exception as error:
                st.error(f"Extraction failed: {error}")


# ==========================================
# 18. DISPLAY EXTRACTED TEXT
# ==========================================

if st.session_state["extracted_text"]:
    st.divider()
    st.header("3. Extracted document")

    edited_text = st.text_area(
        "Review or correct the extracted text:",
        height=300,
        key="editable_text",
        on_change=clear_audio
    )

    st.download_button(
        "Download extracted text",
        data=edited_text,
        file_name="naijasight_text.txt",
        mime="text/plain",
        use_container_width=True
    )

    # ======================================
    # 19. TEXT TO SPEECH
    # ======================================

    st.divider()
    st.header("4. Listen to your document")

    if st.button(
        "Generate voice",
        type="primary",
        use_container_width=True
    ):
        try:
            with st.spinner("Generating speech..."):
                st.session_state["audio_bytes"] = generate_speech(edited_text)
            st.success("Your document is ready to listen to.")
        except ValueError as val_err:
            st.warning(str(val_err))
        except Exception as error:
            st.error(f"Speech generation failed: {error}")

    st.download_button(
        "Download corrected text",
        data=edited_text,
        file_name="naijasight_corrected_text.txt",
        mime="text/plain"
    )


# ==========================================
# 20. VOICE COMMAND SYSTEM
# ==========================================

st.divider()
st.header("5. Voice commands")

st.write(
    "Use your microphone to give NaijaSight "
    "a spoken instruction."
)

st.markdown("""
**Available commands:**

- Read document
- Repeat
- Stop
- Help
""")

voice_recording = st.audio_input(
    "Record your voice command",
    sample_rate=16000
)

if st.button(
    "Process voice command",
    use_container_width=True
):
    if voice_recording is None:
        st.warning("Please record a command first.")
    else:
        recognizer = sr.Recognizer()
        try:
            with st.spinner("Understanding your command..."):
                audio_bytes = voice_recording.getvalue()
                if not audio_bytes:
                    raise ValueError("The recorded audio was empty. Please try recording again.")

                audio_file = io.BytesIO(audio_bytes)
                with sr.AudioFile(audio_file) as source:
                    recorded_audio = recognizer.record(source)

                command = recognizer.recognize_google(
                    recorded_audio,
                    language="en-NG"
                ).lower().strip()

                st.session_state["voice_message"] = command

            st.write(f"**You said:** {command}")

            if "read" in command or "start reading" in command:
                text_to_read = st.session_state["editable_text"]
                if text_to_read.strip():
                    with st.spinner("Preparing your document audio..."):
                        st.session_state["audio_bytes"] = generate_speech(text_to_read)
                    st.success(
                        "Your document is ready. "
                        "Use the audio player below to listen."
                    )
                else:
                    st.warning("Please extract or upload a document first.")

            elif "repeat" in command:
                if st.session_state["audio_bytes"]:
                    st.success(
                        "Your previous audio is available below. "
                        "Press play to listen again."
                    )
                elif st.session_state["editable_text"].strip():
                    with st.spinner("Regenerating audio..."):
                        st.session_state["audio_bytes"] = generate_speech(
                            st.session_state["editable_text"]
                        )
                    st.success("Audio generated. Press play below.")
                else:
                    st.warning("There is no document to repeat.")

            elif "stop" in command:
                st.session_state["audio_bytes"] = None
                st.success("The audio player has been cleared.")

            elif "help" in command:
                st.info(
                    "Say 'Read document' to prepare speech, "
                    "'Repeat' to make the audio available again, "
                    "or 'Stop' to clear the audio player."
                )

            else:
                st.warning(
                    "Command not recognised. Say 'Help' "
                    "to hear the available commands."
                )

        except sr.UnknownValueError:
            st.error(
                "I could not understand the recording. "
                "Please speak clearly and try again."
            )
        except sr.RequestError:
            st.error(
                "Speech recognition service is unavailable. "
                "Check your internet connection."
            )
        except ValueError as val_err:
            st.warning(str(val_err))
        except Exception as error:
            st.error(f"Voice command error: {error}")


# ==========================================
# 21. AUDIO PLAYER
# ==========================================

if st.session_state["audio_bytes"]:
    st.divider()
    st.subheader("Your document audio")

    st.audio(
        st.session_state["audio_bytes"],
        format="audio/mp3"
    )

    st.download_button(
        "Download MP3 audio",
        data=st.session_state["audio_bytes"],
        file_name="naijasight_audio.mp3",
        mime="audio/mpeg",
        use_container_width=True
    )


# ==========================================
# 22. PRIVACY AND ACCURACY NOTICE
# ==========================================

st.divider()

st.caption("NaijaSight | Making information accessible.")

st.warning(
    "OCR can make mistakes. Always verify important information. "
    "Google speech recognition and text-to-speech require internet "
    "access and send the relevant audio or text to Google's service. "
    "Do not use confidential documents or recordings."
)

st.caption(
    "Camera and microphone access depend on browser permissions. "
    "Deployed versions should use HTTPS."
)