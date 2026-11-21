
import easyocr
import pyttsx3
from pathlib import Path

# Project settings
IMAGE_FOLDER = Path("images")
OUTPUT_FILE = Path("extracted_text.txt")
SUPPORTED_FORMATS = [".png", ".jpg", ".jpeg"]

print("\n================================")
print("       WELCOME TO NAIJASIGHT")
print("================================")

print("\nLoading AI text recognition...")
reader = easyocr.Reader(['en'], gpu=False)

speaker = pyttsx3.init()
speaker.setProperty("rate", 150)

print("NaijaSight is ready!")


def read_document():
    filename = input("\nEnter image filename: ").strip().strip('"')

    image_path = Path(filename)

    if not image_path.is_absolute():
        image_path = IMAGE_FOLDER / image_path

    if not image_path.exists():
        print("Error: Image not found. Check the filename.")
        return

    if image_path.suffix.lower() not in SUPPORTED_FORMATS:
        print("Error: Unsupported image format.")
        return

    try:
        print("\nScanning document...")

        results = reader.readtext(str(image_path))

        full_text = "\n".join(
            result[1] for result in results
        )

        if not full_text.strip():
            print("No readable text found in this image.")
            return

        print("\n========== EXTRACTED TEXT ==========\n")
        print(full_text)

        # Save extracted text
        with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
            file.write(full_text)

        print("\nText saved to extracted_text.txt")

        # Speak extracted text
        print("\nNaijaSight is speaking...")
        speaker.say(full_text)
        speaker.runAndWait()

        print("\nFinished reading successfully!")

    except Exception as error:
        print("\nSomething went wrong:", error)


# Main application menu
while True:
    print("\n========== MAIN MENU ==========")
    print("1. Read a document")
    print("2. Exit")

    choice = input("\nChoose an option: ").strip()

    if choice == "1":
        read_document()

    elif choice == "2":
        print("\nThank you for using NaijaSight!")
        break

    else:
        print("\nInvalid choice. Please select 1 or 2.")
