
# NaijaSight

### An AI-Powered Accessibility Assistant

NaijaSight is an assistive technology project designed to help visually impaired individuals access printed information more independently.

## Current Features

- Optical Character Recognition (OCR) using EasyOCR.
- Text extraction from document images.
- Text-to-Speech using pyttsx3.
- Automatic saving of extracted text.
- Support for PNG, JPG and JPEG images.

## Technologies Used

- Python
- EasyOCR
- PyTorch
- pyttsx3

## Installation

1. Clone this repository.
2. Install the required packages:

   pip install -r requirements.txt

3. Create an images folder.
4. Place a document image inside the folder.
5. Run the application:

   python main.py

## Usage

Choose option 1 to read a document.

Enter the image filename, for example:

document.png

NaijaSight extracts the text, displays it, saves it and reads it aloud.

## Future Development

- Nigerian Naira denomination recognition.
- Mobile-friendly interface.
- Android camera integration.
- Voice-based interaction.
- Support for Nigerian languages.

## Project Status

Early MVP development.

## Responsible Development

NaijaSight aims to be developed with input from visually impaired users and accessibility testers.
