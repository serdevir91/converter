Advanced File Converter
A modern, easy-to-use desktop application for converting a wide variety of file formats. Built with Python and the CustomTkinter library, this tool provides a clean and intuitive interface for all your document and image conversion needs.

📝 Description
This project was created to offer a simple yet powerful solution for common file conversion tasks. Instead of relying on web-based converters, this application runs locally on your machine, ensuring your files remain private and secure. The user interface is designed to be straightforward: just pick a conversion type, select your file, and get the output instantly.

✨ Features
The application supports a comprehensive range of conversion types:

Document Conversion
PDF to Word: Convert .pdf files to editable .docx documents.

Word to PDF: Convert .docx documents to .pdf files.

Image Conversion
PNG to JPG: Convert .png images to .jpg format.

JPG to PNG: Convert .jpg images to .png format.

Image to WEBP: Convert common image formats to the modern .webp format.

WEBP to PNG: Convert .webp images back to .png.

Image to ICO: Create .ico icon files from your images, with multiple sizes included.

Image to BMP: Convert images to the .bmp format.

HEIC to JPG: Convert iPhone photos (.heic) to the widely supported .jpg format.

Image to Grayscale: Convert any colored image to black and white.

🛠️ Requirements
To run this application from the source code, you'll need Python and the following libraries:

customtkinter

pdf2docx

docx2pdf

pypiwin32

Pillow

pillow-heif (Required for HEIC to JPG conversion)

🚀 Installation & Usage
Clone the repository:

git clone https://github.com/your-username/your-repository-name.git
cd your-repository-name

Install the required libraries:

pip install customtkinter pdf2docx docx2pdf pypiwin32 Pillow pillow-heif

Run the application:

python your_script_name.py

📦 Building the .EXE
You can create a standalone executable file for Windows using PyInstaller.

Install PyInstaller:

pip install pyinstaller

Run the build command:
(Make sure you have an icon.ico file in the same directory if you use the --icon option)

pyinstaller --name "FileConverter" --windowed --onefile --icon="icon.ico" "your_script_name.py"

The final executable will be located in the dist folder.
