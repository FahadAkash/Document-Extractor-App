# 📄✂️ Document Extractor Pro

[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Open Source Love](https://badges.frapsoft.com/os/v2/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badges/)

A modern GUI application for extracting pages from PDF and Word documents, with multiple export options and beautiful thumbnails preview.

![App Preview]([https://raw.githubusercontent.com/FahadAkash/Document-Extractor-App/refs/heads/main/view/appview.png)) <!-- Add actual screenshot later -->

## ✨ Features

- 🖼️ Visual PDF/DOCX preview with hover zoom
- 📑 Extract pages as:
  - Combined image (PNG)
  - Separate images
  - New PDF file
- 📝 DOC/DOCX to PDF conversion
- 🎯 Range selection (e.g., "1,3-5,7")
- 🖱️ Drag & drop support
- 🌈 Modern UI with emoji support
- 📊 Status bar notifications
- 🛠️ Error handling with user-friendly messages

## 🚀 Installation

### Prerequisites
- Python 3.8+
- [LibreOffice](https://www.libreoffice.org/) (for DOC/DOCX conversion)

### Recommended: Virtual Environment
```bash
# Create virtual environment
python -m venv venv

# Activate environment
# Windows:
venv\Scripts\activate
# Linux/macOS:
source venv/bin/activate
```
# Install Dependencies
```bash
pip install -r requirements.txt
```
System Requirements
Component	Minimum Version
LibreOffice	7.0
Python	3.8
RAM	2 GB
Disk Space	500 MB
## 🛠️ Building from Source
```bash
# Clone repository
git clone https://github.com/yourusername/document-extractor.git
cd document-extractor

# Set up environment
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\\Scripts\\activate     # Windows

# Install dependencies
pip install -r requirements.txt

# Run application
python main.py
```
## ⚠️ Troubleshooting
```bash
# Verify LibreOffice installation
soffice --version

# Add to PATH if needed (Linux/macOS)
export PATH=$PATH:/usr/lib/libreoffice/program
```


