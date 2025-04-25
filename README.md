# 📄✂️ Document Extractor Pro

[![Python 3.8+](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Open Source Love](https://badges.frapsoft.com/os/v2/open-source.svg?v=103)](https://github.com/ellerbrock/open-source-badges/)

A modern GUI application for extracting pages from PDF and Word documents, with multiple export options and beautiful thumbnails preview.

![App Preview](screenshots/app-preview.gif) <!-- Add actual screenshot later -->

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
