# Word to PDF Converter that doesn't steal your data (I swear)

Minimalist CLI tool for converting Word documents (.docx) to PDF while preserving quality, structure, and layout.

Runs on your machine, so no data is sent to any third party.

## Requirements

- Python 3.7+
- Microsoft Word installed (uses Word's native conversion engine for best quality)

## Installation

```bash
pip install -r requirements.txt
```

## Usage

> **Default output:** When no `-o` option is specified, the PDF is saved in the **same folder as the source file**, with the same name but a `.pdf` extension (e.g., `document.docx` → `document.pdf`).

```bash
# Convert a single file (saves as document.pdf in same folder)
python convert.py document.docx

# Specify output name
python convert.py document.docx -o output.pdf

# Convert all .docx files in a folder
python convert.py ./documents/

# Convert folder to a different output folder
python convert.py ./docs/ -o ./pdfs/

# Quiet mode (no output messages)
python convert.py document.docx -q
```

## Options

| Option | Description |
|--------|-------------|
| `input` | Input .docx file or directory |
| `-o, --output` | Output PDF file or directory |
| `-q, --quiet` | Suppress output messages |
| `-h, --help` | Show help message |
