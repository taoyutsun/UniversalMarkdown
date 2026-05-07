# Universal Markdown Converter

[繁體中文](README.md) | [English](README.en.md)

A Windows desktop GUI utility built with **PyQt6** for converting common document formats into Markdown.

This project combines two mature conversion paths:

- `mammoth` for stable Word `.docx -> Markdown` conversion.
- Microsoft [markitdown](https://github.com/microsoft/markitdown) for multi-format document conversion and OCR / vision-assisted workflows.

Project repository: [taoyutsun/UniversalMarkdown](https://github.com/taoyutsun/UniversalMarkdown)

## Features

- Windows desktop GUI.
- Convert common document formats to Markdown.
- Preserve images from Word documents through the Mammoth route.
- Use MarkItDown for PDF, PowerPoint, Excel, and other supported formats.
- Optional OCR-enhanced conversion for image-heavy documents.
- Provider settings stored locally.
- API keys stored outside the repository through local settings or environment files.
- Portable packaging workflow for Windows users.

## Supported Formats And Routing

The application selects a conversion route based on file type and user settings.

Typical routes include:

- Word `.docx`: Mammoth-based conversion.
- PDF / PPTX / XLSX / XLS and other supported formats: MarkItDown route.
- OCR-enhanced workflow: optional route for scanned or image-heavy documents.

## Requirements

- Windows 10 or Windows 11.
- Python 3.10 or newer.
- A Python virtual environment is recommended.

Main dependencies include:

- `PyQt6`
- `mammoth`
- `openai`
- `python-dotenv`
- `markitdown[pdf,pptx,xlsx,xls]`
- `markitdown-ocr`

## Installation

Create and activate a virtual environment, then install dependencies:

```powershell
python -m pip install --upgrade pip
python -m pip install PyQt6 mammoth openai python-dotenv "markitdown[pdf,pptx,xlsx,xls]" markitdown-ocr
```

## Run

### Option A: Start With Python

```powershell
python .\app.py
```

### Option B: Start With The Batch File

```powershell
.\UniversalMarkdown.bat
```

## First-Time Setup

On first launch, review the provider settings. Depending on the conversion mode you use, you may need to configure:

- Local provider settings.
- API keys for OCR or vision-enhanced routes.
- Output folder preferences.

## Settings And Security

The project is designed so that private values are not committed to Git.

Typical local files include:

- `_UserSettings/settings.json`
- `.env`

Do not publish real API keys or personal provider credentials.

## DOCX Mode

For Word `.docx` files, the Mammoth route is used because it is stable and preserves practical Markdown structure. Images can be extracted and referenced in the generated Markdown output.

## OCR-Enhanced Mode

For scanned PDFs or image-heavy files, an OCR-enhanced route can be used if the required provider and dependencies are configured.

Use this mode only when the regular conversion route is not enough, because OCR/vision processing may be slower and may require cloud API usage.

## Portable Build

The repository includes build files for packaging a portable Windows version. Use the included batch or PyInstaller configuration as a starting point.

Example project files may include:

- `app.py`
- `converter.py`
- `UniversalMarkdown.bat`
- `Build_UniversalMarkdown_Exe.bat`
- `.env.example`
- `_UserSettings/`
- `dist/UniversalMarkdown_Portable/`

## Notes

This utility is intended to make documents easier to use with AI workflows by converting them into Markdown. It does not replace manual review, especially for scanned documents, tables, and complex layouts.

## License

See the repository license for details.
