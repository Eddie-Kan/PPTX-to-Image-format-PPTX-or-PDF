# PPTX-to-Image-format-PPTX-or-PDF

- [中文版](README_ZH-CN.md)
---

## Introduction

PPTX-to-Image-format-PPTX-or-PDF is a tool designed to convert PPTX files into high-resolution images and, as needed, merge these images into a PDF file or a new PPTX file.

## Features

- **pptx-to-image_PDF.py**: Exports each slide of a PPTX file as high-resolution PNG images and merges them into a single PDF file.
- **pptx-to-image_PPTX.py**: Exports each slide of a PPTX file as high-resolution PNG images and inserts them into a new PPTX file, with each image as a separate slide.

## Dependencies

- **Operating System**: Windows only
- **Python Version**: Python 3.x
- **Required Libraries**:
  - `pywin32`
  - `Pillow` (PIL)
- **Software Requirements**:
  - Microsoft Office PowerPoint

## Install Dependencies

Install the required Python libraries using the following command:

```Bash
pip install pywin32 Pillow
```

## Usage
- Ensure Microsoft Office PowerPoint is installed.
- Run the script:

1. For generating a PDF:
```Bash
python pptx-to-image_PDF.py
```

2. For generating an image-based PPTX:
```Bash
python pptx-to-image_PPTX.py
```
3. Follow the prompts to input the PPTX file path and desired resolution (DPI).

## Notes
The script only runs on Windows systems as it uses the Windows COM interface to control PowerPoint.
Make sure you have a compatible version of Microsoft Office PowerPoint installed.
Before running the script, please close any open PowerPoint applications to avoid potential conflicts.

## License
This project is licensed under the MIT License.
