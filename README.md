# Amiga Wordworth Converter

A tool for converting Amiga Wordworth (IFF WORD/WOWO) files to modern .docx (Word) and .txt (plain text) formats.

---

## Project by Dana Burman
Vibe-coded with AI assistance (Probably Claude)

---

## Features
- Supports both Wordworth 2/3 (WORD) and Wordworth 4+ (WOWO) formats
- Converts to .docx (Word) and/or .txt (plain text)
- Batch conversion and drag-and-drop GUI (Tkinter)
- Preserves basic formatting, bullets, and headings

## Requirements
- Python 3.7+
- [python-docx](https://pypi.org/project/python-docx/) (`pip install python-docx`)
- [tkinterdnd2](https://pypi.org/project/tkinterdnd2/) (`pip install tkinterdnd2`) (optional, for drag-and-drop)

## Usage

### 1. Install dependencies
```
pip install python-docx
pip install tkinterdnd2   # optional, for drag-and-drop
```

### 2. Run the converter
```
python wordworth_converter.py
```

- Use the GUI to drag-and-drop files, or click to browse and select files.
- Choose output format(s): .docx, .txt, or both.
- Optionally, select a custom output directory.
- Click "Convert Files" to process the queue.

### 3. Command-line/Script Use
You can also import and use the conversion functions (`render_txt`, `render_docx`, `WordworthFile`) in your own Python scripts.

## OS Compatibility
- The converter is written in pure Python and should work on Windows, macOS, and Linux.
- No Amiga emulator or special hardware required.
- Only standard Python libraries and the above dependencies are used.


## License
This project is provided as-is, with no warranty. See source code for details.

---

## Third-Party Licenses

- **python-docx** © 2009–2026, MIT License
	- https://github.com/python-openxml/python-docx
- **tkinterdnd2** © Eliav2/pmgagne, Public Domain
	- https://github.com/Eliav2/tkinterdnd2

Both libraries are permissively licensed for commercial and non-commercial use.

---

## Credits
- Project by Dana Burman
- Vibe-coded with AI (Probably Claude)
