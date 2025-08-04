# IS2 Tool
IS2 Tool is a PyQt-based desktop application for reviewing and renaming `.is2` files exported from Fluke IR cameras. It enables efficient processing of thermal image files, export of visible light photos, and proper handling of Windows file metadata such as creation timestamps.

This utility is especially useful for engineers, technicians, or commissioning teams who deal with thermal imaging files and require structured organization, preview, and export of relevant image data.

## 📚 Table of Contents

- [Features](#-features)
- [Getting Started](#-getting-started)
- [Requirements](#-requirements)
- [Excel Format for Locations](#-excel-format-for-locations)
- [Troubleshooting](#-troubleshooting)
- [Credits](#-credits)


## ✨ Features
📂 **Batch processing** of .is2 files (renaming, metadata updates, and export)

🔎 **Image preview:** infrared, visible light, and up to 3 photo notes per file

🧠 **Smart renaming** using multi-tiered location metadata

📸 **Export** visible thumbnails for documentation or reports

🕒 **Correct** Windows 'Date Created' to match image capture date (NTFS only)

🔍 **Zoom & scroll** viewer for full-resolution image inspection

📥 **Excel import** to load predefined Procore Location hierarchies

🌗 **Dark-themed UI** with scalable layout

💾 **Automatically** cleans up extracted files after use

## 🚀 Getting Started
### 📦 Installation
1. Clone the repository:
```bash
git clone https://github.com/joshuahamsa/is2-tool.git
cd is2-tool
```

2. Create virtual environment (optional but recommended):
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

4. Run the app:
```bash
python is2tool.py
```

📌 **Note:** Make sure you're on Windows (NTFS filesystem) and have the .is2 files available for processing.

### 🧰 Requirements
- Python 3.8+
- Windows OS (required for metadata functions)
- Packages:
   - PyQt5
   - openpyxl
   - pywin32

Add these to your requirements.txt:
```txt
PyQt5>=5.15.0
openpyxl>=3.1.0
pywin32>=306
```

## 🖼️ Sample Workflow
1. Launch IS2 Tool
2. Click "Get Started" and select a folder containing .is2 files
3. Optionally import an Excel file for tiered location data
4. Use dropdowns and/or a custom suffix to rename the file
5. Preview and export visible images automatically
6. Proceed to the next file using Save & Next

All exported .jpg images will have their creation dates aligned with the original .is2 file's modified date for consistency.

## 📝 Excel Format for Locations
To use the tiered dropdowns for structured naming:

| Tier 1 | Tier 2	| Tier 3 |
| ------ | ------ | ------ |
| Site A	| Inverter 1	| String 3 |
| Site A	| Inverter 2	| String 1 |
| Site B	| Inverter 1	| - |

Each row defines a nested structure used to populate the tiered dropdowns for renaming.

## 🖥️ Packaging to EXE (Optional)
To create a standalone executable:
```bash
pip install pyinstaller
pyinstaller --onefile --noconsole is2_tool.py
```
### For better compression (optional):
```bash
pip install upx
pyinstaller --onefile --noconsole --upx-dir="path\to\upx" is2_tool.py
```

## 🛠 Troubleshooting
- **Date Created not updating?** Ensure the script is run with appropriate permissions on an NTFS volume.
- **No visible image found?** Some .is2 files may not include all expected JPEGs.
- **Zoom not working?** Double-click the image or use your mouse wheel to zoom in/out.

## 🙏 Credits
Created with ❤️ by Joshua Hamsa
Special thanks to Fluke and the commissioning teams who inspired this workflow.

