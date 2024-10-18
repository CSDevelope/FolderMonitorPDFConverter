# Folder Monitor PDF Converter

This is a simple Python application that monitors a folder for new files and automatically converts supported files to PDF.

## Features
- Monitors a folder for new files.
- Converts `.docx`, `.png`, `.jpg`, `.jpeg`, `.xls`, `.xlsx` files to PDF.
- Deletes unsupported files.
- Customizable folders for monitoring and output.

## Prerequisites
- Python 3.8 or higher
- pip (Python package manager)

## Setup Instructions

### 1. Clone the Repository
Clone the repository to your local machine:
```bash
git clone https://github.com/CSDevelope/FolderMonitorPDFConverter.git
```

### 2. Navigate to the Project Directory
```bash
cd FolderMonitorPDFConverter
```

### 3. Create a Virtual Environment (Recommended)
```bash
python -m venv venv
```

Activate the virtual environment:

- On Windows:
  ```bash
  venv\Scripts\activate
  ```
- On macOS/Linux:
  ```bash
  source venv/bin/activate
  ```

### 4. Install Dependencies
Install all necessary dependencies:
```bash
pip install -r requirements.txt
```

### 5. Run the Application
Run the Python script:
```bash
python folder_monitor_pdf_converter.py
```

The GUI window will open, allowing you to select the monitored folder and output folder.

## Building the Application as an Executable
If you want to share the application as an executable, you can use `PyInstaller` to bundle it:

1. **Install PyInstaller**:
   ```sh
   pip install pyinstaller
   ```
   
2. **Create the Executable**:
   ```sh
   pyinstaller --onefile --windowed --icon=images/dhl.ico folder_monitor_pdf_converter.py
   ```
   - **`--onefile`**: Packages everything into a single executable.
   - **`--windowed`**: Removes the terminal window (for GUI-based apps).
   - **`--icon`**: Adds an icon to the executable.

The executable will be created in the `dist` folder.

## Notes
- Make sure that the `DejaVuSans.ttf` font is available in the `fonts` folder, and that `f1.png` and `dhl.ico` are in the `images` folder.
- This application only works on Windows due to its reliance on `pywin32`.

## Contact
For any questions, contact [Andrew Tufarella](mailto:andrew.tufarella@dhl.com).

