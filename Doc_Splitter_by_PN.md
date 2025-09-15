# Documentation for PDF Splitter by Person Number

## Overview

This script splits large PDF files (like salary review letters) into individual PDFs per employee, based on detecting:

- The employee’s name (using "Dear ...")
- Their person number (from explicit "Person Number:" field or directly below the name)

It uses PyPDF2 for PDF processing and Tkinter for a simple GUI with progress reporting.

## Features

- Select multiple PDFs to process at once
- Extract employee first/last name and person number
- Save each employee’s PDF separately with structured filenames:

```py
	YYYY-MM_Salary Review_LASTNAME_Firstname_PersonNumber.pdf
```

- Handles continuation pages (multiple pages per employee
- GUI shows progress and automatically closes when finished

## Requirements
Install dependencies:

```bash
pip install PyPDF2
```
(Tkinter is included with standard Python on Windows/Mac; on Linux you may need to install python3-tk.)

## Usage

1. Run the script (python pdf_splitter.py)
2. Use the file dialog to select input PDF(s)
3. Choose an output folder
4. Wait while it processes — progress is shown in the GUI
5. Individual PDFs are saved into a timestamped subfolder

## Code Structure
### Utilities

- ``sanitize_filename(name)``

	Removes invalid characters for filenames. Keeps letters, digits, ``-_.()``, and spaces.

- ``page_verification(text)``

	Checks if a page is the start of a new letter by scanning for key phrases like ``"Dear"`` or ``"Chief People Officer"``.

- ``find_number_below_name(text, first_name, last_name, max_lines_below=2)``

	Looks for a numeric ID in the lines immediately below the employee’s name.

- ``extract_name(text, max_lines_above=8)``

	Extracts the first and last name:

	- Finds ``"Dear <FirstName>"``.

	- Searches up to 8 lines above to locate the last name.

### GUI (PDFSplitterGUI)

- ``__init__()`` → Initializes the Tkinter window.
- ``setup_gui()`` → Builds the UI (button, progress bar, labels).
- ``safe_update_gui(status, detail, progress)`` → Thread-safe updates for GUI labels/progress bar.
- ``start_processing()`` → Starts PDF processing in a background thread.

- ``process_pdfs()`` → Core logic:
	- Opens PDF(s).
	- Splits per employee.
	- Extracts name and person number.
	- Saves PDFs into a timestamped output folder.
	- Updates GUI.

### Main Entry Point
```py
if __name__ == "__main__":
    app = PDFSplitterGUI()
    app.root.mainloop()
```

Starts the GUI when the script is executed directly.

### Error Handling
- Exceptions are caught and shown in the console with a full traceback.
- The GUI shows a friendly error message without crashing.


### Example Workflow

Input PDF contains:

```nginx
Dear John,
John Doe
104
Human Resources
```

➡ Output file:

```yaml
2025-09_Salary Review_DOE_John_104.pdf
```