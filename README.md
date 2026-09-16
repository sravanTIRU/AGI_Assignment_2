# Student Certificate Mail-Merge Automation Tool

REG NO: 261FB04042 | 1st Mtech, CSE

A Python-based certificate mail-merge automation tool that generates individual student certificates from an Excel file (`.xlsx`) and a DOCX certificate template (`.docx`), exporting layout-preserving PDF certificates into `output/certificates/` with the exact filename format `<REG_NO>.pdf`.

---

## 1. Quick Start Workflow

The tool is designed for zero-configuration, plug-and-play operation:

1. Place your Excel file (`students.xlsx`) and certificate template (`certificate_template.docx`) in any folder.
2. Open a terminal in that folder.
3. Run:
   ```bash
   python main.py
   ```
4. The tool automatically discovers the input files in the current working directory.
5. Pre-flight checks validate data integrity, required columns, template placeholders, and duplicate registration numbers.
6. Confirm generation (`Y/n`).
7. Collect your PDF certificates in:
   ```text
   output/certificates/
   ```
8. Every certificate is named strictly using only the student's registration number:
   ```text
   23CS001.pdf
   23CS002.pdf
   ...
   ```

---

## 2. Project Folder Structure

```text
Mail_Merge_Agent/
│
├── main.py                     # CLI entrypoint and argument parser
├── requirements.txt            # Python dependencies
├── README.md                   # Complete documentation and usage guide
│
├── app/                        # Modular application core
│   ├── __init__.py             # Package initializer
│   ├── excel_reader.py         # Auto-detection, column normalization, & row parser
│   ├── template_processor.py   # Run-level placeholder replacement & layout preservation
│   ├── validator.py            # Pre-flight data & duplicate validation
│   ├── pdf_converter.py        # Headless LibreOffice & MS Word COM converter
│   ├── certificate_generator.py# End-to-end orchestration & conflict protection
│   └── logger.py               # CSV audit logging (output/generation_log.csv)
│
├── tests/
│   └── test_suite.py           # Automated unit and integration tests
│
├── Student_Details.xlsx        # Sample/Active student details
├── certificate_template.docx   # Master certificate template
│
└── output/
    ├── generation_log.csv      # Status log of every student processed
    └── certificates/           # Generated PDF certificates (<REG_NO>.pdf)
        ├── 261FB04001.pdf
        ├── 261FB04002.pdf
        └── ...
```

---

## 3. Installation Instructions

### Prerequisites
- Python 3.10+ (Python 3.12 recommended)
- Windows, macOS, or Linux

### Install Python Dependencies
Open your command prompt or terminal in the project directory and run:

```bash
pip install -r requirements.txt
```

`requirements.txt` includes:
- `openpyxl`: For reading Excel `.xlsx` files and evaluating cached formula cells.
- `python-docx`: For manipulating Word `.docx` documents across runs and tables.
- `pywin32`: Enables native Microsoft Word COM automation on Windows.
- `pypdf`: Used for automated PDF extraction and test verification.

---

## 4. PDF Conversion Engine Setup

The tool converts DOCX certificates to PDF while guaranteeing 100% layout, font, margin, and graphic fidelity:

1. **LibreOffice (Headless Mode - Preferred across platforms)**
   The tool automatically detects `soffice` or `libreoffice` in your system `PATH` and standard installation directories (`C:\Program Files\LibreOffice\program\soffice.exe`, `/usr/bin/libreoffice`, `/Applications/LibreOffice.app`, etc.).
   
   - **Windows Installation (winget):**
     ```powershell
     winget install TheDocumentFoundation.LibreOffice
     ```
   - **Ubuntu/Debian Installation:**
     ```bash
     sudo apt update && sudo apt install -y libreoffice
     ```
   - **macOS Installation (Homebrew):**
     ```bash
     brew install --cask libreoffice
     ```
   - **Manual Download:**
     Download directly from [https://www.libreoffice.org/download/](https://www.libreoffice.org/download/).

2. **Microsoft Word COM Automation (Windows Native)**
   On Windows machines with Microsoft Word installed, the tool can also use Microsoft Word's native COM interface (`win32com`). Word provides native rendering fidelity and requires no additional software.

---

## 5. Input Data Format

### Excel File Format (`*.xlsx`)
The Excel file must contain these three core columns in the header row:
- `SI NO` (or `S.NO`, `SL NO`)
- `REG NO` (or `REG_NO`, `Registration No`)
- `NAME` (or `Student Name`)

Column headers are automatically trimmed and normalized (case-insensitive, ignores accidental spacing or underscores).

Example:

| SI NO | REG NO  | NAME         |
| ----: | ------- | ------------ |
|     1 | 23CS001 | Rahul Kumar  |
|     2 | 23CS002 | Priya Sharma |
|     3 | 23CS003 | Arun Kumar   |

#### Important Data Rules:
- **String Registration Numbers**: `REG NO` is always preserved as a string (e.g. `00123` will never be truncated to `123`).
- **Alphanumeric & Hyphenated IDs**: Supports registration numbers containing letters, numbers, hyphens, and slashes.
- **Unicode Names**: Student names support spaces, accents, and Indian Unicode scripts.
- **Blank Rows**: Completely blank rows are automatically skipped.

### DOCX Template Format (`*.docx`)
The certificate template must contain the following placeholder tags:
- `{{REG_NO}}` (or `{{REG NO}}`)
- `{{NAME}}`

Example paragraph in template:
```text
This certificate is proudly presented to

{{NAME}}

Registration Number: {{REG_NO}}

for successfully completing the course.
```

#### Formatting Preservation:
- **No Redesign**: The template is the single source of truth. All page margins, orientations, background graphics, borders, logos, shapes, tables, and typography are preserved.
- **Split-Run Handling**: In Word documents, placeholders often get broken across multiple internal XML runs (e.g. `{{` + `NAME` + `}}`). The tool reconstructs and replaces placeholders cleanly across runs without resetting character formatting.
- **Long Name Auto-Scaling**: For students with long names (e.g. > 28 characters), the tool slightly scales the font size to prevent text overflow or page wrapping.
- **Template Safety**: The original template file is **never modified**; operations take place on temporary working copies.

---

## 6. Command-Line Usage

### Default Usage (Auto-Discovery)
Simply run the script from inside any folder containing your Excel sheet and template:

```powershell
python main.py
```

### Optional Command-Line Arguments
For advanced automation or non-interactive CI/CD pipelines, several flags are available:

```powershell
# Specify input files explicitly
python main.py --excel students.xlsx --template certificate_template.docx

# Run non-interactively (auto-confirm generation)
python main.py --yes

# Overwrite existing certificates without interactive confirmation
python main.py --overwrite

# Select a specific PDF conversion engine
python main.py --engine libreoffice
python main.py --engine word

# Dry-run validation only (previews validation checks without creating PDFs)
python main.py --dry-run

# Customize student detail font size and weight (defaults: 20.0 pt, bold)
python main.py --font-size 22.0
python main.py --no-bold

# Specify custom root or output directories
python main.py --root-dir "D:\MyCertificates" --output "D:\MyCertificates\output\certificates"
```

---

## 7. Pre-Flight Validations & Safety Features

1. **Missing Column Validation**:
   If a required column is missing, the tool stops immediately and displays:
   ```text
   ERROR:
   The Excel file is missing the required column:
   NAME
   ```

2. **Template Placeholder Validation**:
   Verifies that both `{{REG_NO}}` and `{{NAME}}` exist in the DOCX file before processing.

3. **Per-Row Validation**:
   Detects missing names or missing registration numbers on individual rows, reporting row numbers:
   ```text
   Row 15:
   REG NO: 23CS015
   NAME: EMPTY

   Status: INVALID
   Reason: NAME is missing
   ```

4. **Duplicate Registration Number Protection**:
   Registration numbers must uniquely identify certificates. If duplicates are found, the tool reports the duplicate and the exact row numbers, aborting generation to prevent overwriting:
   ```text
   Duplicate REG NO detected:

   23CS001

   Rows:
   2
   15
   ```

5. **Existing File Conflict Protection**:
   If an output file (e.g. `output/certificates/23CS001.pdf`) already exists, the tool will never silently overwrite it. It prompts the user for confirmation or requires the `--overwrite` flag.

---

## 8. Generation Log (`generation_log.csv`)

Every run produces an audit log at `output/generation_log.csv`:

| SI NO | REG NO  | NAME         | STATUS  | OUTPUT      | ERROR |
| ----: | ------- | ------------ | ------- | ----------- | ----- |
|     1 | 23CS001 | Rahul Kumar  | SUCCESS | 23CS001.pdf |       |
|     2 | 23CS002 | Priya Sharma | SUCCESS | 23CS002.pdf |       |
|     3 | 23CS003 | Arun Kumar   | SUCCESS | 23CS003.pdf |       |

---

## 9. Automated Testing

To run the complete automated test suite (including unit tests and the 3-student end-to-end integration test):

```bash
python -m unittest tests/test_suite.py
```

The test suite validates:
- Column header normalization and whitespace handling.
- Leading zero preservation in registration numbers (e.g. `'00123'`).
- Multi-run XML placeholder replacements in paragraphs and shapes.
- Template file immutability (SHA-256 hash comparison before and after execution).
- Duplicate registration number detection.
- Generation of `23CS001.pdf`, `23CS002.pdf`, `23CS003.pdf`.
- Verification of text inside the generated PDFs using `pypdf`.
- Output log formatting.
