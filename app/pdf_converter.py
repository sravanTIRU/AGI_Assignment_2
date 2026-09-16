"""
PDF Converter Module for Student Certificate Generator.

Converts DOCX certificates to PDF format using LibreOffice in headless mode
(preferred) or Microsoft Word COM automation (Windows fallback), ensuring 100%
layout, font, and design preservation.
"""

import os
import sys
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Optional, Tuple


class PDFConversionError(Exception):
    """Raised when PDF conversion fails."""
    pass


class PDFEngineNotFoundError(PDFConversionError):
    """Raised when neither LibreOffice nor Word is available."""
    pass


def find_libreoffice_executable() -> Optional[str]:
    """
    Locate the LibreOffice / soffice executable across platforms.
    Checks PATH and standard operating system installation locations.
    """
    # 1. Check PATH
    for cmd in ("soffice", "libreoffice", "soffice.exe", "libreoffice.exe"):
        found = shutil.which(cmd)
        if found and Path(found).is_file():
            return str(Path(found).resolve())

    # 2. Check Windows common paths
    if sys.platform == "win32":
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")

        candidate_paths = [
            Path(program_files) / "LibreOffice" / "program" / "soffice.exe",
            Path(program_files_x86) / "LibreOffice" / "program" / "soffice.exe",
            Path(local_app_data) / "Programs" / "LibreOffice" / "program" / "soffice.exe",
            Path(r"C:\Program Files\LibreOffice 7\program\soffice.exe"),
            Path(r"C:\Program Files\LibreOffice 24\program\soffice.exe"),
            Path(r"C:\Program Files\LibreOffice 25\program\soffice.exe"),
            Path(r"C:\Program Files\LibreOffice 26\program\soffice.exe"),
        ]
        for p in candidate_paths:
            if p.exists() and p.is_file():
                return str(p.resolve())

    # 3. Check macOS common paths
    elif sys.platform == "darwin":
        mac_paths = [
            Path("/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            Path("~/Applications/LibreOffice.app/Contents/MacOS/soffice").expanduser(),
        ]
        for p in mac_paths:
            if p.exists() and p.is_file():
                return str(p.resolve())

    # 4. Check Linux common paths
    else:
        linux_paths = [
            Path("/usr/bin/libreoffice"),
            Path("/usr/bin/soffice"),
            Path("/usr/local/bin/libreoffice"),
            Path("/usr/local/bin/soffice"),
            Path("/snap/bin/libreoffice"),
        ]
        for p in linux_paths:
            if p.exists() and p.is_file():
                return str(p.resolve())

    return None


def is_msword_available() -> bool:
    """Check if Microsoft Word COM automation is available on Windows."""
    if sys.platform != "win32":
        return False
    try:
        import win32com.client
        word = win32com.client.Dispatch("Word.Application")
        word.Quit()
        return True
    except Exception:
        return False


class PDFConverter:
    """
    Manages DOCX to PDF conversion.
    Prefers LibreOffice headless; falls back to MS Word COM on Windows.
    """

    def __init__(self, force_engine: Optional[str] = None):
        self.force_engine = force_engine
        self.engine_type: str = "none"
        self.soffice_path: Optional[str] = None
        self._word_app = None

        self._detect_engine()

    def _detect_engine(self) -> None:
        """Detect and initialize the best available conversion engine."""
        soffice = find_libreoffice_executable()
        word_ok = is_msword_available() if sys.platform == "win32" else False

        if self.force_engine == "libreoffice":
            if soffice:
                self.engine_type = "libreoffice"
                self.soffice_path = soffice
            else:
                self._raise_libreoffice_missing()
        elif self.force_engine == "word":
            if word_ok:
                self.engine_type = "word"
            else:
                raise PDFEngineNotFoundError("Microsoft Word COM automation is not available.")
        else:
            # Default auto-detection: prefer LibreOffice if available
            if soffice:
                self.engine_type = "libreoffice"
                self.soffice_path = soffice
            elif word_ok:
                self.engine_type = "word"
            else:
                self._raise_libreoffice_missing()

    def _raise_libreoffice_missing(self) -> None:
        msg = (
            "LibreOffice was not found. Please install LibreOffice to enable PDF conversion.\n"
            "Download from: https://www.libreoffice.org/download/download-libreoffice/\n"
            "Or install on Windows via winget:\n"
            "  winget install TheDocumentFoundation.LibreOffice"
        )
        raise PDFEngineNotFoundError(msg)

    def convert_docx_to_pdf(self, docx_path: str, output_pdf_path: str) -> None:
        """
        Convert a DOCX file to PDF, saving exactly to output_pdf_path.
        """
        input_path = Path(docx_path).resolve()
        target_pdf = Path(output_pdf_path).resolve()

        if not input_path.exists():
            raise FileNotFoundError(f"DOCX file not found: {docx_path}")

        target_pdf.parent.mkdir(parents=True, exist_ok=True)

        if self.engine_type == "libreoffice":
            self._convert_with_libreoffice(input_path, target_pdf)
        elif self.engine_type == "word":
            self._convert_with_word(input_path, target_pdf)
        else:
            self._raise_libreoffice_missing()

    def _convert_with_libreoffice(self, input_path: Path, target_pdf: Path) -> None:
        """Execute headless LibreOffice command."""
        # LibreOffice outputs to outdir with same basename as input file.
        # We convert in target_pdf's directory, then rename if necessary.
        outdir = target_pdf.parent
        cmd = [
            self.soffice_path,
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            str(outdir),
            str(input_path)
        ]

        try:
            result = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=90
            )
        except subprocess.TimeoutExpired:
            raise PDFConversionError(f"LibreOffice conversion timed out for {input_path.name}")
        except Exception as e:
            raise PDFConversionError(f"Failed to execute LibreOffice: {e}")

        expected_lo_pdf = outdir / f"{input_path.stem}.pdf"
        if not expected_lo_pdf.exists():
            err = result.stderr.strip() or result.stdout.strip() or "Unknown error"
            raise PDFConversionError(f"LibreOffice failed to generate PDF for {input_path.name}: {err}")

        # Rename to target_pdf if stem differs
        if expected_lo_pdf != target_pdf:
            if target_pdf.exists():
                target_pdf.unlink()
            expected_lo_pdf.rename(target_pdf)

    def _convert_with_word(self, input_path: Path, target_pdf: Path) -> None:
        """Convert using Microsoft Word COM automation."""
        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        word = None
        doc = None
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0

            # 17 = wdFormatPDF
            doc = word.Documents.Open(str(input_path))
            doc.SaveAs(str(target_pdf), FileFormat=17)
        except Exception as e:
            raise PDFConversionError(f"Microsoft Word PDF conversion failed for {input_path.name}: {e}")
        finally:
            if doc is not None:
                try:
                    doc.Close(SaveChanges=0)
                except Exception:
                    pass
            if word is not None:
                try:
                    word.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    def get_engine_name(self) -> str:
        """Return human-readable name of the active conversion engine."""
        if self.engine_type == "libreoffice":
            return "LibreOffice (Headless)"
        if self.engine_type == "word":
            return "Microsoft Word (Native COM)"
        return "None"
