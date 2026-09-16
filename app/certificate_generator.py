"""
Certificate Generator Orchestration Module.

Coordinates file detection, validation, template merging, PDF conversion,
conflict resolution, and progress logging.
"""

import os
import sys
import tempfile
import shutil
from pathlib import Path
from typing import Optional, List, Dict, Any, Tuple

from app.excel_reader import detect_student_excel, read_students_from_excel
from app.template_processor import detect_docx_template, inspect_docx_template, create_individual_certificate_docx
from app.validator import (
    validate_student_records,
    check_duplicate_registration_numbers,
    format_duplicate_error,
    format_invalid_student_row,
    ValidationError
)
from app.pdf_converter import PDFConverter, PDFEngineNotFoundError, PDFConversionError
from app.logger import GenerationLogger


def _symbol(char: str, fallback: str) -> str:
    """Return unicode character if supported by stdout, else fallback."""
    try:
        encoding = sys.stdout.encoding or "utf-8"
        char.encode(encoding)
        return char
    except Exception:
        return fallback


CHECK = _symbol("✓", "[OK]")
CROSS = _symbol("✗", "[X]")


class CertificateGenerator:
    """End-to-end Certificate Generation workflow manager."""

    def __init__(
        self,
        root_dir: Optional[str] = None,
        excel_path: Optional[str] = None,
        template_path: Optional[str] = None,
        output_dir: Optional[str] = None,
        auto_confirm: bool = False,
        overwrite: bool = False,
        force_engine: Optional[str] = None,
        font_size: float = 20.0,
        bold: bool = True
    ):
        self.root_dir = Path(root_dir or os.getcwd()).resolve()
        self.explicit_excel = excel_path
        self.explicit_template = template_path
        self.font_size = font_size
        self.bold = bold
        
        # Default output directory is root_dir / output / certificates
        if output_dir:
            self.certificates_dir = Path(output_dir).resolve()
            self.output_root = self.certificates_dir.parent
        else:
            self.output_root = self.root_dir / "output"
            self.certificates_dir = self.output_root / "certificates"

        self.auto_confirm = auto_confirm
        self.overwrite = overwrite
        self.force_engine = force_engine

        self.excel_file: Optional[str] = None
        self.template_file: Optional[str] = None
        self.students: List[Dict[str, Any]] = []
        self.valid_students: List[Dict[str, Any]] = []
        self.invalid_students: List[Dict[str, Any]] = []
        self.duplicates: Dict[str, List[int]] = {}

    def discover_and_validate(self) -> bool:
        """
        Step 1: Discover input files and perform thorough pre-flight validation.
        Prints clean diagnostic reports matching project guidelines.
        Returns True if pre-flight passed and ready to proceed.
        """
        print("\nStudent Certificate Generator")
        print("=" * 32)
        print(f"\nRoot directory:\n{self.root_dir}\n")

        # 1. Discover & Validate Excel File
        try:
            self.excel_file = detect_student_excel(str(self.root_dir), self.explicit_excel)
            print(f"Excel file detected:\n{Path(self.excel_file).name}\n")
        except Exception as e:
            print(f"ERROR:\n{e}\n")
            return False

        # 2. Discover & Validate DOCX Template
        try:
            self.template_file = detect_docx_template(str(self.root_dir), self.explicit_template)
            print(f"Certificate template detected:\n{Path(self.template_file).name}\n")
        except Exception as e:
            print(f"ERROR:\n{e}\n")
            return False

        # 3. Read Student Data
        try:
            self.students = read_students_from_excel(self.excel_file)
            print(f"Students detected:\n{len(self.students)}\n")
        except Exception as e:
            print(f"ERROR reading Excel file:\n{e}\n")
            return False

        if not self.students:
            print("ERROR:\nNo student records found in the Excel file.\n")
            return False

        # 4. Validate Template Placeholders
        is_tmpl_valid, missing_placeholders = inspect_docx_template(self.template_file)
        print("Template placeholders:")
        if "{{REG_NO}}" in missing_placeholders:
            print(f"{CROSS} {{{{REG_NO}}}} (Missing)")
        else:
            print(f"{CHECK} {{{{REG_NO}}}}")

        if "{{NAME}}" in missing_placeholders:
            print(f"{CROSS} {{{{NAME}}}} (Missing)")
        else:
            print(f"{CHECK} {{{{NAME}}}}")
        print()

        if not is_tmpl_valid:
            print("ERROR:")
            print(f"The certificate template does not contain the required placeholder:\n{missing_placeholders[0]}\n")
            return False

        # 5. Validate Student Rows
        self.valid_students, self.invalid_students = validate_student_records(self.students)

        if self.invalid_students:
            print("WARNING: Invalid student row(s) detected:")
            for inv in self.invalid_students:
                print(format_invalid_student_row(inv))
                print()

        if not self.valid_students:
            print("ERROR:\nNo valid student records found to generate certificates.\n")
            return False

        # 6. Check Duplicate Registration Numbers
        self.duplicates = check_duplicate_registration_numbers(self.valid_students)
        if self.duplicates:
            print("ERROR:")
            print(format_duplicate_error(self.duplicates))
            print("\nGeneration aborted: Duplicate REG NO values must be resolved.\n")
            return False

        print("Validation:")
        print(f"{CHECK} Excel structure valid")
        print(f"{CHECK} Certificate template valid")
        print(f"{CHECK} No duplicate registration numbers\n")

        print(f"Ready to generate {len(self.valid_students)} certificates.\n")
        return True

    def check_existing_files(self) -> Tuple[List[str], bool]:
        """
        Check if certificates already exist in the output directory.
        Returns:
            (existing_pdf_filenames, should_proceed)
        """
        existing: List[str] = []
        for s in self.valid_students:
            target_name = f"{s['reg_no']}.pdf"
            target_path = self.certificates_dir / target_name
            if target_path.exists():
                existing.append(target_name)

        if not existing:
            return [], True

        if self.overwrite:
            return existing, True

        print(f"WARNING: {len(existing)} certificate(s) already exist in output/certificates/:")
        for name in existing[:5]:
            print(f" - {name}")
        if len(existing) > 5:
            print(f"   ... and {len(existing) - 5} more")

        if self.auto_confirm:
            print("Overwrite flag was not specified. Aborting to protect existing files.")
            return existing, False

        # Ask user interactively
        while True:
            response = input("\nDo you want to overwrite existing certificates? [y/N]: ").strip().lower()
            if response in ("y", "yes"):
                return existing, True
            if response in ("n", "no", ""):
                print("Generation cancelled by user to protect existing certificates.\n")
                return existing, False
            print("Please answer 'y' or 'n'.")

    def run(self) -> bool:
        """Execute the full generation process."""
        if not self.discover_and_validate():
            return False

        # Check existing files
        _, should_proceed = self.check_existing_files()
        if not should_proceed:
            return False

        # Prompt user confirmation
        if not self.auto_confirm:
            try:
                choice = input("Generate certificates? [Y/n]: ").strip().lower()
                if choice in ("n", "no"):
                    print("Generation cancelled.")
                    return False
            except (KeyboardInterrupt, EOFError):
                print("\nCancelled.")
                return False

        # Initialize PDF converter
        try:
            converter = PDFConverter(force_engine=self.force_engine)
        except Exception as e:
            print(f"\nERROR initializing PDF engine:\n{e}\n")
            return False

        # Ensure output directories exist
        self.certificates_dir.mkdir(parents=True, exist_ok=True)
        logger = GenerationLogger(str(self.output_root))

        # Log any pre-existing invalid student rows first
        for inv in self.invalid_students:
            logger.log_failure(
                si_no=inv.get("si_no", ""),
                reg_no=inv.get("reg_no", ""),
                name=inv.get("name", ""),
                error_message=inv.get("error", "Invalid row")
            )

        print("\nGenerating certificates...")
        print(f"Conversion engine: {converter.get_engine_name()}\n")

        temp_dir = Path(tempfile.mkdtemp(prefix="cert_merge_"))

        total = len(self.valid_students)
        success_count = 0
        failed_count = 0

        try:
            for idx, student in enumerate(self.valid_students, start=1):
                reg_no = student["reg_no"]
                name = student["name"]
                si_no = student["si_no"]

                progress_prefix = f"[{idx}/{total}] {reg_no} - {name}"

                # Paths
                temp_docx = temp_dir / f"{reg_no}.docx"
                target_pdf = self.certificates_dir / f"{reg_no}.pdf"

                try:
                    # 1. Create modified DOCX copy
                    create_individual_certificate_docx(
                        template_path=self.template_file,
                        output_docx_path=str(temp_docx),
                        reg_no=reg_no,
                        name=name,
                        font_size=self.font_size,
                        bold=self.bold
                    )

                    # 2. Convert to PDF
                    converter.convert_docx_to_pdf(
                        docx_path=str(temp_docx),
                        output_pdf_path=str(target_pdf)
                    )

                    # 3. Clean up temp DOCX
                    if temp_docx.exists():
                        temp_docx.unlink()

                    # 4. Log Success
                    logger.log_success(
                        si_no=si_no,
                        reg_no=reg_no,
                        name=name,
                        output_filename=target_pdf.name
                    )
                    success_count += 1
                    print(f"{progress_prefix}")

                except Exception as e:
                    failed_count += 1
                    error_str = str(e)
                    logger.log_failure(
                        si_no=si_no,
                        reg_no=reg_no,
                        name=name,
                        error_message=error_str
                    )
                    print(f"{progress_prefix} [FAILED: {error_str}]")

        finally:
            # Clean up temp working directory
            shutil.rmtree(temp_dir, ignore_errors=True)

        print("\nGeneration complete.\n")
        print(f"Successful: {success_count}")
        print(f"Failed: {failed_count}\n")
        print("Certificates saved to:")
        print(f"{self.certificates_dir}\\\n")
        print("Log saved to:")
        print(f"{logger.log_path}\n")

        return failed_count == 0
