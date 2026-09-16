"""
Automated Test Suite for Certificate Mail-Merge Agent.

Covers:
1. Excel data parsing and string REG NO preservation (e.g. '00123').
2. Multi-run DOCX placeholder replacement without modifying the original template.
3. Validation rules: missing columns, duplicate REG NOs, invalid rows, missing placeholders.
4. Full End-to-End test with 3 sample students (23CS001, 23CS002, 23CS003) and PDF content verification.
"""

import os
import sys
import hashlib
import tempfile
import unittest
from pathlib import Path
import openpyxl
import docx
import pypdf

# Add project root to sys.path
PROJECT_ROOT = Path(__file__).parent.parent.resolve()
sys.path.insert(0, str(PROJECT_ROOT))

from app.excel_reader import (
    normalize_col_name,
    format_reg_no,
    inspect_excel_columns,
    read_students_from_excel,
    detect_student_excel
)
from app.template_processor import (
    inspect_docx_template,
    detect_docx_template,
    create_individual_certificate_docx,
    replace_placeholder_in_paragraph,
    REG_NO_REGEX,
    NAME_REGEX
)
from app.validator import (
    validate_student_records,
    check_duplicate_registration_numbers,
    format_duplicate_error,
    format_invalid_student_row
)
from app.pdf_converter import PDFConverter
from app.certificate_generator import CertificateGenerator


def file_sha256(filepath: str) -> str:
    """Compute SHA256 checksum of a file."""
    h = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(8192):
            h.update(chunk)
    return h.hexdigest()


class TestExcelReader(unittest.TestCase):
    """Tests for Excel discovery, normalization, and data parsing."""

    def test_normalize_col_name(self):
        self.assertEqual(normalize_col_name("  SI NO  "), "SI NO")
        self.assertEqual(normalize_col_name("S.NO"), "SI NO")
        self.assertEqual(normalize_col_name("SL NO"), "SI NO")
        self.assertEqual(normalize_col_name("REG NO"), "REG NO")
        self.assertEqual(normalize_col_name("REG_NO"), "REG NO")
        self.assertEqual(normalize_col_name("Registration No"), "REG NO")
        self.assertEqual(normalize_col_name("  NAME  "), "NAME")
        self.assertEqual(normalize_col_name("Student Name"), "NAME")

    def test_format_reg_no_preserves_string(self):
        # Leading zero preservation
        self.assertEqual(format_reg_no("00123"), "00123")
        self.assertEqual(format_reg_no(123), "123")
        self.assertEqual(format_reg_no(123.0), "123")
        self.assertEqual(format_reg_no("23CS001"), "23CS001")
        self.assertEqual(format_reg_no(" 23CS-005 "), "23CS-005")

    def test_read_sample_excel_with_formulas_and_blanks(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = Path(tmpdir) / "test_students.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Students"
            # Header
            ws.append([" SI NO ", " REG NO ", " NAME "])
            # Valid row
            ws.append([1, "00123", "Rahul Kumar"])
            # Formula row
            ws.append(["=ROW()-1", "23CS002", "Priya Sharma"])
            # Blank row (should be skipped)
            ws.append([None, None, None])
            # Unicode / Indian name row
            ws.append([3, "23CS003", "Arun Kumar"])
            # Invalid row (missing name)
            ws.append([4, "23CS004", None])
            # Invalid row (missing reg no)
            ws.append([5, None, "Sneha Patel"])
            wb.save(excel_path)

            students = read_students_from_excel(str(excel_path))
            # Expect 5 student records (1 blank was skipped)
            self.assertEqual(len(students), 5)

            # Check first student preserves leading zeroes
            self.assertEqual(students[0]["reg_no"], "00123")
            self.assertEqual(students[0]["name"], "Rahul Kumar")
            self.assertTrue(students[0]["is_valid"])

            # Check second student formula handled
            self.assertEqual(students[1]["reg_no"], "23CS002")
            self.assertEqual(students[1]["name"], "Priya Sharma")
            self.assertTrue(students[1]["is_valid"])

            # Check invalid rows
            self.assertFalse(students[3]["is_valid"])
            self.assertEqual(students[3]["error"], "NAME is missing")

            self.assertFalse(students[4]["is_valid"])
            self.assertEqual(students[4]["error"], "REG NO is missing")


class TestTemplateProcessor(unittest.TestCase):
    """Tests for template inspection and run-level placeholder replacement."""

    def test_multi_run_placeholder_replacement(self):
        doc = docx.Document()
        p = doc.add_paragraph()
        # Split placeholder across 3 runs
        r1 = p.add_run("Awarded to ")
        r2 = p.add_run("{{NA")
        r3 = p.add_run("ME}}")
        r4 = p.add_run(" with Reg: {{")
        r5 = p.add_run("REG_NO}}.")

        # Replace NAME
        replace_placeholder_in_paragraph(p, NAME_REGEX, "Rahul Kumar")
        self.assertIn("Rahul Kumar", p.text)
        self.assertNotIn("{{NAME}}", p.text)

        # Replace REG_NO
        replace_placeholder_in_paragraph(p, REG_NO_REGEX, "23CS001")
        self.assertIn("23CS001", p.text)
        self.assertNotIn("{{REG_NO}}", p.text)

        self.assertEqual(p.text, "Awarded to Rahul Kumar with Reg: 23CS001.")

    def test_original_template_not_modified(self):
        template_file = PROJECT_ROOT / "certificate_template.docx"
        if not template_file.exists():
            self.skipTest("certificate_template.docx not in project root")

        hash_before = file_sha256(str(template_file))

        with tempfile.TemporaryDirectory() as tmpdir:
            out_docx = Path(tmpdir) / "output.docx"
            create_individual_certificate_docx(
                template_path=str(template_file),
                output_docx_path=str(out_docx),
                reg_no="23CS001",
                name="Rahul Kumar"
            )

            # Verify output docx has replaced fields
            out_doc = docx.Document(str(out_docx))
            out_text = " ".join(p.text for p in out_doc.paragraphs)
            self.assertIn("Rahul Kumar", out_text)
            self.assertIn("23CS001", out_text)
            self.assertNotIn("{{NAME}}", out_text)
            self.assertNotIn("{{REG NO}}", out_text)
            self.assertNotIn("{{REG_NO}}", out_text)

        hash_after = file_sha256(str(template_file))
        self.assertEqual(hash_before, hash_after, "Original template MUST remain unchanged!")


class TestValidator(unittest.TestCase):
    """Tests for duplicate REG NO detection and row validation."""

    def test_duplicate_registration_numbers(self):
        students = [
            {"row_num": 2, "reg_no": "23CS001", "name": "Rahul Kumar", "is_valid": True},
            {"row_num": 3, "reg_no": "23CS002", "name": "Priya Sharma", "is_valid": True},
            {"row_num": 15, "reg_no": "23CS001", "name": "Rahul Duplicate", "is_valid": True},
        ]
        duplicates = check_duplicate_registration_numbers(students)
        self.assertIn("23CS001", duplicates)
        self.assertEqual(duplicates["23CS001"], [2, 15])

        error_msg = format_duplicate_error(duplicates)
        self.assertIn("Duplicate REG NO detected:", error_msg)
        self.assertIn("23CS001", error_msg)
        self.assertIn("Rows:", error_msg)
        self.assertIn("2", error_msg)
        self.assertIn("15", error_msg)


class TestEndToEndSampleStudents(unittest.TestCase):
    """
    End-to-End Test with 3 sample students specified in prompt:
    1 | 23CS001 | Rahul Kumar
    2 | 23CS002 | Priya Sharma
    3 | 23CS003 | Arun Kumar

    Verifies:
    - Generation of 23CS001.pdf, 23CS002.pdf, 23CS003.pdf
    - Filename is exactly REG_NO.pdf
    - Content inside each PDF contains exact REG NO and student NAME
    - generation_log.csv records SUCCESS for all 3 students
    - Original template remains completely untouched
    """

    def test_three_students_workflow(self):
        template_source = PROJECT_ROOT / "certificate_template.docx"
        if not template_source.exists():
            self.skipTest("certificate_template.docx not found")

        hash_original = file_sha256(str(template_source))

        with tempfile.TemporaryDirectory() as tmpdir:
            work_dir = Path(tmpdir)

            # Copy template into work_dir
            tmpl_copy = work_dir / "certificate_template.docx"
            tmpl_copy.write_bytes(template_source.read_bytes())

            # Create sample 3-student Excel file
            excel_path = work_dir / "students.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Students"
            ws.append(["SI NO", "REG NO", "NAME"])
            ws.append([1, "23CS001", "Rahul Kumar"])
            ws.append([2, "23CS002", "Priya Sharma"])
            ws.append([3, "23CS003", "Arun Kumar"])
            wb.save(excel_path)

            # Run CertificateGenerator
            generator = CertificateGenerator(
                root_dir=str(work_dir),
                auto_confirm=True,
                overwrite=True
            )
            success = generator.run()
            self.assertTrue(success, "Generator run should succeed")

            # Verify output files
            out_cert_dir = work_dir / "output" / "certificates"
            self.assertTrue(out_cert_dir.exists())

            pdf1 = out_cert_dir / "23CS001.pdf"
            pdf2 = out_cert_dir / "23CS002.pdf"
            pdf3 = out_cert_dir / "23CS003.pdf"

            self.assertTrue(pdf1.exists(), "23CS001.pdf must exist")
            self.assertTrue(pdf2.exists(), "23CS002.pdf must exist")
            self.assertTrue(pdf3.exists(), "23CS003.pdf must exist")

            # Verify PDF contents with pypdf
            def get_pdf_text(p: Path) -> str:
                reader = pypdf.PdfReader(str(p))
                return "".join(page.extract_text() or "" for page in reader.pages)

            text1 = get_pdf_text(pdf1)
            self.assertIn("23CS001", text1)
            self.assertIn("Rahul Kumar", text1)
            self.assertNotIn("Priya Sharma", text1)
            self.assertNotIn("Arun Kumar", text1)

            text2 = get_pdf_text(pdf2)
            self.assertIn("23CS002", text2)
            self.assertIn("Priya Sharma", text2)
            self.assertNotIn("Rahul Kumar", text2)
            self.assertNotIn("Arun Kumar", text2)

            text3 = get_pdf_text(pdf3)
            self.assertIn("23CS003", text3)
            self.assertIn("Arun Kumar", text3)
            self.assertNotIn("Rahul Kumar", text3)
            self.assertNotIn("Priya Sharma", text3)

            # Verify generation_log.csv
            log_csv = work_dir / "output" / "generation_log.csv"
            self.assertTrue(log_csv.exists(), "generation_log.csv must exist")
            log_content = log_csv.read_text(encoding="utf-8-sig")
            self.assertIn("23CS001,Rahul Kumar,SUCCESS,23CS001.pdf", log_content)
            self.assertIn("23CS002,Priya Sharma,SUCCESS,23CS002.pdf", log_content)
            self.assertIn("23CS003,Arun Kumar,SUCCESS,23CS003.pdf", log_content)

            # Verify original template copy in work_dir was not altered
            hash_after_run = file_sha256(str(tmpl_copy))
            self.assertEqual(hash_after_run, hash_original)


if __name__ == "__main__":
    unittest.main(verbosity=2)
