"""
Template Processor Module for Student Certificate Generator.

Responsible for discovering DOCX templates, validating required placeholders,
and performing formatting-preserving placeholder replacement across runs,
tables, headers, footers, and textboxes.
"""

import os
import re
import shutil
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Any
import docx
from docx.text.paragraph import Paragraph
from docx.oxml import OxmlElement
from docx.shared import Pt


IGNORE_DIRS = {
    ".git", ".venv", "venv", "env", "__pycache__",
    "output", ".idea", ".vscode", "scratch", ".system_generated"
}

# Regex to detect REG_NO placeholder: {{REG_NO}} or {{REG NO}} (case-insensitive, optional whitespace)
REG_NO_REGEX = re.compile(r"\{\{\s*REG[_ ]?NO\s*\}\}", re.IGNORECASE)
# Regex to detect NAME placeholder: {{NAME}} (case-insensitive, optional whitespace)
NAME_REGEX = re.compile(r"\{\{\s*NAME\s*\}\}", re.IGNORECASE)


def find_docx_templates(root_dir: str) -> List[str]:
    """
    Search root_dir and its subdirectories for .docx files,
    ignoring temporary files (~$*) and common build/system directories.
    """
    root_path = Path(root_dir).resolve()
    found_files: List[str] = []

    # First search immediate directory
    try:
        for entry in os.scandir(root_path):
            if entry.is_file() and entry.name.lower().endswith(".docx") and not entry.name.startswith("~$"):
                found_files.append(str(Path(entry.path).resolve()))
    except OSError:
        pass

    # Then search subdirectories
    for current_root, dirs, files in os.walk(root_path):
        dirs[:] = [d for d in dirs if d.lower() not in IGNORE_DIRS and not d.startswith(".")]
        
        if Path(current_root).resolve() == root_path:
            continue

        for f in files:
            if f.lower().endswith(".docx") and not f.startswith("~$"):
                found_files.append(str(Path(current_root, f).resolve()))

    return found_files


def get_all_document_paragraphs(doc: docx.Document) -> List[Paragraph]:
    """
    Collect all Paragraph objects in a document, including:
    - Normal body paragraphs
    - Table cells (nested tables included)
    - Headers and footers (all sections, first page, even page)
    - Textboxes / Shapes (w:txbxContent)
    """
    paragraphs: List[Paragraph] = []

    # 1. Body paragraphs
    paragraphs.extend(doc.paragraphs)

    # 2. Table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paragraphs.extend(cell.paragraphs)
                # Check nested tables in cells
                for nested_table in cell.tables:
                    for n_row in nested_table.rows:
                        for n_cell in n_row.cells:
                            paragraphs.extend(n_cell.paragraphs)

    # 3. Headers and Footers
    for section in doc.sections:
        for header in [section.header, section.first_page_header, section.even_page_header]:
            if header is not None:
                paragraphs.extend(header.paragraphs)
                for table in header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            paragraphs.extend(cell.paragraphs)
        for footer in [section.footer, section.first_page_footer, section.even_page_footer]:
            if footer is not None:
                paragraphs.extend(footer.paragraphs)
                for table in footer.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            paragraphs.extend(cell.paragraphs)

    # 4. Textboxes / Shapes in XML
    try:
        txbx_paras = doc._element.xpath(".//w:txbxContent//w:p")
        for p_elm in txbx_paras:
            paragraphs.append(Paragraph(p_elm, doc))
    except Exception:
        pass

    return paragraphs


def inspect_docx_template(file_path: str) -> Tuple[bool, List[str]]:
    """
    Inspect DOCX file to see if it contains {{REG_NO}} (or {{REG NO}}) and {{NAME}}.
    Returns:
        (is_valid, missing_placeholders)
    """
    try:
        doc = docx.Document(file_path)
    except Exception:
        return False, ["{{REG_NO}}", "{{NAME}}"]

    has_reg_no = False
    has_name = False

    # Check paragraph texts
    paragraphs = get_all_document_paragraphs(doc)
    for p in paragraphs:
        text = p.text
        if REG_NO_REGEX.search(text):
            has_reg_no = True
        if NAME_REGEX.search(text):
            has_name = True

    # Also check full document XML in case placeholder was fragmented across complex tags
    if not (has_reg_no and has_name):
        try:
            full_xml = doc._element.xml
            if not has_reg_no and REG_NO_REGEX.search(full_xml):
                has_reg_no = True
            if not has_name and NAME_REGEX.search(full_xml):
                has_name = True
        except Exception:
            pass

    missing: List[str] = []
    if not has_reg_no:
        missing.append("{{REG_NO}}")
    if not has_name:
        missing.append("{{NAME}}")

    return (len(missing) == 0), missing


def detect_docx_template(root_dir: str, explicit_path: Optional[str] = None) -> str:
    """
    Auto-detect the certificate template in root_dir or validate explicit_path.
    Raises descriptive RuntimeError/FileNotFoundError if missing or ambiguous.
    """
    if explicit_path:
        resolved = Path(explicit_path)
        if not resolved.is_absolute():
            resolved = Path(root_dir) / resolved
        if not resolved.exists():
            raise FileNotFoundError(f"Specified certificate template not found: {explicit_path}")
        is_valid, missing = inspect_docx_template(str(resolved))
        if not is_valid:
            raise ValueError(
                f"The certificate template '{resolved.name}' does not contain required placeholder(s): {', '.join(missing)}"
            )
        return str(resolved)

    docx_files = find_docx_templates(root_dir)
    if not docx_files:
        raise FileNotFoundError(
            f"No DOCX certificate template found in current directory:\n{root_dir}"
        )

    # Inspect each docx file for required placeholders
    valid_templates = []
    for f in docx_files:
        is_valid, _ = inspect_docx_template(f)
        if is_valid:
            valid_templates.append(f)

    if not valid_templates:
        sample = docx_files[0]
        _, missing = inspect_docx_template(sample)
        raise ValueError(
            f"Found DOCX file '{Path(sample).name}', but it does not contain the required placeholder(s): {', '.join(missing)}.\n"
            "Expected placeholders: {{REG_NO}} and {{NAME}}"
        )

    if len(valid_templates) > 1:
        root_path = Path(root_dir).resolve()
        root_files = [f for f in valid_templates if Path(f).parent == root_path]
        if len(root_files) == 1:
            return root_files[0]

        file_list = "\n".join(f" - {f}" for f in valid_templates)
        raise RuntimeError(
            f"Multiple valid certificate templates found:\n{file_list}\n"
            "Please specify the certificate template using the --template argument."
        )

    return valid_templates[0]


def replace_placeholder_in_paragraph(
    p: Paragraph,
    target_pattern: re.Pattern,
    replacement: str,
    reduce_font_for_long_text: bool = False,
    font_size: Optional[float] = 20.0,
    bold: Optional[bool] = True
) -> bool:
    """
    Replace target_pattern in paragraph runs while strictly preserving font, style,
    and surrounding text.
    Handles placeholders split across multiple runs.
    Applies prominent font size and bold styling to ensure certificate readability.
    """
    match = target_pattern.search(p.text)
    if not match:
        return False

    replaced_any = False

    while True:
        match = target_pattern.search(p.text)
        if not match:
            break

        start_pos, end_pos = match.span()
        runs = p.runs
        if not runs:
            # Fallback if no runs
            p.text = target_pattern.sub(replacement, p.text, count=1)
            replaced_any = True
            break

        # Map character positions to runs
        curr = 0
        start_run_idx = -1
        start_run_offset = -1
        end_run_idx = -1
        end_run_offset = -1

        for idx, r in enumerate(runs):
            r_len = len(r.text)
            next_curr = curr + r_len
            if start_run_idx == -1 and curr <= start_pos < next_curr:
                start_run_idx = idx
                start_run_offset = start_pos - curr
            if curr < end_pos <= next_curr:
                end_run_idx = idx
                end_run_offset = end_pos - curr
                break
            curr = next_curr

        if start_run_idx == -1 or end_run_idx == -1:
            # Fallback: couldn't map indices cleanly, do run[0] replacement
            runs[0].text = target_pattern.sub(replacement, p.text, count=1)
            for r in runs[1:]:
                r.text = ""
            replaced_any = True
            break

        target_run = runs[start_run_idx]

        if start_run_idx == end_run_idx:
            # Entire placeholder is within a single run
            target_run.text = (
                target_run.text[:start_run_offset]
                + replacement
                + target_run.text[end_run_offset:]
            )
        else:
            # Placeholder spans multiple runs
            target_run.text = target_run.text[:start_run_offset] + replacement
            for mid_idx in range(start_run_idx + 1, end_run_idx):
                runs[mid_idx].text = ""
            runs[end_run_idx].text = runs[end_run_idx].text[end_run_offset:]

        # Apply prominent font styling if specified
        if font_size is not None:
            effective_size = font_size
            if reduce_font_for_long_text and len(replacement) > 28:
                if len(replacement) > 40:
                    effective_size = max(font_size * 0.72, 12.0)
                else:
                    effective_size = max(font_size * 0.85, 14.0)
            
            # Apply to all runs in this student detail paragraph if they lack size
            for r in runs:
                if r.font.size is None or r.font.size.pt < 14.0:
                    r.font.size = Pt(effective_size)
                if bold is not None and r.font.bold is None:
                    r.font.bold = bold
                if r.font.name is None:
                    r.font.name = "Book Antiqua"

        replaced_any = True

    return replaced_any


def adjust_font_size_for_long_name(
    run: docx.text.run.Run,
    text_length: int,
    base_size: float = 20.0
) -> None:
    """
    Intelligently reduce font size for unusually long student names so they fit
    neatly without breaking certificate layout.
    """
    if text_length > 40:
        new_pt = max(base_size * 0.72, 12.0)
    elif text_length > 28:
        new_pt = max(base_size * 0.85, 14.0)
    else:
        new_pt = base_size
    run.font.size = Pt(new_pt)


def create_individual_certificate_docx(
    template_path: str,
    output_docx_path: str,
    reg_no: str,
    name: str,
    font_size: Optional[float] = 20.0,
    bold: Optional[bool] = True
) -> None:
    """
    Load a fresh copy of the DOCX template, replace student placeholders,
    apply prominent font sizing, and save the output DOCX.
    The original template is NEVER modified.
    """
    doc = docx.Document(template_path)
    paragraphs = get_all_document_paragraphs(doc)

    for p in paragraphs:
        # Replace student name
        replace_placeholder_in_paragraph(
            p,
            NAME_REGEX,
            name,
            reduce_font_for_long_text=True,
            font_size=font_size,
            bold=bold
        )
        # Replace student registration number
        replace_placeholder_in_paragraph(
            p,
            REG_NO_REGEX,
            reg_no,
            reduce_font_for_long_text=False,
            font_size=font_size,
            bold=bold
        )

    doc.save(output_docx_path)
