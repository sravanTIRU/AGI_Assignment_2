"""
Validation Module for Student Certificate Generator.

Validates Excel structure, per-student rows, duplicate registration numbers,
and certificate template placeholders with human-readable error reporting.
"""

from typing import List, Dict, Any, Tuple
from collections import defaultdict


class ValidationError(Exception):
    """Raised when validation fails."""
    pass


def validate_student_records(
    students: List[Dict[str, Any]]
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Separate students into valid and invalid records.
    Returns:
        (valid_students, invalid_students)
    """
    valid_students: List[Dict[str, Any]] = []
    invalid_students: List[Dict[str, Any]] = []

    for s in students:
        if s.get("is_valid", False):
            valid_students.append(s)
        else:
            invalid_students.append(s)

    return valid_students, invalid_students


def check_duplicate_registration_numbers(
    students: List[Dict[str, Any]]
) -> Dict[str, List[int]]:
    """
    Check for duplicate REG NO values among student records.
    Returns a dict mapping duplicate REG NO -> list of 1-based Excel row numbers.
    """
    reg_map = defaultdict(list)
    for s in students:
        reg_no = s.get("reg_no", "").strip()
        if reg_no:
            reg_map[reg_no].append(s.get("row_num", 0))

    duplicates = {reg: rows for reg, rows in reg_map.items() if len(rows) > 1}
    return duplicates


def format_duplicate_error(duplicates: Dict[str, List[int]]) -> str:
    """Format duplicate REG NO error message."""
    lines = ["Duplicate REG NO detected:\n"]
    for reg_no, rows in duplicates.items():
        lines.append(f"{reg_no}\n")
        lines.append("Rows:")
        for r in rows:
            lines.append(f"{r}")
        lines.append("")
    return "\n".join(lines).strip()


def format_invalid_student_row(student: Dict[str, Any]) -> str:
    """Format individual invalid row error message."""
    row_num = student.get("row_num", "?")
    reg_no = student.get("reg_no", "").strip() or "EMPTY"
    name = student.get("name", "").strip() or "EMPTY"
    reason = student.get("error", "Unknown validation error")

    return (
        f"Row {row_num}:\n"
        f"REG NO: {reg_no}\n"
        f"NAME: {name}\n\n"
        f"Status: INVALID\n"
        f"Reason: {reason}"
    )
