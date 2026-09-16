"""
Logger Module for Student Certificate Generator.

Maintains generation_log.csv tracking every student's conversion status,
output file path, and any errors encountered during execution.
"""

import os
import csv
from pathlib import Path
from typing import Optional, Dict, Any, List


LOG_COLUMNS = ["SI NO", "REG NO", "NAME", "STATUS", "OUTPUT", "ERROR"]


class GenerationLogger:
    """Handles logging certificate generation results to CSV."""

    def __init__(self, output_dir: str, log_filename: str = "generation_log.csv"):
        self.output_dir = Path(output_dir).resolve()
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = self.output_dir / log_filename
        self.records: List[Dict[str, str]] = []
        self._init_csv()

    def _init_csv(self) -> None:
        """Create or overwrite the CSV file with headers."""
        with open(self.log_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(LOG_COLUMNS)

    def log_success(self, si_no: str, reg_no: str, name: str, output_filename: str) -> None:
        """Record a successful certificate generation."""
        record = {
            "SI NO": str(si_no),
            "REG NO": str(reg_no),
            "NAME": str(name),
            "STATUS": "SUCCESS",
            "OUTPUT": str(output_filename),
            "ERROR": ""
        }
        self.records.append(record)
        self._append_row(record)

    def log_failure(self, si_no: str, reg_no: str, name: str, error_message: str) -> None:
        """Record a failed certificate generation."""
        record = {
            "SI NO": str(si_no),
            "REG NO": str(reg_no),
            "NAME": str(name),
            "STATUS": "FAILED",
            "OUTPUT": "",
            "ERROR": str(error_message)
        }
        self.records.append(record)
        self._append_row(record)

    def _append_row(self, record: Dict[str, str]) -> None:
        """Append a single record row to the CSV file immediately."""
        with open(self.log_path, "a", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow([
                record["SI NO"],
                record["REG NO"],
                record["NAME"],
                record["STATUS"],
                record["OUTPUT"],
                record["ERROR"]
            ])

    @property
    def total_count(self) -> int:
        return len(self.records)

    @property
    def success_count(self) -> int:
        return sum(1 for r in self.records if r["STATUS"] == "SUCCESS")

    @property
    def failure_count(self) -> int:
        return sum(1 for r in self.records if r["STATUS"] == "FAILED")
