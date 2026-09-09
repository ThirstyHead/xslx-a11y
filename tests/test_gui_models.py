"""Tests for GUI batch models."""
from pathlib import Path
import openpyxl
from xlsx_a11y.gui.models import BatchItem, BatchQueue


def test_batch_queue_add_and_deduplicate(tmp_path: Path):
    f1 = tmp_path / "Book1.xlsx"
    f2 = tmp_path / "Book2.xlsx"
    temp_lock = tmp_path / "~$Book1.xlsx"

    wb = openpyxl.Workbook()
    wb.save(f1)
    wb.save(f2)
    temp_lock.write_bytes(b"lock")

    queue = BatchQueue()
    item1 = queue.add_file(f1)
    assert item1 is not None
    assert len(queue) == 1

    # Duplicate should not be added
    assert queue.add_file(f1) is None
    assert len(queue) == 1

    # Lock file ignored
    assert queue.add_file(temp_lock) is None
    assert len(queue) == 1

    # Directory addition
    added = queue.add_directory(tmp_path)
    assert added == 1  # only Book2 was new
    assert len(queue) == 2

    queue.clear()
    assert len(queue) == 0
