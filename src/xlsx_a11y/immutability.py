"""Document immutability and SHA-256 integrity verification.

Guarantees that input workbooks are never modified in place, maintaining
complete non-destructive audit and remediation guarantees.
"""
import hashlib
from pathlib import Path
from typing import Optional, Union


def calculate_sha256(path: Union[str, Path]) -> str:
    """Calculates the hex-encoded SHA-256 digest of a file."""
    p = Path(path)
    h = hashlib.sha256()
    with open(p, "rb") as f:
        while chunk := f.read(65536):
            h.update(chunk)
    return h.hexdigest()


def get_remediated_path(input_path: Union[str, Path], out_path: Optional[Union[str, Path]] = None) -> Path:
    """Determines the output path for a remediated workbook."""
    in_p = Path(input_path).resolve()
    if out_path:
        out_p = Path(out_path).resolve()
    else:
        stem = in_p.stem
        suffix = in_p.suffix or ".xlsx"
        out_p = in_p.with_name(f"{stem}-remediated{suffix}")
    return out_p


def assert_not_same_path(src: Union[str, Path], dest: Union[str, Path]) -> None:
    """Asserts that the destination path is not the exact same file as the source."""
    src_p = Path(src).resolve()
    dest_p = Path(dest).resolve()
    if src_p == dest_p:
        raise ValueError(
            f"Destination {dest_p} matches source {src_p}; xlsx-a11y strictly guarantees "
            "that original files remain untouched. Specify a distinct output path."
        )


def verify_immutability(path: Union[str, Path], expected_sha256: str) -> bool:
    """Verifies that the target file has not changed from its original SHA-256 digest."""
    current_sha256 = calculate_sha256(path)
    if current_sha256 != expected_sha256:
        raise RuntimeError(
            f"Original document was modified! Expected SHA-256 {expected_sha256}, "
            f"but found {current_sha256}."
        )
    return True
