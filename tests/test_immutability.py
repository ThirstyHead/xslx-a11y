"""Tests for document immutability and SHA-256 verification."""
import hashlib
from pathlib import Path
import pytest
from xlsx_a11y.immutability import (
    calculate_sha256,
    verify_immutability,
    get_remediated_path,
    assert_not_same_path,
)


def test_calculate_sha256(tmp_path: Path):
    sample = tmp_path / "test.xlsx"
    sample.write_bytes(b"dummy-excel-content-bytes")
    expected = hashlib.sha256(b"dummy-excel-content-bytes").hexdigest()
    assert calculate_sha256(sample) == expected


def test_get_remediated_path():
    in_path = Path("/path/to/Sample.xlsx")
    out_path = get_remediated_path(in_path)
    assert out_path == Path("/path/to/Sample-remediated.xlsx")

    # Explicit output
    explicit = Path("/somewhere/custom.xlsx")
    assert get_remediated_path(in_path, explicit) == explicit


def test_assert_not_same_path():
    p1 = Path("/path/to/file.xlsx")
    p2 = Path("/path/to/file.xlsx")
    with pytest.raises(ValueError, match="strictly guarantees that original files remain untouched"):
        assert_not_same_path(p1, p2)


def test_verify_immutability(tmp_path: Path):
    src = tmp_path / "original.xlsx"
    src.write_bytes(b"untouched-data")
    sha_before = calculate_sha256(src)

    # Should pass when unchanged
    assert verify_immutability(src, sha_before) is True

    # Should raise error when mutated
    src.write_bytes(b"mutated-data")
    with pytest.raises(RuntimeError, match="Original document was modified"):
        verify_immutability(src, sha_before)
