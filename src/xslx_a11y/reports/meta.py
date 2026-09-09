"""Canonical W3C WCAG 2.1 metadata, Understanding URLs, and POUR taxonomy for Excel.

Delegates to engine_a11y.reports.meta.
"""
from engine_a11y.reports.meta import (
    POUR_INTROS,
    PRINCIPLES,
    SC_META,
    W3C_QUICKREF,
    W3C_UNDERSTANDING_BASE,
)

__all__ = [
    "W3C_UNDERSTANDING_BASE",
    "W3C_QUICKREF",
    "SC_META",
    "PRINCIPLES",
    "POUR_INTROS",
]
