"""Excel Diff public API."""
from .core import CellChange, DiffReport, ExcelDiffError, compare, load_spreadsheet

__all__ = ["CellChange", "DiffReport", "ExcelDiffError", "compare", "load_spreadsheet"]
__version__ = "1.0.0"
