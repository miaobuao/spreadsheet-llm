"""
SpreadsheetLLM - A Python package for spreadsheet compression and LLM-based analysis.

This package provides tools for compressing spreadsheet data and using LLMs to extract
structured information from spreadsheets.
"""

from spreadsheet_llm.SpreadsheetLLMWrapper import (
    SpreadsheetLLMWrapper,
    CellRangeList,
    CellRangeItem,
    RECOGNIZE_PROMPT,
)
from spreadsheet_llm.SheetCompressor import SheetCompressor
from spreadsheet_llm.IndexColumnConverter import IndexColumnConverter
from spreadsheet_llm import CellRangeUtils

# Optional: Legacy SpreadsheetLLM class
try:
    from spreadsheet_llm.SpreadsheetLLM import SpreadsheetLLM
    __all_with_legacy__ = ["SpreadsheetLLM"]
except ImportError:
    __all_with_legacy__ = []

__version__ = "0.1.0"

__all__ = [
    "SpreadsheetLLMWrapper",
    "CellRangeList",
    "CellRangeItem",
    "RECOGNIZE_PROMPT",
    "SheetCompressor",
    "IndexColumnConverter",
    "CellRangeUtils",
] + __all_with_legacy__
