"""
Pytest configuration file for test discovery and setup.

This file adds the src directory to Python path so that imports work correctly.
"""

import sys
from pathlib import Path

# Add src directory to Python path for imports
src_path = Path(__file__).parent.parent / "src"
sys.path.insert(0, str(src_path))
