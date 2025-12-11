#!/usr/bin/env python3
"""
BANTMS Cash Flow Workbook Generator
====================================

Command-line script to generate the complete Cash Flow Management Excel workbook.

Usage:
    python generate_workbook.py [options]

Options:
    -o, --output PATH    Output file path
    --no-sample-data     Generate without sample data

Example:
    python generate_workbook.py -o cashflow.xlsx
"""

import sys
import os

# Add src to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from src.generators.workbook import main

if __name__ == "__main__":
    main()
