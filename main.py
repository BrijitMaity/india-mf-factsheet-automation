#!/usr/bin/env python
"""
================================================================================
MUTUAL FUND MASTER PIPELINE - UNIFIED SINGLE COMMAND ENTRYPOINT
================================================================================
Execute the entire pipeline end-to-end with:
    python main.py

Optional Flags:
    python main.py --date 2026-10-12    # Specific dated folder
    python main.py --excel-only         # Run only Excel data extraction
    python main.py --download-only      # Run only Download and validation
    python main.py --all                # Extract Excel across all date folders
================================================================================
"""
import sys
import os

from run_pipeline import main

if __name__ == "__main__":
    main()
