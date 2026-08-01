#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Test script for logging functionality in xlsxjinja.

This script demonstrates:
1. Logger initialization with debug mode
2. Logger behavior in normal mode vs debug mode
3. Integration with BookWriter
"""

import os
import sys

# Add parent directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from xlsxjinja import get_logger, setup_logger

print("=" * 60)
print("Testing xlsxjinja Logger")
print("=" * 60)

# Test 1: Logger in normal mode (no debug)
print("\n1. Testing logger in NORMAL mode (debug=False):")
print("-" * 60)
logger_normal = setup_logger("xlsxjinja_test1", debug=False)
logger_normal.debug("This DEBUG message should NOT appear")
logger_normal.info("This INFO message should NOT appear")
logger_normal.warning("This WARNING message SHOULD appear")
logger_normal.error("This ERROR message SHOULD appear")

# Test 2: Logger in debug mode
print("\n2. Testing logger in DEBUG mode (debug=True):")
print("-" * 60)
logger_debug = setup_logger("xlsxjinja_test2", debug=True)
logger_debug.debug("This DEBUG message SHOULD appear")
logger_debug.info("This INFO message SHOULD appear")
logger_debug.warning("This WARNING message SHOULD appear")
logger_debug.error("This ERROR message SHOULD appear")

# Test 3: BookWriter integration
print("\n3. Testing BookWriter with debug mode:")
print("-" * 60)

# Create a simple test template if it doesn't exist
template_file = "test_template.xlsx"
if not os.path.exists(template_file):
    print(f"Note: Template file '{template_file}' not found.")
    print("Skipping BookWriter integration test.")
else:
    from xlsxjinja import BookWriter

    print("Creating BookWriter with debug=True...")
    try:
        writer = BookWriter(template_file, debug=True)
        print("✓ BookWriter created successfully with debug logging enabled")
    except Exception as e:
        print(f"✗ Error creating BookWriter: {e}")

print("\n" + "=" * 60)
print("Logger Test Complete")
print("=" * 60)
print("\nSummary:")
print("- Normal mode: Only WARNING and ERROR messages shown")
print("- Debug mode: All messages (DEBUG, INFO, WARNING, ERROR) shown")
print("- Messages are colored for better visibility")
print("\nUsage in your code:")
print("  from xlsxjinja import BookWriter")
print("  writer = BookWriter('template.xlsx', debug=True)")
print("=" * 60)
