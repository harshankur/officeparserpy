#!/usr/bin/env python3

"""test_generator.py

This module is responsible for generating test files and their content for the office parser.

Usage:
    python test_generator.py <extension>

Arguments:
    extension: The file extension for which the test content file will be generated.

Supported Extensions:
    - docx
    - pptx
    - xlsx
    - odt
    - odp
    - ods
    - pdf
"""

import sys
from officeparserpy import parse_office
from supported_extensions import supported_extensions

def get_filename(ext, is_content_file=False):
    """
    Generate the filename for a given extension.

    Args:
        ext (str): The file extension.
        is_content_file (bool): If True, generates the filename for the content file.

    Returns:
        str: The generated filename.
    """
    return f"test/files/test.{ext}" + (".txt" if is_content_file else "")

def create_content_file(ext):
    """
    Create the content file for a given extension using the office parser.

    Args:
        ext (str): The file extension.
    """
    text = parse_office(get_filename(ext), config={"preserve_temp_files": True})
    with open(get_filename(ext, True), "w", encoding="utf-8") as content_file:
        content_file.write(text)

if len(sys.argv) != 2:
    print("Usage: test_generator.py <extension>")
    sys.exit(1)

requested_extension = sys.argv[1]
if requested_extension in supported_extensions:
    create_content_file(requested_extension)
    print(f"Created text content file for {requested_extension} => {get_filename(requested_extension, True)}")
else:
    print("The requested extension test is not currently available.")
