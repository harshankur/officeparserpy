#!/usr/bin/env python3


"""test_runner.py

This module is responsible for running tests on the office parser.

Usage:
    python test_runner.py
    python test_runner.py <extension>

Arguments:
    extension: Optional. The file extension for which the specific test will be run.

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
from officeparserpy.officeparserpy import ERROR_HEADER, ExtensionUnsupported, FileCorrupted, FileDoesNotExist, ImproperBuffers, get_output_error_to_console, parse_office, OfficeParserConfig
from supported_extensions import supported_extensions

# List of all supported extensions with office Parser
supportedExtensionTests = [
    {
        'ext': 'docx',
        'testAvailable': True
    },
    {
        'ext': 'xlsx',
        'testAvailable': True
    },
    {
        'ext': 'pptx',
        'testAvailable': True
    },
    {
        'ext': 'odt',
        'testAvailable': True
    },
    {
        'ext': 'odp',
        'testAvailable': True
    },
    {
        'ext': 'ods',
        'testAvailable': True
    },
    {
        'ext': 'pdf',
        'testAvailable': True
    },
]

# Config file for performing tests
config: OfficeParserConfig = {
    'preserve_temp_files': True,
    'output_error_to_console': True,
}

# Local list of supported extensions in the test file
local_supported_extensions_list = [test['ext'] for test in supportedExtensionTests]

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

def run_test(ext):
    """
    Run a test for a given extension.

    Args:
        ext (str): The file extension.
    """
    try:
        text = parse_office(get_filename(ext), config)
        expected_text = ''
        with open(get_filename(ext, True), 'r', encoding='utf-8') as file:
            expected_text = file.read()

        if text == expected_text:
            print(f"[{ext}]=> Passed")
        else:
            print(f"[{ext}]=> Failed")
    except (ExtensionUnsupported, FileCorrupted, FileDoesNotExist,
            ImproperBuffers) as e:
        if get_output_error_to_console(config):
            print(ERROR_HEADER + e.message)

def run_all_tests():
    """Run all available tests."""
    for test in supportedExtensionTests:
        if test['testAvailable']:
            run_test(test['ext'])
        else:
            print(f"[{test['ext']}]=> Skipped")

if len(sys.argv) != 1 and len(sys.argv) != 2:
    print("Usage: test_runner.py")
    print("Usage: test_runner.py <extension>")
    sys.exit(1)

# Run all test files with test content if no argument passed.
if len(sys.argv) == 1:
    # Test to check all items in the local extension list are present in supportedExtensions.py file
    if all(ext in supported_extensions for ext in local_supported_extensions_list):
        print('All extensions in test files found in the primary supportedExtensions.py file')
    else:
        print('Extension in test files missing from the primary supportedExtensions.py file')

    # Test to check all items in supportedExtensions.py file are present in the local extension list
    if all(ext in local_supported_extensions_list for ext in supported_extensions):
        print('All extensions in the primary supportedExtensions.py file found in the test file')
    else:
        print('Extension in the primary supportedExtensions.py file missing from the test file')

    run_all_tests()
elif len(sys.argv) == 2:
    if sys.argv[1] in local_supported_extensions_list:
        text = parse_office(get_filename(sys.argv[1]), config)
        print(text)
    else:
        print('The requested extension test is not currently available.')
