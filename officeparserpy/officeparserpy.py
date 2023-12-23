#!/usr/bin/env python3

"""
Module providing a function to parse office files.
Currently supports docx, pptx, xlsx, odt, odp, ods, pdf files.
"""

# Standard imports
from xml.dom.minidom import parseString
from typing import Dict, List, Union, Literal
from io import BytesIO
import re
import os
import sys
import shutil
import zipfile
import time
# External imports
from pdfminer.high_level import extract_text as extract_text_from_pdf
import filetype



################################################### Zip Extractor ###################################################
def extract_files_with_regex(zip_path: str, extract_path: str, regex_pattern: str) -> List[str]:
    """
    Extract files from a ZIP archive based on a regex pattern.

    Args:
        zip_path (str): Path to the ZIP archive.
        extract_path (str): Directory where the files will be extracted.
        regex_pattern (str): Regular expression pattern to match filenames.

    Returns:
        List[str]: List of file names that were extracted.

    Raises:
        zipfile.BadZipFile: If the specified file (`zip_path`) is not a valid ZIP archive.
    """
    extracted_files = []
    
    try:
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            matching_files = [file for file in zip_ref.namelist() if re.search(regex_pattern, file)]

            for file in matching_files:
                zip_ref.extract(file, extract_path)
                extracted_files.append(file)

    except Exception as e:
        raise FileCorrupted(zip_path) from e

    return extracted_files

def display_extracted_files(extract_path: str) -> None:
    """
    Display the paths of the extracted files.

    Args:
        extract_path (str): Directory where the files were extracted.

    Returns:
        None

    Raises:
        FileNotFoundError: If the specified directory (`extract_path`) does not exist.
    """
    if not os.path.exists(extract_path):
        raise FileNotFoundError(f"The specified directory '{extract_path}' does not exist.")

    print("Extracted Files:")
    for root, _, files in os.walk(extract_path):
        for file in files:
            print(os.path.join(root, file))
#####################################################################################################################
###################################################### Config #######################################################
# Define a set of allowed keys
AllowedKeys = Literal[
    'temp_files_location',
    'preserve_temp_files',
    'output_error_to_console',
    'newline_delimiter',
    'ignore_notes',
    'put_notes_at_last',
]

# Define the type of the config variable
OfficeParserConfig = Dict[AllowedKeys, Union[bool, str, None]]

__all__ = [
    'OfficeParserConfig',
    'get_temp_files_location',
    'get_preserve_temp_files',
    'get_output_error_to_console',
    'get_newline_delimiter',
    'get_ignore_notes',
    'get_put_notes_at_last',
]

default_temp_files_location = 'officeparser_temp'

def get_temp_files_location(config: OfficeParserConfig) -> str:
    """Get the 'temp_files_location' property with a default value."""
    return config.get('temp_files_location', default_temp_files_location)


def get_preserve_temp_files(config: OfficeParserConfig) -> bool:
    """Get the 'preserve_temp_files' property with a default value."""
    return config.get('preserve_temp_files', False)


def get_output_error_to_console(config: OfficeParserConfig) -> bool:
    """Get the 'output_error_to_console' property with a default value."""
    return config.get('output_error_to_console', False)


def get_newline_delimiter(config: OfficeParserConfig) -> str:
    """Get the 'newline_delimiter' property with a default value."""
    return config.get('newline_delimiter', '\n')


def get_ignore_notes(config: OfficeParserConfig) -> bool:
    """Get the 'ignore_notes' property with a default value."""
    return config.get('ignore_notes', False)


def get_put_notes_at_last(config: OfficeParserConfig) -> bool:
    """Get the 'put_notes_at_last' property with a default value."""
    return config.get('put_notes_at_last', False)

#####################################################################################################################
#################################################### File Utils #####################################################
GLOBAL_FILE_NAME_ITERATOR = 0

def get_new_file_name(temp_files_location: str, ext: str) -> str:
    """
    File Name generator that takes the extension as an input and returns a file name that comprises a timestamp and an incrementing number
    to allow the files to be sorted in chronological order.

    Args:
        temp_files_location (str): Directory where this new file needs to be stored.
        ext (str): File extension for this new generated file name.

    Returns:
        str: The generated file name.
    """
    global GLOBAL_FILE_NAME_ITERATOR    # Declare that we are using the global variable

    # Get the iterator part of the file name
    iterator_part = str(GLOBAL_FILE_NAME_ITERATOR).zfill(5)
    GLOBAL_FILE_NAME_ITERATOR = (GLOBAL_FILE_NAME_ITERATOR + 1) % 100000

    # Return the file name
    return f"{temp_files_location}/tempfiles/{int(time.time())}{iterator_part}.{ext}"

def read_bytes_from_file(file_path):
    """
    Read the bytes of a file in binary mode.

    Args:
        file_path (str): The path to the file.

    Returns:
        bytes: The bytes read from the file.
               Returns None if the file is not found or if there's an error reading the file.

    Raises:
        FileNotFoundError: If the specified file is not found.
        Exception: If there's an error reading the file.

    Example:
        >>> file_path = 'path/to/your/file.txt'
        >>> file_bytes = read_bytes_from_file(file_path)
        >>> if file_bytes is not None:
        ...     print(f"Bytes read from {file_path}: {file_bytes}")
    """
    try:
        with open(file_path, 'rb') as file:
            file_bytes = file.read()
        return file_bytes
    except FileNotFoundError:
        print(f"Error: File not found - {file_path}")
        return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None



def get_file_extension_from_bytes(file_content: bytes) -> str:
    """
    Identify the file type based on magic bytes and return the corresponding extension.

    Args:
        file_content (bytes): The content of the file as bytes.

    Returns:
        str or None: The file extension corresponding to the identified file type.
                     Returns None if the type is not recognized.
    """
    # Identify the file type based on magic bytes
    file_info = filetype.guess(file_content)

    # Map identified types to specific extensions
    type_to_extension = {
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/pdf': 'pdf',
        'application/vnd.oasis.opendocument.text': 'odt',
        'application/vnd.oasis.opendocument.presentation': 'odp',
        'application/vnd.oasis.opendocument.spreadsheet': 'ods',
    }

    # Check if the identified type is in the mapping
    if file_info and file_info.mime in type_to_extension:
        return type_to_extension[file_info.mime]

    # If no match, return None
    return None
#####################################################################################################################
################################################# Custom Exceptions #################################################
ERROR_HEADER = '[officeparserpy]: '

class FileCorrupted(Exception):
    """
    Exception raised for a corrupted file.

    Attributes:
        filepath -- The path of the file that triggered the exception.
        message  -- A customized error message.
    """
    def __init__(self, filepath):
        self.value = filepath
        self.message = f"Your file {filepath} seems to be corrupted. If you are sure it is fine, please create a ticket in Issues on github with the file to reproduce error."
        super().__init__(self.message)

class ExtensionUnsupported(Exception):
    """
    Exception raised for an unsupported file extension.

    Attributes:
        ext      -- The unsupported file extension.
        message  -- A customized error message.
    """
    def __init__(self, ext):
        self.value = ext
        self.message = f"Sorry, officeparser currently supports docx, pptx, xlsx, odt, odp, ods, pdf files only. Create a ticket in Issues on github to add support for {ext} files. Stay tuned for further updates."
        super().__init__(self.message)

class FileDoesNotExist(Exception):
    """
    Exception raised for a non-existent file.

    Attributes:
        filepath -- The path of the non-existent file.
        message  -- A customized error message.
    """
    def __init__(self, filepath):
        self.value = filepath
        self.message = f"File {filepath} could not be found! Check if the file exists or verify if the relative path to the file is correct from your terminal's location."
        super().__init__(self.message)

class LocationNotFound(Exception):
    """
    Exception raised for an unreachable directory location.

    Attributes:
        location -- The unreachable directory location.
        message  -- A customized error message.
    """
    def __init__(self, location):
        self.value = location
        self.message = f"Entered location {location} is not reachable! Please make sure that the entered directory location exists. Check relative paths and reenter."
        super().__init__(self.message)

class ImproperArguments(Exception):
    """Exception raised for improper function arguments."""
    def __init__(self):
        self.message = "Improper arguments"
        super().__init__(self.message)

class ImproperBuffers(Exception):
    """Exception raised for errors while reading file buffers."""
    def __init__(self):
        self.message = "Error occurred while reading the file buffers"
        super().__init__(self.message)
#####################################################################################################################

def _parse_word(filepath: str, config: OfficeParserConfig) -> str:
    """
    This function parses word files and returns the parsed text.
    It decides a few configurations of the parsing using the config object passed in the argument.

    Args:
        filepath: The path of the docx file that needs to be parsed.
        config: A config dictionary for parsing text from this file.

    Returns:
        The parsed text.

    Raises:
        FileCorrupted: If the file represented in the parsed file path is not in the expected format.
    """

    # The target content xml files for the docx file
    target_files_regex = r"word\/(document|footnotes|endnotes)\.xml"

    # The decompress location which contains the filename in it.
    decompress_location = f"{get_temp_files_location(config)}/{filepath.split('/').pop()}"

    # Decompress docx files into the target xml files.
    extracted_files = extract_files_with_regex(zip_path=filepath, extract_path=decompress_location,
                                                             regex_pattern=target_files_regex)

    # Verify if atleast the document xml file exists in the extracted files list.
    # Otherwise, raise FileCorrupted error
    if 'word/document.xml' not in extracted_files:
        raise FileCorrupted(filepath)

    # List of all file content
    xml_content_array = []

    # Read values in each file into an array.
    try:
        for local_file_path in extracted_files:
            with open(file=f"{decompress_location}/{local_file_path}", mode='r', encoding='utf8') as file_content:
                # Read the entire content of the file
                xml_content_array.append(file_content.read())

    except Exception as e:
        raise FileCorrupted(filepath) from e

    # ************************************* word xml files explanation *************************************
    # Structure of xmlContent of a word file is simple.
    # All text nodes are within w:t tags and each of the text nodes that belong in one paragraph are clubbed together within a w:p tag.
    # So, we will filter out all the empty w:p tags and then combine all the w:t tag text inside for creating our response text.
    # ******************************************************************************************************

    # Holds the response text
    response_text = []

    # Iterate over the content array, get dom of the xml content of each file and fetch the text info from within
    for xml_content in xml_content_array:
        # Get iterable list of w:p elements
        xml_paragraph_nodes_list = list(parseString(xml_content).getElementsByTagName('w:p'))
        # Iterate over w:p elements to extract the w:t text from within
        paragraph_text = []
        for paragraph_node in xml_paragraph_nodes_list:
            text_node_list = list(paragraph_node.getElementsByTagName('w:t'))
            if text_node_list:
                paragraph_text.append(''.join(text_node.firstChild.nodeValue for text_node in text_node_list))
        response_text.append(get_newline_delimiter(config).join(paragraph_text))

    # Join all response_text array
    response_text = get_newline_delimiter(config).join(response_text)

    # Return the response text
    return response_text


def _parse_powerpoint(filepath: str, config: OfficeParserConfig) -> str:
    """
    This function parses PowerPoint files and returns the parsed text.
    It decides a few configurations of the parsing using the config object passed in the argument.

    Args:
        filepath: The path of the pptx file that needs to be parsed.
        config: A config dictionary for parsing text from this file.

    Returns:
        The parsed text.

    Raises:
        FileCorrupted: If the file represented in the parsed file path is not in the expected format.
    """

    # Files regex that hold our content of interest
    all_files_regex = r"ppt/(notesSlides|slides)/(notesSlide|slide)\d+.xml"
    slides_regex = r"ppt/slides/slide\d+.xml"

    # The decompress location which contains the filename in it
    decompress_location = f"{get_temp_files_location(config)}/{filepath.split('/').pop()}"

    # Decompress pptx files into the target xml files.
    extracted_files = extract_files_with_regex(zip_path=filepath, extract_path=decompress_location,
                                                             regex_pattern=all_files_regex if not get_ignore_notes(
                                                                 config) else slides_regex)

    # Verify if atleast the slides xml files exist in the extracted files list.
    # Otherwise, raise FileCorrupted error
    if not any(re.match(slides_regex, filename) for filename in extracted_files):
        raise FileCorrupted(filepath)

    # Check if any sorting is required.
    if not get_ignore_notes(config) and get_put_notes_at_last(config):
        # Sort files according to previous order of taking text out of ppt/slides followed by ppt/notesSlides
        # For this, we are looking at the presence of notes string in the file name. If it exists, it goes to the end.
        extracted_files.sort(key=lambda x: (1 if 'notes' in x else 0, x.find('notes') if 'notes' in x else len(x)))

    # Returning an array of all the xml contents read using fs.readFileSync
    xml_content_array = []

    # Read values in each file into an array.
    try:
        for local_file_path in extracted_files:
            with open(file=f"{decompress_location}/{local_file_path}", mode='r', encoding='utf8') as file_content:
                # Read the entire content of the file
                xml_content_array.append(file_content.read())

    except Exception as e:
        raise FileCorrupted(filepath) from e

    # ******************************** powerpoint xml files explanation ************************************
    # Structure of xmlContent of a powerpoint file is simple.
    # There are multiple xml files for each slide and correspondingly their notesSlide files.
    # All text nodes are within a:t tags and each of the text nodes that belong in one paragraph are clubbed together within a a:p tag.
    # So, we will filter out all the empty a:p tags and then combine all the a:t tag text inside for creating our response text.
    # ******************************************************************************************************

    # Holds the response text
    response_text = []

    # Iterate over the content array, get dom of the xml content of each file and fetch the text info from within
    for xml_content in xml_content_array:
        # Get iterable list of a:p elements
        xml_paragraph_nodes_list = list(parseString(xml_content).getElementsByTagName('a:p'))
        # Iterate over a:p elements to extract the a:t text from within
        paragraph_text = []
        for paragraph_node in xml_paragraph_nodes_list:
            text_node_list = list(paragraph_node.getElementsByTagName('a:t'))
            if text_node_list:
                paragraph_text.append(''.join(text_node.firstChild.nodeValue for text_node in text_node_list))
        response_text.append(get_newline_delimiter(config).join(paragraph_text))

    # Join all response_text array
    response_text = get_newline_delimiter(config).join(response_text)

    # Return the response text
    return response_text


def _parse_excel(filepath: str, config: OfficeParserConfig) -> str:
    """
    This function parses Excel files and returns the parsed text.
    It decides a few configurations of the parsing using the config object passed in the argument.

    Args:
        filepath: The path of the xlsx file that needs to be parsed.
        config: A config dictionary for parsing text from this file.

    Returns:
        The parsed text.

    Raises:
        FileCorrupted: If the file represented in the parsed file path is not in the expected format.
    """

    # Files regex that hold our content of interest
    sheets_regex = r"xl/worksheets/sheet\d+.xml"
    drawings_regex = r"xl/drawings/drawing\d+.xml"
    charts_regex = r"xl/charts/chart\d+.xml"
    strings_file_path = 'xl/sharedStrings.xml'

    # The decompress location which contains the filename in it
    decompress_location = f"{get_temp_files_location(config)}/{filepath.split('/').pop()}"

    # Decompress xlsx files into the target xml files.
    extracted_files = extract_files_with_regex(zip_path=filepath, extract_path=decompress_location,
                                                             regex_pattern=f"{sheets_regex}|{drawings_regex}|{charts_regex}|{strings_file_path}")

    # Verify if atleast the slides xml files exist in the extracted files list.
    # Otherwise, raise FileCorrupted error
    if not any(re.match(sheets_regex, filename) for filename in extracted_files):
        raise FileCorrupted(filepath)

    # Read values in each file into an object.
    xml_content_files_object: Dict[str, List[str]] = {
        'sheet_files': [],
        'drawing_files': [],
        'chart_files': [],
        'shared_strings_file': ''
    }

    try:
        for local_file_path in extracted_files:
            with open(file=f"{decompress_location}/{local_file_path}", mode='r', encoding='utf8') as file_content:
                # Read the entire content of the file
                content = file_content.read()
                if local_file_path.startswith('xl/worksheets'):
                    xml_content_files_object['sheet_files'].append(content)
                elif local_file_path.startswith('xl/drawings'):
                    xml_content_files_object['drawing_files'].append(content)
                elif local_file_path.startswith('xl/charts'):
                    xml_content_files_object['chart_files'].append(content)
                elif local_file_path == strings_file_path:
                    xml_content_files_object['shared_strings_file'] = content

    except Exception as e:
        raise FileCorrupted(filepath) from e

    # ********************************** excel xml files explanation ***************************************
    # Structure of xmlContent of an excel file is a bit complex.
    # We have a sharedStrings.xml file which has strings inside t tags
    # Each sheet has an individual sheet xml file which has numbers in v tags (probably value) inside c tags (probably cell)
    # Each value of v tag is to be used as it is if the 't' attribute (probably type) of c tag is not 's' (probably shared string)
    # If the 't' attribute of c tag is 's', then we use the value to select value from sharedStrings array with the value as its index.
    # Drawing files contain all text for each drawing and have text nodes in a:t and paragraph nodes in a:p.
    # ******************************************************************************************************

    # Holds the response text
    response_text = []

    # Find text nodes with t tags in sharedStrings xml file
    shared_strings_xml_t_nodes_list = list(
        parseString(xml_content_files_object['shared_strings_file']).getElementsByTagName('t'))
    # Create shared string array. This will be used as a map to get strings from within sheet files.
    shared_strings = [t_node.childNodes[0].nodeValue for t_node in shared_strings_xml_t_nodes_list]

    # Parse Sheet files
    for sheet_xml_content in xml_content_files_object['sheet_files']:
        # Find text nodes with c tags in sharedStrings xml file
        sheets_xml_c_nodes_list = list(parseString(sheet_xml_content).getElementsByTagName('c'))
        # Traverse through the nodes list and fill response_text with either the number value in its v node or find a mapped string from shared_strings.
        response_text.append(
            get_newline_delimiter(config).join([
                (shared_strings[
                     int(c_node.getElementsByTagName('v')[0].childNodes[0].nodeValue)] if c_node.getAttribute(
                    't') == 's' else c_node.getElementsByTagName('v')[0].childNodes[0].nodeValue)
                for c_node in sheets_xml_c_nodes_list if c_node.getElementsByTagName('v')
            ])
        )

    # Parse Drawing files
    for drawing_xml_content in xml_content_files_object['drawing_files']:
        # Find text nodes with a:p tags
        drawings_xml_paragraph_nodes_list = list(parseString(drawing_xml_content).getElementsByTagName('a:p'))
        # Store all the text content to respond
        response_text.append(
            get_newline_delimiter(config).join([
                ''.join([
                    text_node.childNodes[0].nodeValue
                    for text_node in paragraph_node.getElementsByTagName('a:t')
                ])
                for paragraph_node in drawings_xml_paragraph_nodes_list if paragraph_node.getElementsByTagName('a:t')
            ])
        )

    # Parse Chart files
    for chart_xml_content in xml_content_files_object['chart_files']:
        # Find text nodes with c:v tags
        charts_xml_cv_nodes_list = list(parseString(chart_xml_content).getElementsByTagName('c:v'))
        # Store all the text content to respond
        response_text.append(get_newline_delimiter(config).join([c_v_node.childNodes[0].nodeValue for c_v_node in charts_xml_cv_nodes_list]))

    # Join all response_text array
    response_text = get_newline_delimiter(config).join(response_text)

    # Return the response text
    return response_text


def _parse_open_office(filepath: str, config: OfficeParserConfig) -> str:
    """
    This function parses OpenOffice files and returns the parsed text.
    It decides a few configurations of the parsing using the config object passed in the argument.

    Args:
        filepath: The path of the ods file that needs to be parsed.
        config: A config dictionary for parsing text from this file.

    Returns:
        The parsed text.

    Raises:
        FileCorrupted: If the file represented in the parsed file path is not in the expected format.
    """

    # The target content xml file for the OpenOffice file
    main_content_file_path = 'content.xml'
    object_content_files_regex = r"Object \d+/content.xml"

    # The decompress location which contains the filename in it
    decompress_location = f"{get_temp_files_location(config)}/{filepath.split('/').pop()}"

    # Decompress ods files into the target xml files.
    extracted_files = extract_files_with_regex(zip_path=filepath, extract_path=decompress_location,
                                                             regex_pattern=f"{main_content_file_path}|{object_content_files_regex}")

    # Verify if atleast the content xml file exists in the extracted files list.
    # Otherwise, raise FileCorrupted error
    if main_content_file_path not in extracted_files:
        raise FileCorrupted(filepath)

    # Read values in each file into an object.
    xml_content_files_object: Dict[str, List[str]] = {
        'main_content_file': '',
        'object_content_files': []
    }

    try:
        for local_file_path in extracted_files:
            with open(file=f"{decompress_location}/{local_file_path}", mode='r', encoding='utf8') as file_content:
                # Read the entire content of the file
                content = file_content.read()
                if local_file_path == main_content_file_path:
                    xml_content_files_object['main_content_file'] = content
                elif local_file_path.startswith("Object"):
                    xml_content_files_object['object_content_files'].append(content)

    except Exception as e:
        raise FileCorrupted(filepath) from e

    # ********************************** openoffice xml files explanation **********************************
    # Structure of xmlContent of OpenOffice files is simple.
    # All text nodes are within text:h and text:p tags with all kinds of formatting within nested tags.
    # All text in these tags are separated by new line delimiters.
    # Objects like charts in ods files are in Object d+/content.xml with the same way as above.
    # ******************************************************************************************************

    # Holds the response text
    response_text = []
    # Holds the notes text
    notes_text = []

    # List of allowed text tags
    allowed_text_tags = ["text:p", "text:h"]
    # Notes tag
    notes_tag = "presentation:notes"

    # Main dfs traversal function that goes from one node to its children and returns the value out.
    def extract_all_texts_from_node(root):
        xml_text_array = []
        for _, child_node in enumerate(root.childNodes):
            traversal(child_node, xml_text_array, True)
        return "".join(xml_text_array)

    # Traversal function that gets recursive calling.
    def traversal(node, xml_text_array, is_first_recursion):
        if not node.childNodes or len(node.childNodes) == 0:
            parent_node = node.parentNode if node.parentNode else None
            if parent_node and getattr(parent_node, 'tagName', '').startswith('text') and node.nodeValue:
                if is_notes_node(parent_node) and (
                        get_put_notes_at_last(config) or get_ignore_notes(
                        config)):
                    notes_text.append(node.nodeValue)
                    if getattr(parent_node, 'tagName', '') in allowed_text_tags and not is_first_recursion:
                        notes_text.append(get_newline_delimiter(config) or "\n")
                else:
                    xml_text_array.append(node.nodeValue)
                    if getattr(parent_node, 'tagName', '') in allowed_text_tags and not is_first_recursion:
                        xml_text_array.append(get_newline_delimiter(config) or "\n")
            return

        for _, child_node in enumerate(node.childNodes):
            traversal(child_node, xml_text_array, False)

    # Checks if the given node has an ancestor which is a notes tag. We use this information to put the notes in the response text and its position.
    def is_notes_node(node):
        if hasattr(node, 'tagName') and node.tagName == notes_tag:
            return True
        if node.parentNode:
            return is_notes_node(node.parentNode)
        return False

    # Checks if the given node has an ancestor which is also an allowed text tag. In that case, we ignore the child text tag.
    def is_invalid_text_node(node):
        if hasattr(node, 'tagName') and getattr(node, 'tagName', '') in allowed_text_tags:
            return True
        if node.parentNode:
            return is_invalid_text_node(node.parentNode)
        return False

    # The xml string parsed as xml array
    xml_content_array = [parseString(xml_content_files_object['main_content_file']),
                         *map(parseString, xml_content_files_object['object_content_files'])]

    # Iterate over each xml_content and extract text from them.
    for xml_content in xml_content_array:
        # Find text nodes with text:h and text:p tags in xml_content
        xml_text_nodes_list = [node for node in list(xml_content.getElementsByTagName("*")) if
                               hasattr(node, 'tagName') and getattr(node, 'tagName',
                                                                    '') in allowed_text_tags and not is_invalid_text_node(
                                   node.parentNode)]
        # Store all the non-empty text content to respond
        non_empty_texts = [text for text_node in xml_text_nodes_list if
                           (text := extract_all_texts_from_node(text_node))]
        if non_empty_texts:
            response_text.append(get_newline_delimiter(config).join(non_empty_texts))

    # Add notes text at the end if the user config says so.
    # Note that we already have pushed the text content to notes_text array while extracting all texts from the nodes.
    if not get_ignore_notes(config) and get_put_notes_at_last(config):
        response_text.extend(notes_text)

    # Join all response_text array
    response_text = get_newline_delimiter(config).join(response_text)

    # Return the response text
    return response_text


def _parse_pdf(filepath: str, config: OfficeParserConfig) -> str:
    """
    This function parses PDF files and returns the parsed text.
    It can also replace newline characters with the specified delimiter in the config object.

    Args:
        filepath: The path of the PDF file that needs to be parsed.
        config: A config dictionary for parsing text from this file.

    Returns:
        The parsed text.

    Raises:
        FileCorrupted: If the file represented in the parsed file path is not in the expected format.
    """
    # Read PDF content
    with open(filepath, "rb") as file:
        pdf_data = BytesIO(file.read())

    try:
        # Extract text using pdfminer.six
        response_text = extract_text_from_pdf(pdf_data)

        # Replace newline characters if specified in the config
        if get_newline_delimiter(config) and get_newline_delimiter(
                config) != "\n":
            response_text = response_text.replace("\n", get_newline_delimiter(config))
    # Handle any error from pdf parsing as a file corrupted error.
    except Exception as e:
        raise FileCorrupted(filepath) from e

    return response_text


def parse_office(file: Union[str, bytes], config: OfficeParserConfig = None) -> str:
    """
    Parse the content of an office file (docx, pptx, xlsx, odt, odp, ods, pdf) and return the text content.

    Args:
        file (Union[str, bytes]): The file path or buffer of the office file.
        config (OfficeParserConfig, optional): Configuration options. Defaults to {}.

    Returns:
        str: The parsed text content.

    Raises:
        ExtensionUnsupported: If the file extension is not supported.
        FileCorrupted: If the file is corrupted.
        FileDoesNotExist: If the file does not exist.
        ImproperBuffers: If the buffers are not proper.
    """
    # Create internal config as a copy of the config passed in argument
    internal_config = config.copy() if config else {}
    # Check if temp_files_location is passed in the config
    if internal_config.get('temp_files_location') is not None:
        internal_config[
            'temp_files_location'] = f"{get_temp_files_location(internal_config)}{'' if get_temp_files_location(internal_config).endswith('/') else '/'}{default_temp_files_location}"

    try:
        # Create temp file subdirectory if it does not exist
        os.makedirs(os.path.join(get_temp_files_location(internal_config), 'tempfiles'),
                    exist_ok=True)

        # New file path for the file passed in argument
        new_file_path = ''
        # Check if buffer
        if isinstance(file, bytes):
            # Guess file type from buffer
            file_type = get_file_extension_from_bytes(file)
            if file_type:
                # temp file name
                new_file_path = get_new_file_name(
                    get_temp_files_location(internal_config), file_type)
                # write new file
                with open(new_file_path, 'wb') as new_file:
                    new_file.write(file)
            else:
                raise ImproperBuffers
        # Not buffers but real file path.
        else:
            # Check if file exists
            if not os.path.exists(file):
                raise FileDoesNotExist(file)

            # temp file name
            new_file_path = get_new_file_name(get_temp_files_location(internal_config),
                                                         file.split('.').pop().lower())
            # Copy the file into a temp location with the temp name
            shutil.copy2(file, new_file_path)

        # File extension. Already in lowercase when we prepared the temp file above.
        extension = new_file_path.split(".").pop()

        # Response text
        response_text = ''
        # Switch between parsing functions depending on extension.
        if extension == 'docx':
            response_text = _parse_word(new_file_path, internal_config)
        elif extension == 'pptx':
            response_text = _parse_powerpoint(new_file_path, internal_config)
        elif extension == 'xlsx':
            response_text = _parse_excel(new_file_path, internal_config)
        elif extension in ['odt', 'odp', 'ods']:
            response_text = _parse_open_office(new_file_path, internal_config)
        elif extension == 'pdf':
            response_text = _parse_pdf(new_file_path, internal_config)
        else:
            raise ExtensionUnsupported(extension)

        # Check if we need to preserve unzipped content files or delete them.
        if not get_preserve_temp_files(internal_config):
            # Delete decompress sublocation
            shutil.rmtree(get_temp_files_location(internal_config))

        return response_text
    # Handle custom exceptions by printing its errors if requested to do so in config.
    except (ExtensionUnsupported, FileCorrupted, FileDoesNotExist,
            ImproperBuffers) as e:
        if get_output_error_to_console(internal_config):
            print(ERROR_HEADER + e.message)
        raise e


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: officeparser.py <argfilepath>")
        sys.exit(1)

    argfilepath = sys.argv[1]
    # Call your parsing function with the provided argfilepath
    print(parse_office(argfilepath))
