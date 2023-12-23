# officeparserpy
A Python library to parse text out of any office file.

### Supported File Types

- [`docx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`pptx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`xlsx`](https://en.wikipedia.org/wiki/Office_Open_XML)
- [`odt`](https://en.wikipedia.org/wiki/OpenDocument)
- [`odp`](https://en.wikipedia.org/wiki/OpenDocument)
- [`ods`](https://en.wikipedia.org/wiki/OpenDocument)
- [`pdf`](https://en.wikipedia.org/wiki/PDF)

## Install via pip

```bash
pip install officeparserpy
```

## Library Usage
```python
from officeparserpy import parse_office

# USING FILE BUFFERS
# instead of file path, you can also pass file buffers of one of the supported files
# on parse_office function.

# get file buffers
file_buffers = open("/path/to/officeFile", "rb").read()
# get parsed text from officeparserpy
# NOTE: Only works with parse_office. Private functions are not supported.
data = parse_office(file_buffers)
print(data)
```

### Configuration Object: OfficeParserConfig

*Optionally add a config object as a parameter to parse_office for the following configurations*

| Flag                 | DataType | Default          | Explanation                                                                                                                                                                                                                                     |
|----------------------|----------|------------------|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| temp_files_location  | string   | officeparser_temp| The directory where officeparserpy stores the temp files. The final decompressed data will be put inside the officeparser_temp folder within your directory. **Please ensure that this directory actually exists.** Default is officeparser_temp. |
| preserve_temp_files  | boolean  | False            | Flag to not delete the internal content files and the possible duplicate temp files that it uses after unzipping office files. Default is False. It always deletes all of those files.                                                          |
| output_error_to_console | boolean  | False            | Flag to show all the logs to the console in case of an error. Default is False.                                                                                                                                                                     |
| newline_delimiter     | string   | '\n'             | The delimiter used for every new line in places that allow multiline text like word. Default is '\n'.                                                                                                                                             |
| ignore_notes          | boolean  | False            | Flag to ignore notes from parsing in files like PowerPoint. Default is False. It includes notes in the parsed text by default.                                                                                                                  |
| put_notes_at_last       | boolean  | False            | Flag, if set to True, will collectively put all the parsed text from notes at last in files like PowerPoint. Default is False. It puts each note right after its main slide content. If ignore_notes is set to True, this flag is also ignored. |
<br>

## Exception Types

`officeparserpy` can raise the following exceptions:

| Exception Type          | Description                                                        |
|--------------------------|--------------------------------------------------------------------|
| FileCorrupted            | Raised when the file is corrupted.                                 |
| ExtensionUnsupported     | Raised when the file extension is unsupported.                     |
| FileDoesNotExist         | Raised when the specified file does not exist.                     |
| LocationNotFound         | Raised when the specified directory location is not reachable.     |
| ImproperArguments         | Raised for improper function arguments.                            |
| ImproperBuffers           | Raised for errors while reading file buffers.                      |
<br>

**Example**
```python
from officeparserpy import parse_office, FileCorrupted, FileDoesNotExist

config = {
    'newline_delimiter': ' ',  # Separate new lines with a space instead of the default '\n'.
    'ignore_notes': True       # Ignore notes while parsing presentation files like pptx or odp.
}

try:
    # relative path is also fine => eg: files/myWorkSheet.ods
    data = parse_office("/Users/harsh/Desktop/files/mySlides.pptx", config)
    new_text = data + " look, I can parse a PowerPoint file"
    call_some_other_function(new_text)

    # Search for a term in the parsed text.
    def search_for_term_in_office_file(search_term, file_path):
        data = parse_office(file_path, config)
        return search_term in data

except FileDoesNotExist as file_not_found_error:
    print(f"Error: {file_not_found_error}")
    # Handle the case where the specified file does not exist.

except FileCorrupted as file_corrupted_error:
    print(f"Error: {file_corrupted_error}")
    # Handle the case where the file is corrupted.

except Exception as generic_error:
    print(f"An unexpected error occurred: {str(generic_error)}")
    # Handle other unexpected errors.

```


## Known Bugs
1. Inconsistency and incorrectness in the positioning of footnotes and endnotes in .docx files where the footnotes and endnotes would end up at the end of the parsed text, whereas it would be positioned exactly after the referenced word in .odt files.
2. The charts and objects information of .odt files are not accurate and may end up showing a few NaN in some cases.
----------

**pip**
https://pypi.org/project/officeparserpy/

**github**
https://github.com/harshankur/officeparserpy