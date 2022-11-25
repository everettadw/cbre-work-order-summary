import os
import re


def get_file_count(path: str) -> int:
    """Returns the number of files in the given directory."""
    if not os.path.isdir(path):
        raise ValueError(f"Directory '{path}' does not exist.")
    return len([name for name in os.listdir(path) if os.path.isfile(os.path.join(path, name)) and name.find('crdownload') == -1 and name.find('.tmp') == -1])


def replace_escapes(string: str) -> str:
    """Escapes a string with tabs and newlines."""
    return re.sub('[\t\r\n]', '', string.strip())
