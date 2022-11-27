import os
import re
import sys


def get_file_count(path: str) -> int:
    """Returns the number of files in the given directory."""
    if not os.path.isdir(path):
        raise ValueError(f"Directory '{path}' does not exist.")
    return len([name for name in os.listdir(path) if os.path.isfile(os.path.join(path, name)) and name.find('crdownload') == -1 and name.find('.tmp') == -1])


def replace_escapes(string: str) -> str:
    """Escapes a string with tabs and newlines."""
    return re.sub('[\t\r\n]', '', string.strip())


def is_executable() -> bool:
    """Returns true if the code is running from inside an .exe file."""
    if getattr(sys, 'frozen', False):
        return True
    return False
