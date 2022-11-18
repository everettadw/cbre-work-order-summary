import os
import re
import sys


def get_file_count(path: str) -> int:
    """Returns the number of files in the given directory."""
    if not os.path.isdir(path):
        raise ValueError(f"Directory '{path}' does not exist.")
    return len([name for name in os.listdir(path) if os.path.isfile(os.path.join(path, name)) and name.find('crdownload') == -1 and name.find('.tmp') == -1])


def isin_production() -> bool:
    """
    Returns an int describing the environment of the application.

    Development -> False
    Production  -> True
    """
    if os.name == "nt" and os.path.join(
            os.path.basename(os.path.dirname(sys.executable)),
            os.path.basename(sys.executable)) == os.path.join("Scripts", "python.exe"):
        return False
    return True


def replace_escapes(string: str) -> str:
    """Escapes a string with tabs and newlines."""
    return re.sub('[\t\r\n]', '', string.strip())
