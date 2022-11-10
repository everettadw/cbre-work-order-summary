import os


def get_file_count(path):
    return len([name for name in os.listdir(path) if os.path.isfile(os.path.join(path, name)) and name.find('crdownload') == -1 and name.find('.tmp') == -1])
