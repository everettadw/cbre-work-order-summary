import os

import tomlkit


class TOMLWrapper:
    def __init__(self, config_file_name):
        self.file_name = config_file_name
        self.file_path = os.path.join(os.getcwd(), config_file_name)

        with open(self.file_path, mode='rt', encoding='utf-8') as f:
            self.config = tomlkit.load(f)

    def get_dict(self):
        return self.config

    def update_dict(self, new_config_dict):
        self.config = new_config_dict

    def update_toml(self):
        with open(self.file_path, mode='wt', encoding='utf-8') as f:
            tomlkit.dump(self.config, f)
