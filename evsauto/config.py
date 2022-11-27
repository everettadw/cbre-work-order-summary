import tomlkit


class TOMLWrapper:
    def __init__(self, config_file_path):
        self.file_path = config_file_path

        with open(self.file_path, mode='rt', encoding='utf-8') as f:
            self.config = tomlkit.load(f)

    def get_dict(self):
        return self.config

    def update_dict(self, new_config_dict):
        self.config = new_config_dict

    def update_toml(self):
        with open(self.file_path, mode='wt', encoding='utf-8') as f:
            tomlkit.dump(self.config, f)
