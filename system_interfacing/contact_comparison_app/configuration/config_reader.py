import importlib.util
import os


class ConfigReader:
    def __init__(self, config_file_path="config.py"):

        if not os.path.exists(config_file_path):
            raise FileNotFoundError(f"Config file not found: {config_file_path}")

        spec = importlib.util.spec_from_file_location("config", config_file_path)
        if spec is None:
            raise ImportError(f"Could not load config file: {config_file_path}")

        config_module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(config_module)

        for key, value in config_module.__dict__.items():
            if not key.startswith("_"):  # Ignore private/special attributes
                setattr(self, key, value)

    def __repr__(self):
        attrs = {k: v for k, v in self.__dict__.items() if not k.startswith("_")}
        return f"ConfigReader({attrs})"
