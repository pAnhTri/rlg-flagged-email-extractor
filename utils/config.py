import configparser
import os

config = configparser.ConfigParser()

# Check if config file exists
if os.path.exists("config.ini"):
    config.read("config.ini")
    print("Config sections:", config.sections())
    if "Folder" in config and "output_folder" in config["Folder"]:
        print("Output folder:", config["Folder"]["output_folder"])
    else:
        print(
            "Warning: 'Folder' section or 'output_folder' key not found in config.ini"
        )
else:
    print("Warning: config.ini file not found")


def get_config(section, key):
    if not os.path.exists("config.ini"):
        raise FileNotFoundError("config.ini file not found")

    if section not in config:
        raise KeyError(f"Section '{section}' not found in config.ini")

    if key not in config[section]:
        raise KeyError(f"Key '{key}' not found in section '{section}'")

    return config[section][key]


def set_config(section, key, value):
    if not os.path.exists("config.ini"):
        raise FileNotFoundError("config.ini file not found")

    if section not in config:
        config[section] = {}

    config[section][key] = value

    with open("config.ini", "w") as configfile:
        config.write(configfile)
