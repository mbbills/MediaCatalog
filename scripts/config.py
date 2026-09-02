import configparser
from pathlib import Path


# config.py lives in MediaCatalog/scripts/, so the project root is one
# directory above the scripts folder.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = PROJECT_ROOT / "settings.ini"
SETTINGS_EXAMPLE_FILE = PROJECT_ROOT / "settings.example.ini"


def load_settings_files(defaults_path, settings_path=None):
    """Load shipped defaults, then overlay any user settings that exist."""
    defaults_path = Path(defaults_path)
    if not defaults_path.exists():
        raise FileNotFoundError(
            f"Default settings were not found at: {defaults_path}"
        )

    config = configparser.ConfigParser()
    config.read(defaults_path, encoding="utf-8")

    if settings_path is not None:
        settings_path = Path(settings_path)
        if settings_path.exists():
            config.read(settings_path, encoding="utf-8")

    return config


def load_settings():
    """Load shipped defaults and overlay settings.ini when present."""
    return load_settings_files(SETTINGS_EXAMPLE_FILE, SETTINGS_FILE)


def resolve_path(value):
    """
    Resolve a settings.ini path.

    Relative paths are interpreted from the MediaCatalog project root.
    Absolute paths are accepted unchanged.
    """
    path = Path(value).expanduser()

    if not path.is_absolute():
        path = PROJECT_ROOT / path

    return path.resolve()
