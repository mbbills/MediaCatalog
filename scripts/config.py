import configparser
from pathlib import Path


# config.py lives in MediaCatalog/scripts/, so the project root is one
# directory above the scripts folder.
PROJECT_ROOT = Path(__file__).resolve().parent.parent
SETTINGS_FILE = PROJECT_ROOT / "settings.ini"


def load_settings():
    """Load settings.ini from the MediaCatalog project root."""
    if not SETTINGS_FILE.exists():
        raise FileNotFoundError(
            f"settings.ini was not found at: {SETTINGS_FILE}"
        )

    config = configparser.ConfigParser()
    config.read(SETTINGS_FILE, encoding="utf-8")

    return config


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
