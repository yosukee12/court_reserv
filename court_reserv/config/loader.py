# -*- coding: utf-8 -*-
import configparser
import os
from functools import lru_cache
from pathlib import Path


PACKAGE_DIR = Path(__file__).resolve().parents[1]
REPO_ROOT = PACKAGE_DIR.parent
PACKAGE_CONFIG_PATH = PACKAGE_DIR / "config.ini"
LOCAL_CONFIG_CANDIDATES = (
    REPO_ROOT / "config.local.ini",
    PACKAGE_DIR / "config.local.ini",
)
ENV_FILE_CANDIDATES = (
    REPO_ROOT / ".env",
    PACKAGE_DIR / ".env",
)

ENV_TO_CONFIG_MAP = {
    "COURT_RESERV_CHROME_DRIVER_PATH": ("PATH", "DRIVER_PATH"),
    "COURT_RESERV_DRIVER_PATH": ("PATH", "DRIVER_PATH"),
    "COURT_RESERV_LOG_PATH": ("PATH", "LOG_PATH"),
    "COURT_RESERV_OUTPUT_CSV_PATH": ("PATH", "OUTPUT_CSV_PATH"),
    "COURT_RESERV_TOP_URL": ("URL", "TOP_URL"),
    "COURT_RESERV_LOG_LEVEL": ("LOG", "LEVEL"),
}


def _read_env_file():
    env_values = {}
    for env_path in ENV_FILE_CANDIDATES:
        if not env_path.exists():
            continue
        for line in env_path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, value = stripped.split("=", 1)
            env_values[key.strip()] = value.strip().strip("'\"")
    return env_values


@lru_cache(maxsize=1)
def load_config():
    config = configparser.ConfigParser()
    if not PACKAGE_CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"config.ini not found at {PACKAGE_CONFIG_PATH}. "
            "Create court_reserv/config.ini from config.example.ini."
        )

    config.read(PACKAGE_CONFIG_PATH, encoding="utf-8")

    existing_local_configs = [str(path) for path in LOCAL_CONFIG_CANDIDATES if path.exists()]
    if existing_local_configs:
        config.read(existing_local_configs, encoding="utf-8")

    env_values = _read_env_file()
    for env_key, (section, option) in ENV_TO_CONFIG_MAP.items():
        env_value = os.environ.get(env_key, env_values.get(env_key))
        if not env_value:
            continue
        if not config.has_section(section):
            config.add_section(section)
        config.set(section, option, env_value)

    if not config.has_section("LOG"):
        config.add_section("LOG")
    if not config.has_option("LOG", "LEVEL"):
        config.set("LOG", "LEVEL", "INFO")

    if not config.has_section("AUTH"):
        config.add_section("AUTH")

    return config


def get_default_credentials():
    env_values = _read_env_file()
    config = load_config()

    user_id = os.environ.get("COURT_RESERV_USER_ID", env_values.get("COURT_RESERV_USER_ID", "")).strip()
    password = os.environ.get("COURT_RESERV_PASSWORD", env_values.get("COURT_RESERV_PASSWORD", "")).strip()

    if not user_id and config.has_option("AUTH", "USER_ID"):
        user_id = config.get("AUTH", "USER_ID", fallback="").strip()
    if not password and config.has_option("AUTH", "PASSWORD"):
        password = config.get("AUTH", "PASSWORD", fallback="").strip()

    return user_id, password


def get_output_base_path():
    config = load_config()
    output_path = config.get("PATH", "OUTPUT_CSV_PATH", fallback="").strip()
    if output_path:
        return Path(output_path)
    return REPO_ROOT / "output"


def get_debug_output_dir():
    return REPO_ROOT / "output" / "debug_pages"
