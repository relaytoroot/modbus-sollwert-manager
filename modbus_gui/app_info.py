from __future__ import annotations

import sys
from pathlib import Path


APP_NAME = "FGH Modbus Sollwert Manager"
APP_VERSION = "1.0.0"
APP_COMPANY = "FGH"
APP_DESCRIPTION = "Werkzeug zum Planen und Ausfuehren von Modbus-Sollwertablaeufen"
APP_AUTHOR = "Made by Yunus Sevgi"
ICON_FILE = "FGH_Logo_prüflabor_gruen.ico"
HEADER_LOGO_FILE = "FGH_Logo_gruen.ico"
COMBO_ARROW_LIGHT_THEME_FILE = "combo_arrow_dark.svg"
COMBO_ARROW_DARK_THEME_FILE = "combo_arrow_light.svg"
SPIN_ARROW_UP_LIGHT_THEME_FILE = "spin_arrow_up_dark.svg"
SPIN_ARROW_UP_DARK_THEME_FILE = "spin_arrow_up_light.svg"
SPIN_ARROW_DOWN_LIGHT_THEME_FILE = "spin_arrow_down_dark.svg"
SPIN_ARROW_DOWN_DARK_THEME_FILE = "spin_arrow_down_light.svg"
CHECKBOX_CHECK_LIGHT_THEME_FILE = "checkbox_check_dark.svg"
CHECKBOX_CHECK_DARK_THEME_FILE = "checkbox_check_light.svg"


def project_root() -> Path:
    return Path(__file__).resolve().parent.parent


def resource_path(filename: str) -> Path:
    base_path = Path(getattr(sys, "_MEIPASS", project_root()))
    return base_path / filename

