import os
import sys

# Application Info
APP_NAME = "Kunhwa PDF Editor"
VERSION = "v3.5"
AUTHOR = "TaeWoong Jang"
COMPANY = "Kunhwa Engineering & Consulting"

# Paths
def get_app_dir():
    """Returns the application directory."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

APP_DIR = get_app_dir()
DATA_DIR = os.path.join(APP_DIR, "data")
USERS_FILE = os.path.join(APP_DIR, "users.json.enc")

# UI Settings
THEME_NAME = "flatly"  # readable, modern, professional
FONT_FAMILY = "맑은 고딕"
FONT_SIZE_NORMAL = 10
FONT_SIZE_LARGE = 12
FONT_SIZE_HEADER = 14

# Colors (Fallback if theme issues)
COLOR_PRIMARY = "#2563EB"
COLOR_SECONDARY = "#6C757D"
COLOR_SUCCESS = "#10B981"
COLOR_WARNING = "#F59E0B"
COLOR_DANGER = "#EF4444"
