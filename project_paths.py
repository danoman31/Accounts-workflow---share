"""Paths shared by the scripts, independent of the current working directory."""
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPTS_FOLDER = PROJECT_ROOT / 'scripts'
OUTPUTS_FOLDER = PROJECT_ROOT / 'outputs'
