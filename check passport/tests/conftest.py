"""Shared test helpers: import path + builders for fake Excels and photos."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

import pytest
from openpyxl import Workbook
from PIL import Image


HEADERS = ["Index", "Name", "Surname", "Passport no"]


def make_excel(path, rows, headers=HEADERS):
    """Write a small .xlsx list. rows = list of tuples matching headers."""
    wb = Workbook()
    ws = wb.active
    ws.append(headers)
    for r in rows:
        ws.append(list(r))
    wb.save(path)
    return path


def make_photo(path, size=(400, 300), color="white"):
    """Write a small real JPEG/PNG so Pillow can open and stamp it."""
    path.parent.mkdir(parents=True, exist_ok=True)
    Image.new("RGB", size, color).save(path)
    return path


@pytest.fixture
def test_config():
    """A config dict mirroring config.json, for calling run_check directly."""
    return {
        "excel_file": "auto",
        "name_column": "Name",
        "surname_column": "Surname",
        "id_column": "Passport no",
        "index_column": "Index",
        "images_folder": "images",
        "output_folder": "output",
        "skip_pdf": True,
        "number_output_files": True,
        "id_text": {
            "prefix": "ID: ",
            "show_index": True,
            "index_prefix": "No ",
            "position": "bottom-right",
            "margin_x": 30,
            "margin_y": 30,
            "font_size": 48,
            "font_color": "#000000",
            "box": True,
            "box_color": "#FFFFFF",
            "box_opacity": 200,
            "box_padding": 14,
        },
    }
