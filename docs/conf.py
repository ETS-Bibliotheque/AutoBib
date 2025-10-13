from __future__ import annotations

import os
import sys
from datetime import datetime

sys.path.insert(0, os.path.abspath(".."))

project = "AutoBib+"
author = "AutoBib+ Team"
copyright = f"{datetime.now():%Y}, AutoBib+ Team"

extensions = [
    "sphinx.ext.autodoc",
    "sphinx.ext.autosummary",
    "sphinx.ext.napoleon",
    "sphinx.ext.viewcode",
    "sphinx.ext.ifconfig",
    "sphinx_autodoc_typehints",
]

autosummary_generate = True
master_doc = "index"

autodoc_default_options = {
    "members": True,
    "undoc-members": False,
    "show-inheritance": True,
}

suppress_warnings = ["autodoc.mocked_object"]

autodoc_mock_imports = [
    "PySide6",
    "PySide6.QtWidgets",
    "PySide6.QtGui",
    "PySide6.QtCore",
    "PySide6_Addons",
    "PySide6_Essentials",
    "shiboken6",
    "win32gui",
    "win32com",
    "win32com.client",
    "win32api",
    "win32con",
    "win32clipboard",
    "pandas",
    "numpy",
    "openpyxl",
    "xlsxwriter",
    "requests",
    "pywin32",
    "pywin32_ctypes",
    "pyinstaller",
    "pyinstaller_hooks_contrib",
    "altgraph",
    "pillow",
    "pefile",
    "Levenshtein",
    "RapidFuzz",
    "fuzzywuzzy",
    "tqdm",
    "docx",
    "lxml",
    "pytz",
    "tzdata",
    "Unidecode",
    "colorama",
    "charset_normalizer",
    "certifi",
    "idna",
    "urllib3",
    "Include.pybliometrics",
]

napoleon_google_docstring = False
napoleon_numpy_docstring = True

templates_path = ["_templates"]
exclude_patterns: list[str] = ["_build"]

html_theme = "sphinx_rtd_theme"
html_static_path = ["_static"]
