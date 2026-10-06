"""Local-only configuration for the synthetic portfolio dataset."""

from pathlib import Path
import os
import re
import sys

ROOT = Path(__file__).resolve().parents[1]
DATA = ROOT / 'demo_data'
FOLDERS = (
    '01-Assemblies',
    '02-Motion Systems',
    '03-Mounting Hardware',
)
PROFILES = {
    'demo': {
        'label': 'DEMO',
        'bom_root': str(DATA / 'files' / 'boms'),
        'drawings_root': str(DATA / 'files' / 'drawings'),
        'math_root': str(DATA / 'files' / 'cad'),
        'dfmea_root': str(DATA / 'files' / 'dfmea'),
        'proto_bom_root': str(DATA / 'files' / 'prototype_boms'),
        'proto_drawings_root': str(DATA / 'files' / 'prototype_drawings'),
        'proto_math_root': str(DATA / 'files' / 'prototype_cad'),
        'prototype_index': str(DATA / 'index' / 'prototype'),
        'index': str(DATA / 'index' / 'combined_data.feather'),
        'pdf_index': str(DATA / 'index' / 'pdf_data.feather'),
        'cells': str(DATA / 'index' / 'bom_cells.feather'),
        'files': str(DATA / 'index' / 'bom_files.feather'),
        'pdfs': str(DATA / 'index' / 'bom_pdfs.feather'),
        'pcr_root': str(DATA / 'files' / 'change_records'),
        'projects': str(DATA / 'index' / 'change_projects.feather'),
        'projects_active': str(DATA / 'files' / 'projects' / 'active'),
        'projects_closed': str(DATA / 'files' / 'projects' / 'closed'),
        'project_folders': str(DATA / 'index' / 'project_folders.feather'),
        'dfmea': str(DATA / 'index' / 'dfmea_files.feather'),
        'usage_dir': str(DATA / 'usage'),
        'feedback_dir': str(DATA / 'feedback'),
    }
}
STALE_INDEX_DAYS = 3650
UPDATE_CHECK_HOURS = 24
SHARE_APP = ''
APP_EXE = 'BOM Search Demo.exe'
APP_EXE_PATTERN = 'BOM Search Demo*.exe'
APP_ICON = 'demo.ico'
HELP_DOCUMENT = 'USER_GUIDE.md'
ECN_FORM_PATTERN = 'Not included in the public demo'

DRAWING_FOLDER_RENAMES = {}
CUSTOMER_PRINT_FOLDERS = {}
DEFAULT_CUSTOMER_PRINT_FOLDER = 'Customer Prints'
PLURAL_NAMES = {}


def active_name():
    name = os.environ.get('DEMO_PROFILE', 'demo')
    if name != 'demo':
        raise SystemExit('This portfolio build supports the demo profile only.')
    return name


def paths():
    return PROFILES['demo']


def is_live():
    return False


def warn_when_index_is_old():
    return False


def describe_v2():
    return 'DEMO profile - synthetic local data'


def drawings_folder_for(name):
    return DRAWING_FOLDER_RENAMES.get(name, name)


def customer_print_folder_for(name):
    return CUSTOMER_PRINT_FOLDERS.get(name, DEFAULT_CUSTOMER_PRINT_FOLDER)


def commodity_name(name):
    text = re.sub(r'^\s*\d+\s*[-_ ]\s*', '', str(name)).strip()
    return text[:1].upper() + text[1:] if text else str(name)


def resolve_table(path):
    return (path, False) if path and os.path.isfile(path) else ('', False)


def obsolete_index():
    return str(DATA / 'index' / 'obsolete_boms.feather')


def network_ready(path):
    return os.path.isdir(os.path.dirname(os.path.abspath(path)))


def usage_log():
    return str(DATA / 'usage' / 'demo-user.txt')


def feedback_file():
    return str(DATA / 'feedback' / 'demo-user.txt')


def feedback_targets():
    return [feedback_file()]


def app_icon():
    return ''


def help_document():
    candidate = ROOT / 'docs' / HELP_DOCUMENT
    return str(candidate) if candidate.is_file() else ''


def ecn_form():
    return ''


def published_exe():
    return None


def running_exe():
    return sys.executable if getattr(sys, 'frozen', False) else None
