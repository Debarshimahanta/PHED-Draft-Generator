from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ASSETS = ROOT / 'assets'
TEMPLATES = ROOT / 'templates'
OUTPUTS = ROOT / 'outputs'

TEMPLATE_ID = 'CE_1'
TEMPLATE_NAME = 'CE_1 (Office Letter)'

OFFICE_HEADER = [
    'GOVERNMENT OF ASSAM',
    'OFFICE OF THE CHIEF ENGINEER (PHE) WATER, ASSAM',
    'HENGRABARI, GUWAHATI - 36',
    'Email - asphe@rediffmail.com',
]

SIGNATORY_LINES = [
    'Chief Engineer (PHE) Water,',
    'Hengrabari, Assam',
]

COPY_CATEGORIES = {
    'Info & Action': 'for favour of information and necessary action.',
    'Kind Information': 'for favour of kind information.',
    'PS Intimation': None,
    'Custom': None,
}
