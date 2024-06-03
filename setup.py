from setuptools import setup

APP = ['main.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['openpyxl', 'numpy', 'pandas'],
    'includes': [
        'openpyxl',
        'numpy',
        'pandas',
        'openpyxl.cell',
        'openpyxl.cell._writer',
        'openpyxl.worksheet._writer'
    ],
    'excludes': [],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
