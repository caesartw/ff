from setuptools import setup

APP = ['main_vfinal.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': True,
    'packages': ['openpyxl', 'pandas'],
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
