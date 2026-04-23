"""
Файл для сборки .app через py2app
Использование:
    pip install py2app
    python setup_mac.py py2app
Готовый .app появится в папке dist/
"""
from setuptools import setup

APP = ['convert_1144_gui.py']
DATA_FILES = []
OPTIONS = {
    'argv_emulation': False,
    'packages': ['pandas', 'openpyxl', 'tkinter'],
    'includes': ['tkinter', '_tkinter'],
    'iconfile': None,
    'plist': {
        'CFBundleName': 'JDE Конвертер',
        'CFBundleDisplayName': 'JDE Конвертер',
        'CFBundleVersion': '1.0.0',
        'CFBundleShortVersionString': '1.0',
        'NSHighResolutionCapable': True,
        'LSMinimumSystemVersion': '12.0',
    },
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
