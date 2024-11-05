from setuptools import setup
import py2exe

py2exe.freeze(
    windows=['convert2.py'],
    options={
        'py2exe': {
            'includes': ['TkEasyGUI', 'jpbizday', 'chardet', 'pandas', 'openpyxl'],
            'compressed' : True,
            'bundle_files' : 1
        }
    }
)