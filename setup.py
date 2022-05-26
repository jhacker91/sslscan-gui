"""
Uso:
    python setup.py py2app
"""

from setuptools import setup

APP = ['sslscangui.py']
DATA_FILES = ['icona_pri.icns','icona_pri.png','icona_pri.ico']
OPTIONS = {'argv_emulation':True, 'includes':['requests','tkinter','xlsxwriter'],'iconfile':'icona_pri.icns'}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)
