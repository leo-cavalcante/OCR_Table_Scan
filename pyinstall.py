# https://pyinstaller.org/en/stable/usage.html

import PyInstaller.__main__

PyInstaller.__main__.run([
    'main.py',
    # '--onefile',
    '--windowed',
    '--noconsole',
    '--icon=appicon.icns'
])