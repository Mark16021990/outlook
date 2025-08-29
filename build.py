import PyInstaller.__main__
import os
import shutil

# Очистка предыдущих сборок
if os.path.exists('build'):
    shutil.rmtree('build')
if os.path.exists('dist'):
    shutil.rmtree('dist')

# Полный путь к main.py
main_script = os.path.join(os.path.dirname(__file__), 'main.py')

PyInstaller.__main__.run([
    main_script,
    '--onefile',
    '--windowed',
    '--name=EmailNotifier',
    '--hidden-import=imapclient',
    '--hidden-import=PIL',
    '--hidden-import=customtkinter',
    '--noconfirm',
    '--clean'
])