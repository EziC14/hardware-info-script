rmdir /s /q build dist __pycache__
del script.spec
python -m PyInstaller --onefile script.py
