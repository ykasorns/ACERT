# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['app_launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('AC Certs for print.pdf', '.'),
        ('NCSA-16March.pdf',       '.'),
        ('NCSA-19March.pdf',       '.'),
        ('fonts',                  'fonts'),
        ('templates',              'templates'),
    ],
    hiddenimports=[
        'flask',
        'jinja2',
        'werkzeug',
        'pypdf',
        'reportlab',
        'reportlab.pdfbase.ttfonts',
        'reportlab.pdfgen.canvas',
        'pandas',
        'openpyxl',
        'openpyxl.cell._writer',
        'et_xmlfile',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['gunicorn', 'tkinter', 'matplotlib', 'scipy', 'numpy.testing'],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ACERT',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,      # no black console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,          # ใส่ path ไปยัง .ico ถ้ามี เช่น icon='icon.ico'
)
