# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['SP_emb+exp_beta.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['tkcalendar', 'mysql.connector', 'mysql.connector.plugins.mysql_native_password', 'mysql.connector.plugins.caching_sha2_password', 'openpyxl', 'os', 'pandas', 'babel.numbers'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='SP_emb+exp_beta',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
