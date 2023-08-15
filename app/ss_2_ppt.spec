# -*- mode: python ; coding: utf-8 -*-


block_cipher = None

import os
print(os.getcwd())


a = Analysis(
    ['ss_2_ppt.py'],
    pathex=['C:\\Users\\ri003\\Documents\\Programming\\PowerPoint\\', 'C:\\Users\\ri003\\Documents\\Programming\\GUI\\'],
    binaries=[],
    datas=[],
    hiddenimports=["../../GUI/lib", "app", "lib", "lib.powerpoint", "../../GUI.lib.get_window_info"],
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
a.datas += [('python.png', '..\\image\\python.png', 'DATA')]
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='ss_2_ppt',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
