import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

block_cipher = None

hidden = []
hidden += collect_submodules('flask')
hidden += collect_submodules('pyomo')
hidden += collect_submodules('engineio')
hidden += collect_submodules('socketio')
hidden += ['pkg_resources', 'packaging']

a = Analysis(
    ['app.py'],
    pathex=['.'],
    binaries=[
        ('solvers/glpsol.exe',    'solvers'),
    ],
    datas=[
        ('templates',                       'templates'),
        ('dashboard_mix30dias_ort.html',    '.'),
        ('sequenciamento_v2_ort.html',      '.'),
        ('dashboard_data_ort.json',         '.'),
        ('sequenciamento_ort30dias_data.json', '.'),
        ('Simulador - Mix de cartões_3__2_.xlsx', '.'),
        ('setup_mp27.xlsx',                 '.'),
        ('setup_mp28.xlsx',                 '.'),
    ],
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['tkinter', 'matplotlib', 'PyQt5', 'PyQt6'],
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
    name='PlanejamentoORT',
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
    icon=None,
)
