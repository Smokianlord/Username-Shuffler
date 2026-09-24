# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['Username-Shuffler.pyw'],
    pathex=[],
    binaries=[],
    datas=[
        ('icon.png', '.'),
        ('icon.ico', '.'),
        ('titlebar.ico', '.')
    ],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'numpy', 'scipy', 'pandas', 'matplotlib', 'torch',
        'tensorflow', 'onnxruntime', 'selenium', 'playwright',
        'trio', 'anyio', 'pytest', 'unittest', 'IPython', 'jupyter',
        'rembg', 'av', 'faster_whisper', 'ctranslate2'
    ],
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Username-Shuffler',
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
    icon=['icon.ico'],
)
