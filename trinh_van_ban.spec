# -*- mode: python ; coding: utf-8 -*-

import sys

a = Analysis(
    ['trinh_van_ban.py'],
    pathex=[],
    binaries=[],
    datas=[],
    # pystray/Pillow: chỉ dùng ở chế độ "chạy ngầm dưới khay" trên Windows. pystray nạp backend
    # theo tên lúc chạy nên phải khai báo tường minh cho PyInstaller. Trên macOS bỏ qua (import
    # lazy — thiếu 2 lib này chương trình vẫn chạy).
    hiddenimports=(
        ['pystray._win32', 'PIL.Image', 'PIL.ImageDraw', 'PIL.ImageFont']
        if sys.platform == 'win32' else []
    ),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='trinh_van_ban',
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

if sys.platform == "darwin":
    app = BUNDLE(
        exe,
        name="trinh_van_ban.app",
        icon=None,
        bundle_identifier="vn.gov.moh.qld.trinhvanban",
        info_plist={"NSHighResolutionCapable": True},
    )
