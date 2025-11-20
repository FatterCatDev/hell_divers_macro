# -*- mode: python ; coding: utf-8 -*-
import sys
from pathlib import Path

# Resolve this spec file path even when __file__ is not set.
_spec_path = Path(__file__) if "__file__" in globals() else Path(sys.argv[0])
project_root = _spec_path.resolve().parents[1]
src_dir = project_root / "src"
package_dir = src_dir / "hell_divers_macro"
data_dir = project_root / "data"

asset_dir = package_dir / "assets"
datas = []
if asset_dir.exists():
    datas += [
        (
            str(path),
            str(Path("hell_divers_macro") / "assets" / path.relative_to(asset_dir).parent),
        )
        for path in asset_dir.rglob("*")
        if path.is_file()
    ]

md_file = data_dir / "helldivers2_stratagem_codes.md"
if md_file.exists():
    datas.append((str(md_file), "data"))

a = Analysis(
    [str(package_dir / "main.py")],
    pathex=[str(src_dir)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
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
    name="helldivers_macro",
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
