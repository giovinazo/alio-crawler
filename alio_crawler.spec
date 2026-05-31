# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — ALIO Crawler 단일 실행파일(Windows onefile, GUI).

빌드는 GitHub Actions(windows-latest)에서 수행한다(PyInstaller는 크로스컴파일
불가 — Windows exe는 Windows에서만 빌드). 로컬 검증 시:
    pip install -r requirements.txt pyinstaller
    pyinstaller --noconfirm alio_crawler.spec

포함물:
- alio_core.py: 일반 import이므로 Analysis가 자동 수집.
- alio_items.json: 항목 메뉴 캐시. 번들에 동봉해 오프라인/최초실행 시 메뉴 즉시 로드
  (없으면 alio_core가 API에서 재수집하므로 동작엔 지장 없음).
"""

a = Analysis(
    ['alio_crawler_v5.4.py'],
    pathex=[],
    binaries=[],
    datas=[('alio_items.json', '.')],
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='ALIO_Crawler_v5.4',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                 # UPX 미사용 — AV 오탐·런너 미설치 회피
    runtime_tmpdir=None,
    console=False,             # GUI 앱 — 콘솔창 숨김
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
