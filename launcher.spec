# -*- mode: python ; coding: utf-8 -*-
# launcher.spec — ARAON실행.exe 빌드 설정
# 빌드: pyinstaller launcher.spec --clean

a = Analysis(
    ['launcher.py'],
    pathex=[],
    binaries=[],
    datas=[
        # 창 아이콘 번들 (ZIP 압축 해제 없이 실행해도 아이콘 표시)
        ('favicon.ico', '.'),
    ],
    hiddenimports=[
        # requests 가 설치된 경우 자동 포함
        'requests',
        'urllib.request',
        'araon_core.updater',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        # 런처에 불필요한 무거운 패키지 제외 → 빌드 크기 축소
        'customtkinter', 'selenium', 'cv2', 'PIL',
        'gspread', 'oauth2client', 'keyboard', 'pyautogui',
        'numpy', 'pandas',
    ],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='launcher',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,       # 콘솔 창 없음
    icon='favicon.ico',
)
