# Copyright (c) 2026 swseokx. All rights reserved.

# -*- mode: python ; coding: utf-8 -*-
# main.spec — main.exe 빌드 설정
# 빌드: pyinstaller main.spec --clean
#
# ⚠ hiddenimports / datas 편집 원칙:
#   lazy-import / 동적 plugin 로딩을 하는 패키지는 collect_submodules 로
#   전부 번들. 특정 서브모듈만 고정하면 향후 업데이트 시 누락으로 런타임
#   ModuleNotFoundError 재발 가능.

from PyInstaller.utils.hooks import collect_data_files, collect_submodules

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # customtkinter 테마/이미지 파일
        *collect_data_files('customtkinter'),
        # tkcalendar 로케일 데이터
        *collect_data_files('tkcalendar'),
        # webdriver_manager 내부 데이터 (드라이버 캐시 템플릿 등)
        *collect_data_files('webdriver_manager'),
        # 내장 데이터 파일 (시간표 템플릿, 아이콘)
        # ZIP 압축해제 없이 실행해도 최소 동작 보장
        ('timetable_data.json', '.'),
        ('favicon.ico', '.'),
    ],
    hiddenimports=[
        # ── 네트워크/인증 ──────────────────────────────────────────
        # keyring — 플랫폼별 백엔드가 동적 로딩됨
        *collect_submodules('keyring'),
        # oauth2client — 내부 하위 모듈 많음
        *collect_submodules('oauth2client'),
        # google-auth 관련 (oauth2client deprecated 대비)
        *collect_submodules('google'),
        # gspread
        *collect_submodules('gspread'),

        # ── 브라우저 자동화 ────────────────────────────────────────
        # selenium — chrome.webdriver 등 lazy-import 서브모듈
        *collect_submodules('selenium'),
        # webdriver_manager — 드라이버 자동 설치 로직
        *collect_submodules('webdriver_manager'),

        # ── 이미지/GUI ─────────────────────────────────────────────
        # Pillow — 이미지 포맷별 플러그인이 동적 로딩됨
        *collect_submodules('PIL'),
        'PIL._tkinter_finder',
        # tkinter 기본
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        # tkcalendar
        *collect_submodules('tkcalendar'),
        # customtkinter — 테마 모듈 동적 로딩
        *collect_submodules('customtkinter'),

        # ── 입력/화면 ──────────────────────────────────────────────
        # pyautogui — 플랫폼별 백엔드(_pyautogui_win 등) 동적 로딩
        *collect_submodules('pyautogui'),
        # pygetwindow — 마찬가지
        *collect_submodules('pygetwindow'),
        # keyboard — 훅 관련 내부 모듈
        *collect_submodules('keyboard'),

        # ── HTTP ──────────────────────────────────────────────────
        # requests / urllib3 — 압축 디코더 등 lazy
        *collect_submodules('requests'),
        *collect_submodules('urllib3'),

        # ── 데이터 처리 ───────────────────────────────────────────
        # numpy — 내부 서브모듈 많음 (cv2 와 함께 사용)
        *collect_submodules('numpy'),

        # ── 기타 ─────────────────────────────────────────────────
        'pkg_resources.py2_warn',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name='main',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[
        # 안티바이러스 오탐 방지: selenium, cv2 DLL 은 UPX 제외
        'vcruntime*.dll', 'msvcp*.dll', 'cv2*.pyd',
    ],
    runtime_tmpdir=None,
    console=False,       # 콘솔 창 없음
    icon='favicon.ico',
)
