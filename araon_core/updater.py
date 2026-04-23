"""
ARAON 자동 업데이트 모듈 (GitHub Releases 기반)
=================================================
흐름:
  1. 앱 시작 → GitHub API 로 최신 릴리즈 확인
  2. 태그(v4.1.0) > 로컬 버전이면 업데이트 정보 반환
  3. 사용자가 수락하면 릴리즈 asset(zip) 다운로드 + 압축 해제
  4. main.exe 재시작

GitHub 설정:
  settings.ini [UPDATE] 섹션:
    repo  = owner/repo          # 예: swseokx/ARAONManagement
    token =                     # 비공개 저장소만 입력, 공개면 빈칸

새 버전 배포 방법:
  make_release.bat 실행 → 자동으로 빌드 + GitHub Release 생성
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

try:
    import requests as _req
    _USE_REQUESTS = True
except ImportError:
    import urllib.request as _urllib_req   # type: ignore
    _USE_REQUESTS = False

_GITHUB_API = "https://api.github.com/repos/{repo}/releases/latest"


# ── 로컬 버전 ─────────────────────────────────────────────────────────────────

def local_version() -> str:
    """
    로컬 version.json 에서 현재 버전 반환. 없으면 '0.0.0'.

    배포 구조: 루트/launcher.exe  루트/bin/version.json
    launcher.exe 기준으로 실행 시 부모 폴더가 루트이므로 bin/ 하위도 탐색.
    """
    base = Path(sys.executable if getattr(sys, 'frozen', False)
                else sys.argv[0]).parent
    for candidate in (base / 'version.json', base / 'bin' / 'version.json'):
        try:
            return json.loads(candidate.read_text(encoding='utf-8')).get('version', '0.0.0')
        except Exception:
            continue
    return '0.0.0'


def _parse_ver(v: str) -> tuple:
    try:
        return tuple(int(x) for x in v.lstrip('v').strip().split('.'))
    except Exception:
        return (0, 0, 0)


# ── GitHub API ────────────────────────────────────────────────────────────────

def _api_get(url: str, token: str = '', timeout: int = 7) -> dict:
    headers = {'Accept': 'application/vnd.github+json',
               'X-GitHub-Api-Version': '2022-11-28'}
    if token:
        headers['Authorization'] = f'Bearer {token}'

    if _USE_REQUESTS:
        resp = _req.get(url, headers=headers, timeout=timeout)
        resp.raise_for_status()
        return resp.json()
    else:
        req = _urllib_req.Request(url, headers=headers)
        with _urllib_req.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode())


def _download_file(url: str, dest: str, token: str = '', progress_cb=None):
    """URL 에서 dest 로 파일 다운로드. progress_cb(ratio 0~1)."""
    headers = {}
    if token:
        headers['Authorization'] = f'Bearer {token}'
        headers['Accept'] = 'application/octet-stream'

    if _USE_REQUESTS:
        resp = _req.get(url, headers=headers, stream=True, timeout=60,
                        allow_redirects=True)
        resp.raise_for_status()
        total = int(resp.headers.get('content-length', 0))
        downloaded = 0
        with open(dest, 'wb') as f:
            for chunk in resp.iter_content(65536):
                if chunk:
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb and total:
                        progress_cb(downloaded / total)
        # content-length 가 없었거나 남은 오차 보정
        if progress_cb:
            progress_cb(1.0)
    else:
        req = _urllib_req.Request(url, headers=headers)
        with _urllib_req.urlopen(req, timeout=60) as resp:
            total = int(resp.headers.get('content-length', 0))
            downloaded = 0
            with open(dest, 'wb') as f:
                while True:
                    chunk = resp.read(65536)
                    if not chunk:
                        break
                    f.write(chunk)
                    downloaded += len(chunk)
                    if progress_cb and total:
                        progress_cb(downloaded / total)
        if progress_cb:
            progress_cb(1.0)


# ── 공개 API ──────────────────────────────────────────────────────────────────

def check_update(repo: str, token: str = '') -> dict | None:
    """
    GitHub 최신 릴리즈와 로컬 버전 비교.

    업데이트가 있으면 반환:
      {'version': '4.1.0', 'notes': '...', 'download_url': '...zip 직접 URL'}
    없거나 오류면 None.
    """
    if not repo:
        return None
    try:
        data = _api_get(_GITHUB_API.format(repo=repo), token)
        tag  = data.get('tag_name', '0.0.0')           # 예: "v4.1.0"

        if _parse_ver(tag) <= _parse_ver(local_version()):
            return None

        # zip asset URL (릴리즈에 업로드한 .zip 파일만 대상)
        asset_url = next(
            (a['browser_download_url']
             for a in data.get('assets', [])
             if a['name'].endswith('.zip')),
            None,
        )
        if not asset_url:
            return None

        return {
            'version':      tag.lstrip('v'),
            'notes':        (data.get('body') or '').strip(),
            'download_url': asset_url,
        }
    except Exception:
        return None


def apply_update(download_url: str, token: str = '', progress_cb=None):
    """
    download_url 에서 zip 다운로드 → 앱 루트 디렉토리에 압축 해제.
    배포 구조: 루트/launcher.exe, 루트/bin/main.exe, 루트/img/...
    실행 중인 EXE(런처)는 Windows에서 덮어쓸 수 없으므로 PermissionError를 무시.
    """
    exe_dir = Path(
        sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
    ).parent.resolve()
    # bin/ 안에서 실행 중이면 상위 폴더(루트)에 압축 해제
    if exe_dir.name == 'bin':
        app_dir = exe_dir.parent
    else:
        app_dir = exe_dir

    fd, tmp_path = tempfile.mkstemp(suffix='.zip')
    os.close(fd)
    try:
        _download_file(download_url, tmp_path, token, progress_cb)
        with zipfile.ZipFile(tmp_path, 'r') as zf:
            for member in zf.infolist():
                try:
                    zf.extract(member, app_dir)
                except PermissionError:
                    # 실행 중인 런처 EXE는 덮어쓸 수 없음 — 무시하고 계속
                    pass
    finally:
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


def restart_app():
    """런처 재시작. launcher.exe → ARAON실행.exe 순으로 탐색."""
    base = Path(
        sys.executable if getattr(sys, 'frozen', False) else sys.argv[0]
    ).parent

    for name in ('launcher.exe', 'ARAON실행.exe'):
        launcher = base / name
        if launcher.exists():
            subprocess.Popen([str(launcher)])
            sys.exit(0)
    subprocess.Popen([sys.executable] + sys.argv)
    sys.exit(0)
