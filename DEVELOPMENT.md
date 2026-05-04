# ARAONManagement 개발 메모

## 프로젝트 구조

- `launcher.py`: 사용자가 실행하는 런처. `settings.ini [UPDATE]`를 읽고 GitHub Releases 최신 버전을 확인한 뒤 `bin/main.exe`를 실행한다.
- `main.py`: 개통/AS 메인 워크스테이션 GUI와 LMS/카카오/구글시트 업무 자동화.
- `admission.py`: 입학식 전용 GUI와 LMS 자동 배정 기능.
- `araon_core/`: 공통 모듈.
  - `config_manager.py`: 설정 파일 생성/마이그레이션, keyring 기반 LMS 계정 저장.
  - `log_manager.py`: 시스템/업무 로그.
  - `sheet_manager.py`: Google Sheets API 연동.
  - `selenium_manager.py`: ChromeDriver 생성, 로그인, 안전 종료.
  - `updater.py`: GitHub Release 확인, zip 다운로드, 앱 폴더 갱신.

## 로컬 실행

```bat
python launcher.py
python main.py
python admission.py
```

의존성은 `requirements.txt` 기준으로 설치한다.

```bat
python -m pip install -r requirements.txt
```

## 릴리즈 흐름

1. `version.json`의 `version` 값을 올린다.
2. `make_release.bat`를 실행한다.
3. 스크립트가 `main.spec`로 `main.exe`, `launcher.spec`로 `ARAON.exe`를 빌드한다.
4. `release/` 폴더를 `ARAON_v{version}.zip`으로 압축한다.
5. GitHub CLI(`gh`) 인증이 되어 있으면 GitHub Release를 생성하고 zip asset을 업로드한다.

배포 zip의 기본 구조:

```text
ARAON.exe
bin/main.exe
bin/version.json
bin/settings.ini
bin/credentials.json
bin/timetable_data.json
img/*.png
img/favicon.ico
```

## 자동 업데이트

- 설정 위치: `settings.ini [UPDATE]`
- 저장소 형식: `owner/repo`
- 공개 저장소면 `token`은 비워둔다.
- 런처는 최신 릴리즈의 `.zip` asset만 업데이트 대상으로 본다.
- `araon_core/updater.py`는 zip 내부 경로가 앱 폴더 밖으로 나가지 않는지 검사한 뒤 압축을 해제한다.

## 정리 원칙

- `settings.ini`, `credentials.json`, `lms_cache.json`, 로그, 가상환경, PyInstaller 산출물은 커밋하지 않는다.
- `.spec` 파일은 빌드 설정 원본이므로 유지한다.
- 기존 사용자 변경이 있는 파일은 되돌리지 않고, 필요한 변경만 누적한다.
