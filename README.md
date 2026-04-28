<!-- Copyright (c) 2026 swseokx. All rights reserved. -->

# ARAON Enterprise v4.0 — 설치 및 사용 가이드

## 📁 파일 구조

```
araon/
├── main.py                  # 개통/AS 워크스테이션 (기존 프로그램)
├── admission.py             # 입학식쌤 전용 프로그램 (신규)
├── timetable_data.json      # 시간표 데이터
├── settings.ini             # 설정 파일 (자동 생성)
├── credentials.json         # 구글 API 서비스 계정 키
├── send_msg_btn.png         # 카카오 매크로용 이미지
├── consult_now_btn.png      # 카카오 매크로용 이미지
├── consult_complete_btn.png # 카카오 매크로용 이미지
├── consult_okay_btn.png     # 카카오 매크로용 이미지
├── araon_core/              # 공통 코어 라이브러리
│   ├── __init__.py
│   ├── config_manager.py    # 설정 + 보안 자격증명
│   ├── log_manager.py       # 통합 로그 관리
│   ├── sheet_manager.py     # 구글 시트 연동
│   └── selenium_manager.py  # 셀레늄 드라이버 관리
├── log/                     # 시스템 로그 (자동 생성)
└── setup_log/               # 개통/AS 실적 로그 (자동 생성)
```

---

## 🛠 설치

```bash
pip install customtkinter gspread oauth2client tkcalendar pyperclip
pip install keyboard pyautogui pygetwindow opencv-python numpy
pip install selenium webdriver-manager
pip install keyring          # ← 보안 저장 (강력 권장)
```

---

## 🔐 계정 보안 (v4.0 신규)

`keyring`이 설치되어 있으면 LMS ID/PW가 OS 자격증명 저장소에 저장됩니다.  
설치 전: ini 파일 평문 저장 (구버전 방식)  
설치 후: Windows 자격증명 관리자 / macOS 키체인에 암호화 저장

---

## 🆕 v4.0 주요 변경사항

### 개통/AS 프로그램 (main.py)
| 항목 | 내용 |
|------|------|
| 🔴 보안 | LMS 계정 keyring 암호화 저장 |
| 🔴 안정성 | `except: pass` 전수 제거, 모든 에러 로그 기록 |
| 🟠 성능 | 구글 시트 배치 API (2번 → 1번 호출) |
| 🟠 성능 | ChromeDriver 경로 캐싱 (매번 재설치 방지) |
| 🟠 오동작 | 자동 새로고침 후 알람 재발송 버그 수정 |
| 🟡 UX | 그리드 Entry 읽기전용 처리 |
| 🟡 기능 | 수당 단가 설정 메뉴 추가 |
| 🟡 기능 | 일괄 등록 / 카카오 매크로 중복 실행 방지 |
| 🟡 구조 | 중복 함수 통합 (load_setup_log 계열) |

### 입학식 프로그램 (admission.py) — 신규
- 학생 목록 구글 시트 연동 및 표시
- 미완료/전체 필터, 이름 검색
- 학년별 시간표 자동 배정 (LMS 매크로)
- 수동 완료 체크 / 되돌리기
- 시간표 클립보드 복사
- 진행률 현황 팝업

---

## 🚀 실행

```bash
# 개통/AS 프로그램
python main.py

# 입학식쌤 전용
python admission.py
```

---

## ⚙ 입학식 시트 설정

1. `admission.py` 실행 후 우측 상단 **⚙ 시트 설정** 클릭
2. 입학식 전용 구글 시트 ID 입력 (비우면 기존 시트와 동일 파일 사용)
3. Sheet 이름 입력 (기본값: `입학식`)
4. 저장 후 자동 로드

### 시트 열 구조 (기본)
| 열 | 내용 |
|----|------|
| A  | 학생 이름 |
| B  | 학년 |
| C  | 연락처 |
| D  | 배정 과목 (콤마 구분) |
| E  | 완료 여부 (✅) |

열 구조가 다르면 `admission.py` 상단 `AdmissionSheetManager` 클래스의 상수를 수정하세요.
