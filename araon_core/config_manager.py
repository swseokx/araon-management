# Copyright (c) 2026 swseokx. All rights reserved.

"""
araon_core/config_manager.py
설정 관리 + 보안 자격증명 저장 (keyring 사용)
"""

import configparser
import os
import shutil

try:
    import keyring
    KEYRING_AVAILABLE = True
except ImportError:
    KEYRING_AVAILABLE = False

KEYRING_SERVICE = "AraonWorkstation"


class ConfigManager:
    def __init__(self, base_path: str):
        self.base_path = base_path
        self.config_path = os.path.join(base_path, 'settings.ini')
        self.config = configparser.ConfigParser()
        self._init_config()

    def _init_config(self):
        if not os.path.exists(self.config_path):
            # settings.ini.template 이 있으면 자동 복사
            template = os.path.join(self.base_path, 'settings.ini.template')
            if os.path.exists(template):
                shutil.copy2(template, self.config_path)
                self.config.read(self.config_path, encoding='utf-8')
                self._ensure_sections()
                return
            # 템플릿도 없으면 기본값으로 생성
            self.config['DEFAULT'] = {
                'CREDENTIALS_FILE': 'credentials.json',
            }
            self.config['MAIN_SHEET'] = {
                'SPREADSHEET_ID': '',
                'SHEET_NAME': 'Sheet1',
            }
            self.config['SETTINGS'] = {
                'k_column_width': '350',
                'hotkey': 'F4',
                'appearance_mode': 'dark',
                'popup_topmost': 'False',
                'kakao_rate': '500',
                'setup_rate': '3000',
            }
            self.config['AS_TEMPLATES'] = {
                'list': '장비교체 완료,현장점검 완료,배선 재연결,신호 불량 조치'
            }
            self.config['COPY_BUTTONS'] = {}
            for i in range(1, 16):
                self.config['COPY_BUTTONS'][f'btn_{i}_title'] = f'업무 {i}'
                self.config['COPY_BUTTONS'][f'btn_{i}_text'] = ''
            self.config['KAKAO_STATS'] = {}
            if not self.config.has_section('COMPANY_SITE'):
                self.config.add_section('COMPANY_SITE')
            self.save()
        else:
            self.config.read(self.config_path, encoding='utf-8')
            self._ensure_sections()

    def _ensure_sections(self):
        for section in ['MAIN_SHEET', 'KAKAO_STATS', 'SETTINGS', 'AS_TEMPLATES',
                        'COPY_BUTTONS', 'COMPANY_SITE', 'LMS', 'UPDATE', 'ADMISSION',
                        'KAKAO_COORDS']:
            if not self.config.has_section(section):
                self.config.add_section(section)
        # 카톡 매크로 좌표 모드 기본값
        if not self.config.has_option('SETTINGS', 'kakao_macro_mode'):
            self.config.set('SETTINGS', 'kakao_macro_mode', 'coords')
        for k, v in [('send_x', '0'), ('send_y', '0'),
                     ('now_x', '0'), ('now_y', '0'),
                     ('complete_x', '0'), ('complete_y', '0'),
                     ('ok_x', '0'), ('ok_y', '0')]:
            if not self.config.has_option('KAKAO_COORDS', k):
                self.config.set('KAKAO_COORDS', k, v)
        # 단가 키 신규 추가 (기존 파일 호환)
        if not self.config.has_option('SETTINGS', 'kakao_rate'):
            self.config.set('SETTINGS', 'kakao_rate', '500')
        if not self.config.has_option('SETTINGS', 'setup_rate'):
            self.config.set('SETTINGS', 'setup_rate', '3000')
        if not self.config.has_option('SETTINGS', 'kakao_confidence'):
            self.config.set('SETTINGS', 'kakao_confidence', '0.6')
        # 자동업데이트 기본값 (기존 파일 호환)
        _CORRECT_REPO = 'swseokx/araon-management'
        current_repo = self.config.get('UPDATE', 'repo', fallback='').strip()
        # 비어있거나 구버전 잘못된 값(ARAONManagement)이면 올바른 값으로 교정
        if not current_repo or current_repo.lower().replace('-', '') == 'swseokx/araonmanagement':
            self.config.set('UPDATE', 'repo', _CORRECT_REPO)
        if not self.config.has_option('UPDATE', 'token'):
            self.config.set('UPDATE', 'token', '')
        # v4.8.0: popup_topmost 기본값 True → False 마이그레이션
        if self.config.get('SETTINGS', 'popup_topmost', fallback='False').strip() == 'True':
            self.config.set('SETTINGS', 'popup_topmost', 'False')
        self.save()

    def save(self):
        with open(self.config_path, 'w', encoding='utf-8') as f:
            self.config.write(f)

    # ------------------------------------------------------------------
    # 일반 설정 get/set
    # ------------------------------------------------------------------
    def get(self, section, key, fallback=''):
        return self.config.get(section, key, fallback=fallback)

    def getboolean(self, section, key, fallback=True):
        return self.config.getboolean(section, key, fallback=fallback)

    def set(self, section, key, value: str):
        if not self.config.has_section(section):
            self.config.add_section(section)
        self.config.set(section, key, value)

    def get_kakao_rate(self) -> int:
        return int(self.get('SETTINGS', 'kakao_rate', '500'))

    def get_setup_rate(self) -> int:
        return int(self.get('SETTINGS', 'setup_rate', '3000'))

    def get_kakao_confidence(self) -> float:
        try:
            return float(self.get('SETTINGS', 'kakao_confidence', '0.6'))
        except ValueError:
            return 0.6

    # ------------------------------------------------------------------
    # 보안 자격증명 (keyring 우선, fallback: ini 평문 — 마이그레이션용)
    # ------------------------------------------------------------------
    def set_credentials(self, lms_id: str, lms_pw: str):
        """LMS 계정을 OS 키링에 안전하게 저장."""
        if KEYRING_AVAILABLE:
            keyring.set_password(KEYRING_SERVICE, "lms_id", lms_id)
            keyring.set_password(KEYRING_SERVICE, "lms_pw", lms_pw)
            # 기존 ini에 평문이 남아있다면 제거
            if self.config.has_option('COMPANY_SITE', 'id'):
                self.config.remove_option('COMPANY_SITE', 'id')
            if self.config.has_option('COMPANY_SITE', 'pw'):
                self.config.remove_option('COMPANY_SITE', 'pw')
            self.save()
        else:
            # keyring 없으면 기존 방식 (사용자에게 경고 표시 권장)
            self.set('COMPANY_SITE', 'id', lms_id)
            self.set('COMPANY_SITE', 'pw', lms_pw)
            self.save()

    def get_credentials(self) -> tuple[str, str]:
        """(lms_id, lms_pw) 반환. keyring 실패 시 ini fallback."""
        if KEYRING_AVAILABLE:
            lms_id = keyring.get_password(KEYRING_SERVICE, "lms_id") or ''
            lms_pw = keyring.get_password(KEYRING_SERVICE, "lms_pw") or ''
            if lms_id:
                return lms_id, lms_pw
        # Fallback: ini 평문 (구버전 마이그레이션)
        lms_id = self.get('COMPANY_SITE', 'id', '')
        lms_pw = self.get('COMPANY_SITE', 'pw', '')
        return lms_id, lms_pw

    def is_keyring_available(self) -> bool:
        return KEYRING_AVAILABLE
