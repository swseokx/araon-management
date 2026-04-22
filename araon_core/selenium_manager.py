"""
araon_core/selenium_manager.py
셀레늄 드라이버 팩토리 + 풀 관리
ChromeDriverManager 는 한 번만 설치 확인 후 경로를 캐싱.
"""

import os
import time
import threading

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


class SeleniumManager:
    """
    드라이버 생성 설정을 중앙화.
    driver_path 는 최초 1회만 resolve 하여 캐싱.
    """

    _driver_path: str | None = None
    _path_lock = threading.Lock()

    @classmethod
    def _resolve_driver_path(cls) -> str:
        with cls._path_lock:
            if cls._driver_path is None:
                cls._driver_path = ChromeDriverManager().install()
        return cls._driver_path

    @staticmethod
    def create_incognito(page_load_strategy: str = 'eager',
                         background: bool = False) -> webdriver.Chrome:
        """시크릿 모드 드라이버 (개통/AS/일괄등록 등 LMS 작업용).

        background=True: 생성 직후 창을 최소화하여 배치 작업이 화면을 가리지 않게 한다.
        """
        path = SeleniumManager._resolve_driver_path()
        opts = Options()
        opts.add_argument('--incognito')
        opts.add_argument('--disable-popup-blocking')
        if background:
            # 화면 밖으로 배치 + 이후 minimize — 포커스 스틸 최소화
            opts.add_argument('--window-position=-32000,-32000')
        opts.page_load_strategy = page_load_strategy
        prefs = {
            "profile.managed_default_content_settings.images": 2,
            "profile.default_content_setting_values.popups": 1,
        }
        opts.add_experimental_option("prefs", prefs)
        drv = webdriver.Chrome(service=Service(path), options=opts)
        if background:
            try:
                drv.minimize_window()
            except Exception:
                pass
        return drv

    @staticmethod
    def create_with_profile(profile_dir: str,
                            page_load_strategy: str = 'eager') -> webdriver.Chrome:
        """
        사용자 프로파일 유지 드라이버 (카카오 로그인 세션 유지 등)
        profile_dir: 절대 경로
        """
        path = SeleniumManager._resolve_driver_path()
        opts = Options()
        opts.add_argument(f'--user-data-dir={profile_dir}')
        opts.page_load_strategy = page_load_strategy
        return webdriver.Chrome(service=Service(path), options=opts)

    @staticmethod
    def safe_quit(driver):
        """드라이버 안전 종료 (alert 처리 포함)"""
        if driver is None:
            return
        try:
            driver.switch_to.alert.accept()
        except Exception:
            pass
        try:
            driver.quit()
        except Exception:
            pass

    @staticmethod
    def lms_login(driver, lms_id: str, lms_pw: str, wait_timeout: int = 10):
        """LMS 공통 로그인 루틴. 로그인 완료까지 대기."""
        from selenium.webdriver.common.by import By
        wait = WebDriverWait(driver, wait_timeout)
        driver.get('https://www.lmsone.com/wcms/')
        driver.execute_script(
            f"document.getElementById('user_id').value = '{lms_id}';"
            f"document.getElementById('user_pw').value = '{lms_pw}';"
            "document.querySelector('.loginBtn a').click();"
        )
        wait.until(EC.url_contains('wcms'))
        return wait
