# --- main.py  ARAON Enterprise Workstation v4.0 ---
# 전면 리팩토링: 보안·안정성·성능 개선판

import sys
import os
import time
import threading
import re
import traceback
import json

import pyperclip
import keyboard
import pyautogui
import pygetwindow as gw
import customtkinter as ctk
import winsound
import cv2
import numpy as np

from tkinter import messagebox, Listbox
from tkcalendar import Calendar
from datetime import datetime, timedelta
from urllib.parse import urlparse

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# 공통 코어
from araon_core import ConfigManager, LogManager, SheetManager, SeleniumManager


# ─────────────────────────────────────────────
#  유틸
# ─────────────────────────────────────────────
def _get_base_path() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ─────────────────────────────────────────────
#  메인 앱
# ─────────────────────────────────────────────
class AraonWorkstation(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.base_path = _get_base_path()

        # 코어 매니저
        self.cfg = ConfigManager(self.base_path)
        self.log = LogManager(self.base_path)
        self.sheet_mgr = SheetManager(self.cfg)

        # UI 상태
        self.selected_date = f"{datetime.now().month}/{datetime.now().day}"
        self.current_data_cache: list[list] = []
        self.row_map: dict[int, int] = {}
        self.row_widgets: dict[int, list] = {}
        self.admission_row_widgets: dict[int, list] = {}
        self.work_drivers: dict = {}
        self.qc_pop = None
        self._monitor_visible: bool = True
        self._admission_needs_render: bool = False

        # 알람 상태 (render_grid 에서 초기화하지 않음)
        self.alarm_states: dict = {}          # 개통/AS 탭: {row_idx: state}
        self.last_alert_times: dict = {}      # 개통/AS 탭: {row_idx: datetime}
        self.admission_alarm_times: dict = {} # 입학식 탭:  {time_str: state}
        self.flashing_rows: set = set()
        self._flash_state = False

        # 매크로 실행 중 플래그 (중복 실행 방지)
        self._bulk_enroll_running = False
        self._macro_running = False

        # LMS 학생 정보 캐시 (일괄 등록 후 프리패치 → 열기 시 즉시 표시)
        self._cache_file = os.path.join(self.base_path, 'lms_cache.json')
        self.lms_info_cache: dict = {}
        self._load_lms_cache()

        self.title("ARAON Enterprise - Workstation v4.0")
        self.geometry("1550x950")
        ctk.set_appearance_mode(
            self.cfg.get('SETTINGS', 'appearance_mode', 'dark')
        )

        self.setup_main_ui()
        self.update_kakao_ui()
        self.update_time_display()
        self._update_flash()

        self.start_hotkey_listener()
        self.start_auto_refresh()
        self.start_time_monitor()

        self.log.write_system('--- 시스템 가동 ---')
        if not self.cfg.is_keyring_available():
            self.log.write_system(
                '⚠ keyring 미설치: 계정이 평문으로 저장됩니다. '
                '`pip install keyring` 권장'
            )
        self.load_sheet_data_async()
        # 시작 시 퀵카피 창 자동 오픈 (UI 완전 렌더링 후)
        self.after(1200, self.open_quick_copy_window)

    # ──────────────────────────────────────────
    #  로그 / 상태바
    # ──────────────────────────────────────────
    def write_system_log(self, msg: str):
        self.log.write_system(msg)
        def _ui():
            if hasattr(self, 'status_bar') and self.status_bar.winfo_exists():
                self.status_bar.configure(text=f'  ● {msg}')
        self.after(0, _ui)

    # ──────────────────────────────────────────
    #  LMS 캐시 관리
    # ──────────────────────────────────────────
    def write_adm_log(self, msg: str):
        """입학식 모니터에 메시지 추가 (메인 스레드에서 호출)."""
        try:
            if hasattr(self, 'adm_monitor') and self.adm_monitor.winfo_exists():
                self.adm_monitor.configure(state='normal')
                self.adm_monitor.insert('end', msg + '\n')
                self.adm_monitor.configure(state='disabled')
                self.adm_monitor.see('end')
        except Exception:
            pass

    def _clear_adm_monitor(self):
        try:
            self.adm_monitor.configure(state='normal')
            self.adm_monitor.delete('1.0', 'end')
            self.adm_monitor.configure(state='disabled')
        except Exception:
            pass

    def _load_lms_cache(self):
        try:
            if os.path.exists(self._cache_file):
                with open(self._cache_file, 'r', encoding='utf-8') as f:
                    self.lms_info_cache = json.load(f)
        except Exception:
            self.lms_info_cache = {}

    def _save_lms_cache(self):
        try:
            with open(self._cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.lms_info_cache, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.write_system_log(f'LMS 캐시 저장 실패: {e}')

    def open_full_log(self):
        pop = ctk.CTkToplevel(self)
        pop.title(f"오늘의 시스템 로그 - {datetime.now().strftime('%Y-%m-%d')}")
        pop.geometry('700x500')
        pop.transient(self)
        pop.attributes('-topmost', True)
        pop.lift()
        pop.focus_force()

        txt = ctk.CTkTextbox(
            pop, font=('Consolas', 13),
            fg_color='#1e1e1e', text_color='#2ecc71'
        )
        txt.pack(fill='both', expand=True, padx=10, pady=10)
        txt.insert('end', self.log.read_system_today())
        txt.configure(state='disabled')
        txt.see('end')

    # ──────────────────────────────────────────
    #  플래시 효과
    # ──────────────────────────────────────────
    def _update_flash(self):
        self._flash_state = not self._flash_state
        for row_idx in list(self.flashing_rows):
            try:
                cache_row = (
                    self.current_data_cache[row_idx]
                    if len(self.current_data_cache) > row_idx else []
                )
                if len(cache_row) > 9 and str(cache_row[9]).strip() in (
                    'ㅇ', '완료', 'O', 'o'
                ):
                    self.flashing_rows.discard(row_idx)
                    if row_idx in self.row_widgets:
                        for w in self.row_widgets[row_idx]:
                            w.configure(fg_color=['#F9F9FA', '#343638'])
                    continue
                if row_idx in self.row_widgets:
                    color = '#7b241c' if self._flash_state else ['#F9F9FA', '#343638']
                    for w in self.row_widgets[row_idx]:
                        w.configure(fg_color=color)
            except Exception:
                self.flashing_rows.discard(row_idx)
        self.after(700, self._update_flash)

    # ──────────────────────────────────────────
    #  토스트 알림
    # ──────────────────────────────────────────
    def show_toast_notification(self, title: str, msg: str, on_click=None):
        """
        on_click: 토스트 클릭 시 실행할 callable (개통 탭에서 학생 팝업 오픈 등)
        """
        try:
            toast = ctk.CTkToplevel(self)
            toast.title('알림')
            toast.overrideredirect(True)
            toast.attributes('-topmost', True)

            sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
            w, h = 330, 110
            toast.geometry(f'{w}x{h}+{sw - w - 20}+{sh - h - 60}')

            border_color = '#27ae60' if on_click else '#e67e22'
            f = ctk.CTkFrame(
                toast, fg_color='#1f538d',
                border_width=2, border_color=border_color, corner_radius=10
            )
            f.pack(fill='both', expand=True)
            if on_click:
                f.configure(cursor='hand2')

            ctk.CTkLabel(
                f, text=title,
                font=('Pretendard', 14, 'bold'), text_color='#f1c40f'
            ).pack(pady=(12, 5))
            ctk.CTkLabel(f, text=msg, font=('Pretendard', 12)).pack(padx=10, pady=(0, 10))
            if on_click:
                hint = ctk.CTkLabel(
                    f, text='▶ 클릭하면 상세 팝업이 열립니다',
                    font=('Pretendard', 10), text_color='#aaaaaa'
                )
                hint.pack(pady=(0, 6))
                hint.configure(cursor='hand2')

            self.after(8000, lambda: toast.destroy() if toast.winfo_exists() else None)

            if on_click:
                def _click(e):
                    try:
                        toast.destroy()
                        on_click()
                    except Exception:
                        pass
                for widget in [toast, f] + list(f.winfo_children()):
                    try:
                        widget.bind('<Button-1>', _click)
                    except Exception:
                        pass
        except Exception as e:
            self.write_system_log(f'토스트 알림 에러: {e}')

    # ──────────────────────────────────────────
    #  시간 모니터 (알람)
    # ──────────────────────────────────────────
    def _parse_time_str(self, time_str: str):
        """HH:MM 형태 파싱. 1~11시는 오후로 보정. 실패 시 None."""
        match = re.search(
            r'(?:(\d{1,2})/(\d{1,2})\s*)?(\d{1,2}):(\d{2})',
            time_str
        )
        if not match:
            return None
        mo_s, d_s, h_s, m_s = match.groups()
        h, m = int(h_s), int(m_s)
        if 1 <= h < 12:
            h += 12
        now = datetime.now()
        try:
            target = now.replace(
                month=int(mo_s) if mo_s else now.month,
                day=int(d_s) if d_s else now.day,
                hour=h, minute=m, second=0, microsecond=0,
            )
            return target
        except ValueError:
            return None

    def trigger_alarm(self, row_idx, student_name, time_str, alarm_type):
        """개통/AS 탭 알람 — 클릭 시 해당 학생 업무 팝업 오픈."""
        if alarm_type == 'past':
            title = '⚠️ 미처리 개통 경고!'
            body = f'[{time_str}] {student_name}님 개통 시간이 지났습니다!'
        else:
            title = '⏰ 개통 10분 전 알림!'
            body = f'[{time_str}] {student_name}님 개통 10분 전입니다.'

        threading.Thread(
            target=lambda: winsound.PlaySound('SystemAsterisk', winsound.SND_ALIAS),
            daemon=True
        ).start()
        self.write_system_log(f'{title} {student_name} ({time_str})')

        # 클릭 시 해당 학생 업무 팝업 오픈
        on_click = lambda ri=row_idx, nm=student_name: self.open_work_popup(ri, nm)
        self.after(0, lambda: self.show_toast_notification(title, body, on_click=on_click))

    def trigger_admission_alarm(self, time_str: str, count: int):
        """입학식 탭 알람 — 시간 단위로 한 번만, 학생 수 표시."""
        title = '🎓 입학식 OT 시간 알림!'
        body = f'[{time_str}] OT 시간입니다. (대상: {count}명)'
        threading.Thread(
            target=lambda: winsound.PlaySound('SystemAsterisk', winsound.SND_ALIAS),
            daemon=True
        ).start()
        self.write_system_log(f'{title} OT {time_str} ({count}명)')
        self.after(0, lambda: self.show_toast_notification(title, body))

    def _get_current_tab(self) -> str:
        try:
            return self.tab_view.get()
        except Exception:
            return ''

    def start_time_monitor(self):
        def check_time():
            while True:
                try:
                    if not self.current_data_cache:
                        time.sleep(5)
                        continue

                    now = datetime.now()
                    current_tab = self._get_current_tab()

                    # ── 개통/AS 탭: M열(index 9) 기준, 학생별 알람 ──────────────
                    if current_tab == '개통/AS':
                        for row_idx, row in enumerate(self.current_data_cache):
                            if len(row) <= 9:
                                continue
                            time_str = str(row[9]).strip()
                            student_name = row[1] if len(row) > 1 else '학생'

                            if time_str in ('ㅇ', '완료', 'O', 'o'):
                                self.alarm_states.pop(row_idx, None)
                                self.flashing_rows.discard(row_idx)
                                continue

                            target = self._parse_time_str(time_str)
                            if not target:
                                continue

                            t_minus_10 = target - timedelta(minutes=10)
                            state = self.alarm_states.get(row_idx, 'pending')

                            if t_minus_10 <= now <= target:
                                self.after(0, lambda r=row_idx: self.flashing_rows.add(r))
                                if state == 'pending':
                                    self.trigger_alarm(row_idx, student_name, time_str, 'before_10')
                                    self.alarm_states[row_idx] = 'notified_10min'
                                    self.last_alert_times[row_idx] = now
                            elif now > target:
                                self.after(0, lambda r=row_idx: self.flashing_rows.add(r))
                                last = self.last_alert_times.get(row_idx, datetime.min)
                                if now >= last + timedelta(minutes=5):
                                    self.trigger_alarm(row_idx, student_name, time_str, 'past')
                                    self.last_alert_times[row_idx] = now

                    # ── 입학식 탭: O열(index 11) 기준, 시간별 1회 알람 ──────────
                    elif current_tab == '입학식':
                        # 같은 OT 시간끼리 그룹핑
                        time_groups: dict[str, int] = {}
                        for row in self.current_data_cache:
                            if len(row) <= 11:
                                continue
                            ts = str(row[11]).strip()
                            if ts:
                                time_groups[ts] = time_groups.get(ts, 0) + 1

                        for time_str, count in time_groups.items():
                            target = self._parse_time_str(time_str)
                            if not target:
                                continue

                            t_minus_10 = target - timedelta(minutes=10)
                            state = self.admission_alarm_times.get(time_str, 'pending')

                            if t_minus_10 <= now <= target:
                                if state == 'pending':
                                    self.trigger_admission_alarm(time_str, count)
                                    self.admission_alarm_times[time_str] = 'notified'
                            elif now > target:
                                # 지난 경우 한 번만 경고 (5분마다 반복 없음)
                                if state == 'pending':
                                    self.trigger_admission_alarm(time_str, count)
                                    self.admission_alarm_times[time_str] = 'notified'

                    time.sleep(5)
                except Exception as e:
                    self.write_system_log(f'[시간모니터 에러] {e}')
                    time.sleep(5)

        threading.Thread(target=check_time, daemon=True).start()

    # ──────────────────────────────────────────
    #  이미지 매칭 (카카오 매크로)
    # ──────────────────────────────────────────
    def find_img_any_scale(self, target_img_path, confidence=0.7, region=None):
        try:
            if region:
                x, y, w, h = region
                screen = pyautogui.screenshot(region=(int(x), int(y), int(w), int(h)))
            else:
                screen = pyautogui.screenshot()

            screen_bgr = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2BGR)
            template_bgr = cv2.imread(target_img_path, cv2.IMREAD_COLOR)
            if template_bgr is None:
                return None

            for scale in [1.0, 1.1, 1.25, 1.5, 0.8, 0.9]:
                if scale == 1.0:
                    tmpl = template_bgr
                else:
                    tmpl = cv2.resize(template_bgr, (0, 0), fx=scale, fy=scale)

                if (tmpl.shape[0] > screen_bgr.shape[0] or
                        tmpl.shape[1] > screen_bgr.shape[1]):
                    continue

                result = cv2.matchTemplate(screen_bgr, tmpl, cv2.TM_CCOEFF_NORMED)
                _, max_val, _, max_loc = cv2.minMaxLoc(result)

                if max_val >= confidence:
                    h_t, w_t = tmpl.shape[:2]
                    mx = max_loc[0] + w_t // 2
                    my = max_loc[1] + h_t // 2
                    if region:
                        mx += region[0]
                        my += region[1]
                    return (mx, my)
            return None
        except Exception as e:
            self.write_system_log(f'이미지 매칭 오류: {e}')
            return None

    def run_kakao_macro(self):
        if self._macro_running:
            self.write_system_log('카카오 매크로 이미 실행 중 — 중복 실행 방지')
            return
        self._macro_running = True
        try:
            og_x, og_y = pyautogui.position()
            win = gw.getActiveWindow()
            if not win:
                return
            self.write_system_log(f'카카오 매크로 가동 (대상: {win.title})')

            win_region = (win.left, win.top, win.width, win.height)
            if win.width <= 0 or win.height <= 0:
                win_region = None

            img_dir = self.base_path

            send_btn = os.path.join(img_dir, 'send_msg_btn.png')
            loc = self.find_img_any_scale(send_btn, confidence=0.7, region=win_region)
            if loc:
                pyautogui.click(loc)
                time.sleep(0.3)
            else:
                self.write_system_log('안내: send_msg_btn.png 없음')

            pyautogui.moveTo(10, 10)

            now_btn = os.path.join(img_dir, 'consult_now_btn.png')
            self.write_system_log('상담중 버튼 탐색 시작...')
            loc = self.find_img_any_scale(now_btn, confidence=0.75, region=win_region)
            if not loc:
                time.sleep(1.5)
                loc = self.find_img_any_scale(now_btn, confidence=0.55, region=win_region)

            if loc:
                pyautogui.click(loc)
                time.sleep(0.4)
            else:
                self.write_system_log('안내: consult_now_btn.png 없음 — 다음 단계 진행')

            pyautogui.moveTo(10, 10)

            comp_btn = os.path.join(img_dir, 'consult_complete_btn.png')
            start = time.time()
            loc = None
            while time.time() - start < 2.5:
                loc = self.find_img_any_scale(comp_btn, confidence=0.6, region=win_region)
                if loc:
                    break
                time.sleep(0.1)

            if loc:
                pyautogui.click(loc)
                time.sleep(0.5)
            else:
                self.write_system_log('에러: consult_complete_btn.png 찾기 실패')
                return

            pyautogui.moveTo(10, 10)

            ok_btn = os.path.join(img_dir, 'consult_okay_btn.png')
            loc = self.find_img_any_scale(ok_btn, confidence=0.7, region=win_region)
            if loc:
                pyautogui.click(loc)
                time.sleep(0.4)
            else:
                self.write_system_log('에러: consult_okay_btn.png 찾기 실패')
                return

            keyboard.press_and_release('ctrl+w')
            pyautogui.moveTo(og_x, og_y)
            self.write_system_log('카톡 상담 완료')
            self.increment_kakao_count()

        except Exception as e:
            self.write_system_log(f'카카오 매크로 에러: {e}')
        finally:
            self._macro_running = False

    def increment_kakao_count(self):
        today = datetime.now().strftime('%Y-%m-%d')
        current = int(self.cfg.get('KAKAO_STATS', today, '0'))
        self.cfg.set('KAKAO_STATS', today, str(current + 1))
        self.cfg.save()
        self.after(0, self.update_kakao_ui)

    def update_kakao_ui(self):
        today = datetime.now().strftime('%Y-%m-%d')
        count = int(self.cfg.get('KAKAO_STATS', today, '0'))
        self.kakao_lbl.configure(text=f'💬 오늘 카톡 상담: {count}건')

    # ──────────────────────────────────────────
    #  단축키 / 자동새로고침
    # ──────────────────────────────────────────
    def start_hotkey_listener(self):
        def listener():
            hk = self.cfg.get('SETTINGS', 'hotkey', 'F4')
            keyboard.unhook_all()
            keyboard.add_hotkey(hk, self.run_kakao_macro, suppress=False)
            self.write_system_log(f'단축키 [{hk}] 활성화')
            keyboard.wait()
        threading.Thread(target=listener, daemon=True).start()

    def start_auto_refresh(self):
        self.after(300000, self.start_auto_refresh)
        self.load_sheet_data_async()

    # ──────────────────────────────────────────
    #  구글 시트 로드
    # ──────────────────────────────────────────
    def load_sheet_data_async(self):
        threading.Thread(target=self.load_sheet_data, daemon=True).start()

    def load_sheet_data(self):
        try:
            self.write_system_log('시트 데이터 동기화 중...')
            data, row_map = self.sheet_mgr.load_day_data(self.selected_date)
            self.current_data_cache = data
            self.row_map = row_map
            # 알람 상태는 유지 (render_grid 에서 초기화하지 않음)
            # 현재 보이는 탭만 즉시 렌더링, 비활성 탭은 전환 시 lazy 렌더링
            self.after(0, lambda: self.render_grid(data))
            if self._get_current_tab() == '입학식':
                self.after(0, lambda: self.render_admission_grid(data))
                self._admission_needs_render = False
            else:
                self._admission_needs_render = True
            self.after(0, self.load_setup_log)
            self.write_system_log(f'데이터 로드 완료 ({len(data)}건)')
        except Exception as e:
            self.write_system_log(f'시트 로드 실패: {e}\n{traceback.format_exc()}')

    # ──────────────────────────────────────────
    #  탭 전환 콜백 (lazy render)
    # ──────────────────────────────────────────
    def _on_tab_switch(self):
        tab = self._get_current_tab()
        if tab == '입학식':
            self.open_monitor_f.pack_forget()
            self.adm_monitor_f.pack(fill='both', expand=True)
            if self._admission_needs_render:
                self.render_admission_grid(self.current_data_cache)
                self._admission_needs_render = False
        else:
            self.adm_monitor_f.pack_forget()
            self.open_monitor_f.pack(fill='both', expand=True)

    # ──────────────────────────────────────────
    #  Setup Monitor 사이드바 토글
    # ──────────────────────────────────────────
    def toggle_monitor_sidebar(self):
        if self._monitor_visible:
            self.monitor_frame.pack_forget()
            self.monitor_toggle_btn.configure(text='▶ 모니터')
            self._monitor_visible = False
        else:
            self.monitor_frame.pack(side='right', fill='both')
            self.monitor_toggle_btn.configure(text='◀ 모니터')
            self._monitor_visible = True

    # ──────────────────────────────────────────
    #  입학식 등록 (→ 입학식 관리 시트)
    # ──────────────────────────────────────────
    def register_to_admission_sheet(self):
        if not self.current_data_cache:
            messagebox.showwarning('알림', '먼저 데이터를 불러오세요.')
            return

        # D:S 인덱스: E=1(학생명), F=2(학년), O=11(입학식시간)
        rows_to_send = []
        for row in self.current_data_cache:
            f_row = row + [''] * 20
            name = str(f_row[1]).strip()
            if not name:
                continue
            grade = str(f_row[2]).strip()
            time_ = str(f_row[11]).strip()
            rows_to_send.append([name, grade, time_])

        if not rows_to_send:
            messagebox.showinfo('알림', '등록할 학생 데이터가 없습니다.')
            return

        def _do_register():
            try:
                self.write_system_log('입학식 관리 시트에 데이터 등록 중...')
                added, skipped = self.sheet_mgr.write_to_admission_sheet(
                    self.selected_date, rows_to_send
                )
                msg = f'등록 완료!\n추가: {added}건  /  중복 스킵: {skipped}건'
                self.write_system_log(f'입학식 등록 완료 (추가:{added}, 스킵:{skipped})')
                self.after(0, lambda: messagebox.showinfo('입학식 등록', msg))
            except Exception as e:
                err = f'입학식 등록 실패: {e}'
                self.write_system_log(err)
                self.after(0, lambda: messagebox.showerror('오류', err))

        threading.Thread(target=_do_register, daemon=True).start()

    # ──────────────────────────────────────────
    #  UI 구성
    # ──────────────────────────────────────────
    def setup_main_ui(self):
        # ── 상단 내비게이션 바 ──
        nav = ctk.CTkFrame(self, height=70, fg_color='#1a1a1a', corner_radius=0)
        nav.pack(side='top', fill='x')

        self.date_btn = ctk.CTkButton(
            nav, text=f'📅 {self.selected_date}',
            font=('Pretendard', 16, 'bold'), command=self.open_calendar
        )
        self.date_btn.pack(side='left', padx=20)
        ctk.CTkButton(nav, text='⚙ 환경설정', width=100,
                      fg_color='#444444', command=self.open_settings_menu
                      ).pack(side='left', padx=5)
        ctk.CTkButton(nav, text='🔄 새로고침', width=100,
                      fg_color='#2c3e50', command=self.load_sheet_data_async
                      ).pack(side='left', padx=5)
        ctk.CTkButton(nav, text='👨‍🎓 일괄 강의방 등록', width=130,
                      fg_color='#8e44ad', hover_color='#732d91',
                      command=self.start_bulk_enroll
                      ).pack(side='left', padx=15)
        ctk.CTkButton(nav, text='📅 시간표 작성', width=100,
                      fg_color='#27ae60', hover_color='#1e8449',
                      command=self.open_timetable_popup
                      ).pack(side='left', padx=5)
        ctk.CTkButton(nav, text='📋 퀵카피 창 열기', width=120,
                      fg_color='#e67e22', hover_color='#d35400',
                      command=self.open_quick_copy_window
                      ).pack(side='left', padx=15)

        # 모니터 토글 버튼
        self.monitor_toggle_btn = ctk.CTkButton(
            nav, text='◀ 모니터', width=90,
            fg_color='#1f538d', hover_color='#174070',
            font=('Pretendard', 12, 'bold'),
            command=self.toggle_monitor_sidebar
        )
        self.monitor_toggle_btn.pack(side='right', padx=10)

        self.time_lbl = ctk.CTkLabel(
            nav, text='', font=('Consolas', 15), text_color='#888888'
        )
        self.time_lbl.pack(side='right', padx=20)

        # ── 상태바 (하단) ──
        self.status_bar = ctk.CTkLabel(
            self, text='  ● 시스템 대기 중', height=30,
            fg_color='#111111', text_color='#00FF00',
            anchor='w', cursor='hand2'
        )
        self.status_bar.pack(side='bottom', fill='x')
        self.status_bar.bind('<Button-1>', lambda e: self.open_full_log())

        # ── 메인 컨테이너 ──
        container = ctk.CTkFrame(self, fg_color='transparent')
        container.pack(fill='both', expand=True, padx=15, pady=10)

        # ── 탭 영역 (왼쪽) ──
        self.tab_view = ctk.CTkTabview(
            container, fg_color='#2b2b2b',
            segmented_button_fg_color='#1a1a1a',
            segmented_button_selected_color='#1f538d',
            segmented_button_selected_hover_color='#174070',
            segmented_button_unselected_color='#1a1a1a',
            segmented_button_unselected_hover_color='#2c3e50',
        )
        self.tab_view.pack(side='left', fill='both', expand=True, padx=(0, 10))
        self.tab_view.add('개통/AS')
        self.tab_view.add('입학식')
        self.tab_view.configure(command=self._on_tab_switch)

        # ── 개통/AS 탭 내부 ──
        tab_open = self.tab_view.tab('개통/AS')
        self.header_f = ctk.CTkFrame(tab_open, fg_color='#1f538d', height=40)
        self.header_f.pack(fill='x', padx=5, pady=(5, 0))
        self.render_header()
        self.sheet_scroll = ctk.CTkScrollableFrame(tab_open, fg_color='#1e1e1e')
        self.sheet_scroll.pack(fill='both', expand=True, padx=5, pady=5)

        # ── 입학식 탭 내부 ──
        tab_adm = self.tab_view.tab('입학식')
        self.adm_header_f = ctk.CTkFrame(tab_adm, fg_color='#16a085', height=40)
        self.adm_header_f.pack(fill='x', padx=5, pady=(5, 0))
        self.render_admission_header()
        self.admission_scroll = ctk.CTkScrollableFrame(tab_adm, fg_color='#1e1e1e')
        self.admission_scroll.pack(fill='both', expand=True, padx=5, pady=5)

        # ── 사이드바: Setup Monitor (오른쪽, 토글 가능) ──
        self.monitor_frame = ctk.CTkFrame(container, width=390, fg_color='#212121')
        self.monitor_frame.pack(side='right', fill='both')
        self.monitor_frame.pack_propagate(False)

        # 카카오 상담 카운터 (항상 표시)
        self.kakao_lbl = ctk.CTkLabel(
            self.monitor_frame, text='💬 오늘 카톡 상담: 0건',
            font=('Pretendard', 14, 'bold'), text_color='#f1c40f'
        )
        self.kakao_lbl.pack(fill='x', padx=10, pady=(15, 0))

        # ── 개통/AS 모니터 프레임 ──
        self.open_monitor_f = ctk.CTkFrame(self.monitor_frame, fg_color='transparent')
        self.open_monitor_f.pack(fill='both', expand=True)

        self.setup_stat_lbl = ctk.CTkLabel(
            self.open_monitor_f,
            text='개통: 0건  /  AS: 0건',
            font=('Pretendard', 13, 'bold'), text_color='#2ecc71'
        )
        self.setup_stat_lbl.pack(fill='x', padx=10, pady=(6, 0))

        monitor_hdr = ctk.CTkFrame(self.open_monitor_f, fg_color='transparent')
        monitor_hdr.pack(fill='x', padx=10, pady=(4, 4))
        ctk.CTkLabel(
            monitor_hdr, text='[ SETUP MONITOR ]',
            font=('Pretendard', 12, 'bold'), text_color='#3498db'
        ).pack(side='left')

        m_btn_f = ctk.CTkFrame(monitor_hdr, fg_color='transparent')
        m_btn_f.pack(side='right')
        for label, cmd in [
            ('전체', self.load_all_setup_logs),
            ('이번달', self.load_month_setup_logs),
            ('오늘', self.load_today_setup_logs),
        ]:
            ctk.CTkButton(
                m_btn_f, text=label, width=40, height=22,
                font=('Pretendard', 10), command=cmd
            ).pack(side='left', padx=2)
        ctk.CTkButton(
            m_btn_f, text='엑셀', width=35, height=22,
            font=('Pretendard', 10), fg_color='#27ae60', hover_color='#1e8449',
            command=self.export_logs_to_excel
        ).pack(side='left', padx=2)
        ctk.CTkButton(
            m_btn_f, text='📝 편집', width=55, height=22,
            font=('Pretendard', 10, 'bold'), fg_color='#e67e22',
            hover_color='#d35400', command=self.open_log_editor_popup
        ).pack(side='left', padx=2)

        self.setup_monitor = ctk.CTkTextbox(
            self.open_monitor_f, fg_color='#000000',
            text_color='#3498db', font=('Pretendard', 14, 'bold')
        )
        self.setup_monitor.pack(fill='both', expand=True, padx=10, pady=10)
        self.setup_monitor.configure(state='disabled')

        # ── 입학식 모니터 프레임 (탭 전환 시 표시) ──
        self.adm_monitor_f = ctk.CTkFrame(self.monitor_frame, fg_color='transparent')
        # 초기엔 숨김 (개통/AS 탭이 기본)

        adm_mon_hdr = ctk.CTkFrame(self.adm_monitor_f, fg_color='transparent')
        adm_mon_hdr.pack(fill='x', padx=10, pady=(6, 4))
        ctk.CTkLabel(
            adm_mon_hdr, text='[ 입학식 MONITOR ]',
            font=('Pretendard', 12, 'bold'), text_color='#16a085'
        ).pack(side='left')
        ctk.CTkButton(
            adm_mon_hdr, text='지우기', width=50, height=22,
            font=('Pretendard', 10),
            command=self._clear_adm_monitor
        ).pack(side='right', padx=2)

        self.adm_monitor = ctk.CTkTextbox(
            self.adm_monitor_f, fg_color='#000000',
            text_color='#2ecc71', font=('Pretendard', 13)
        )
        self.adm_monitor.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        self.adm_monitor.configure(state='disabled')

    # ─── 개통/AS 탭 헤더 ───────────────────────────────────────────────────────
    # D:S 범위 인덱스: D=0 E=1 F=2 G=3 H=4 I=5 J=6 K=7 L=8 M=9 N=10
    #                 O=11 P=12 Q=13 R=14 S=15
    # 표시 열: E(1) F(2) K(7) M(9) N(10) O(11) P(12) Q(13) R(14) S(15)
    def render_header(self):
        for w in self.header_f.winfo_children():
            w.destroy()
        kw = int(self.cfg.get('SETTINGS', 'k_column_width', '350'))
        cols = [
            ('학생명',  100), ('학년', 50), ('특이사항(K)', kw),
            ('개통시간', 80), ('개통담당', 80),
        ]
        for t, w in cols:
            ctk.CTkLabel(
                self.header_f, text=t, width=w,
                font=('Pretendard', 12, 'bold')
            ).pack(side='left', padx=2)
        ctk.CTkLabel(self.header_f, text='작업', width=180).pack(side='right', padx=5)

    def render_grid(self, data: list[list]):
        for w in self.sheet_scroll.winfo_children():
            w.destroy()
        self.row_widgets = {}
        # ⚠ alarm_states / last_alert_times / flashing_rows 는 초기화하지 않음
        self.flashing_rows.clear()

        kw = int(self.cfg.get('SETTINGS', 'k_column_width', '350'))
        # E(1) F(2) K(7) M(9) N(10) — 개통담당까지만 표시
        COL_SPEC = [
            (1, 100), (2, 50), (7, kw),
            (9, 80), (10, 80),
        ]
        ROW_BG = '#1e1e1e'

        def _make_cell(parent, width, value):
            """CTkLabel: Entry보다 훨씬 가볍고 빠름."""
            return ctk.CTkLabel(
                parent, text=str(value), width=width, height=30,
                anchor='w', fg_color=ROW_BG,
                font=('Pretendard', 12), corner_radius=4
            )

        def render_chunk(start_idx):
            end_idx = min(start_idx + 20, len(data))
            for i in range(start_idx, end_idx):
                row = data[i]
                f_row = row + [''] * 20
                f = ctk.CTkFrame(self.sheet_scroll, fg_color=ROW_BG, corner_radius=4)
                f.pack(fill='x', pady=1)

                row_w_list = []
                for col_idx, w in COL_SPEC:
                    lbl = _make_cell(f, w, f_row[col_idx])
                    lbl.pack(side='left', padx=2)
                    row_w_list.append(lbl)
                self.row_widgets[i] = row_w_list

                ctk.CTkButton(
                    f, text='열기', width=55, height=28, fg_color='#d35400',
                    command=lambda idx=i, n=f_row[1]: self.open_work_popup(idx, n)
                ).pack(side='right', padx=2)
                ctk.CTkButton(
                    f, text='강의배정<', width=65, height=28, fg_color='#2980b9',
                    command=lambda n=f_row[1]: self.start_individual_assign(n)
                ).pack(side='right', padx=2)
                ctk.CTkButton(
                    f, text='💬카톡', width=55, height=28,
                    fg_color='#f1c40f', text_color='black',
                    hover_color='#f39c12',
                    command=lambda n=f_row[1]: self.start_kakao_search(n)
                ).pack(side='right', padx=2)

            if end_idx < len(data):
                self.after(10, lambda: render_chunk(end_idx))
            else:
                self.write_system_log(f'UI 렌더링 완료 ({len(data)}건)')

        if data:
            render_chunk(0)

    # ─── 입학식 탭 헤더 ───────────────────────────────────────────────────────
    # 표시 열: E(1) F(2) K(7) O(11) Q(13) R(14) S(15)
    def render_admission_header(self):
        for w in self.adm_header_f.winfo_children():
            w.destroy()
        kw = int(self.cfg.get('SETTINGS', 'k_column_width', '350'))
        cols = [
            ('학생명',    100), ('학년',   50), ('특이사항(K)', kw),
            ('입학식시간', 100), ('시간표배정', 90), ('첫수업일', 90),
            ('카톡등록',   90),
        ]
        for t, w in cols:
            ctk.CTkLabel(
                self.adm_header_f, text=t, width=w,
                font=('Pretendard', 12, 'bold')
            ).pack(side='left', padx=2)
        ctk.CTkLabel(self.adm_header_f, text='작업', width=135).pack(side='right', padx=5)

    def render_admission_grid(self, data: list[list]):
        for w in self.admission_scroll.winfo_children():
            w.destroy()
        self.admission_row_widgets = {}

        kw = int(self.cfg.get('SETTINGS', 'k_column_width', '350'))
        ADM_COL_SPEC = [
            (1, 100), (2, 50), (7, kw),
            (11, 100), (13, 90), (14, 90), (15, 90),
        ]
        ROW_BG = '#1e1e1e'

        def _make_cell(parent, width, value):
            return ctk.CTkLabel(
                parent, text=str(value), width=width, height=30,
                anchor='w', fg_color=ROW_BG,
                font=('Pretendard', 12), corner_radius=4
            )

        def render_chunk(start_idx):
            end_idx = min(start_idx + 20, len(data))
            for i in range(start_idx, end_idx):
                row = data[i]
                f_row = row + [''] * 20
                f = ctk.CTkFrame(self.admission_scroll, fg_color=ROW_BG, corner_radius=4)
                f.pack(fill='x', pady=1)

                row_w_list = []
                for col_idx, w in ADM_COL_SPEC:
                    lbl = _make_cell(f, w, f_row[col_idx])
                    lbl.pack(side='left', padx=2)
                    row_w_list.append(lbl)
                self.admission_row_widgets[i] = row_w_list

                ot_time = str(f_row[11]).strip()
                stu_name = str(f_row[1]).strip()
                ctk.CTkButton(
                    f, text='OT 팝업', width=65, height=28, fg_color='#8e44ad',
                    command=lambda ts=ot_time: self.open_admission_popup(ts)
                ).pack(side='right', padx=2)
                ctk.CTkButton(
                    f, text='열기', width=55, height=28, fg_color='#d35400',
                    command=lambda idx=i, n=stu_name: self.open_work_popup(idx, n)
                ).pack(side='right', padx=2)

            if end_idx < len(data):
                self.after(10, lambda: render_chunk(end_idx))

        if data:
            render_chunk(0)

    def update_time_display(self):
        self.time_lbl.configure(text=datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
        self.after(1000, self.update_time_display)

    # ──────────────────────────────────────────
    #  퀵카피 창
    # ──────────────────────────────────────────
    def open_quick_copy_window(self):
        if self.qc_pop is not None and self.qc_pop.winfo_exists():
            self.qc_pop.lift()
            self.qc_pop.focus_force()
            return

        self.qc_pop = ctk.CTkToplevel(self)
        self.qc_pop.title('퀵카피 도구')
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = 170, 650
        x = sw - w - 10
        y = (sh - h) // 2 - 100
        self.qc_pop.geometry(f'{w}x{h}+{x}+{y}')

        top_frame = ctk.CTkFrame(self.qc_pop, fg_color='transparent')
        top_frame.pack(fill='x', padx=10, pady=(10, 0))
        topmost_var = ctk.BooleanVar(value=True)
        self.qc_pop.attributes('-topmost', True)

        def toggle_topmost():
            self.qc_pop.attributes('-topmost', topmost_var.get())

        ctk.CTkSwitch(
            top_frame, text='항상 위', font=('Pretendard', 12, 'bold'),
            variable=topmost_var, command=toggle_topmost
        ).pack(side='right')
        ctk.CTkLabel(
            self.qc_pop, text='📋 퀵카피',
            font=('Pretendard', 13, 'bold'), text_color='#28a745'
        ).pack(pady=(0, 10))

        scroll = ctk.CTkScrollableFrame(self.qc_pop, fg_color='transparent')
        scroll.pack(fill='both', expand=True, padx=5, pady=(0, 10))

        for i in range(1, 16):
            t = self.cfg.get('COPY_BUTTONS', f'btn_{i}_title', f'업무 {i}')
            c = self.cfg.get('COPY_BUTTONS', f'btn_{i}_text', '')
            if t or c:
                ctk.CTkButton(
                    scroll, text=t, fg_color='#333333', height=40,
                    font=('Pretendard', 12),
                    command=lambda txt=c: [
                        pyperclip.copy(txt),
                        self.write_system_log('퀵카피 복사 완료')
                    ]
                ).pack(fill='x', pady=3)

    # ──────────────────────────────────────────
    #  셋업 모니터 (개통/AS 로그)
    # ──────────────────────────────────────────
    def _display_setup(self, content: str, open_cnt: int, as_cnt: int, label: str):
        total = open_cnt + as_cnt
        money = total * self.cfg.get_setup_rate()
        summary = (
            f'┌──────────────────────────────────┐\n'
            f'  📊 {label} 정산 현황\n'
            f'  - 개통: {open_cnt}건 / AS: {as_cnt}건\n'
            f'  - 총합: {total}건\n'
            f'  💰 누적 수당: {money:,}원\n'
            f'└──────────────────────────────────┘\n\n'
        )
        # 상단 stat 레이블 갱신
        if hasattr(self, 'setup_stat_lbl'):
            self.setup_stat_lbl.configure(
                text=f'개통: {open_cnt}건  /  AS: {as_cnt}건'
            )
        self.setup_monitor.configure(state='normal')
        self.setup_monitor.delete('1.0', 'end')
        self.setup_monitor.insert(
            '1.0', summary + (content or '해당 기간의 기록이 없습니다.')
        )
        self.setup_monitor.configure(state='disabled')
        self.setup_monitor.see('1.0')

    def _date_str_from_selected(self) -> str:
        try:
            year = datetime.now().year
            obj = datetime.strptime(f'{year}/{self.selected_date}', '%Y/%m/%d')
            return obj.strftime('%Y-%m-%d')
        except Exception:
            return datetime.now().strftime('%Y-%m-%d')

    def load_setup_log(self):
        date_str = self._date_str_from_selected()
        content, oc, ac = self.log.read_setup(date_str)
        self._display_setup(content, oc, ac, f'{self.selected_date} 일일')

    def load_today_setup_logs(self):
        date_str = datetime.now().strftime('%Y-%m-%d')
        content, oc, ac = self.log.read_setup(date_str)
        self._display_setup(content, oc, ac, f'{date_str} (오늘)')

    def load_month_setup_logs(self):
        ym = datetime.now().strftime('%Y-%m')
        content, oc, ac = self.log.read_setup_month(ym)
        self._display_setup(content, oc, ac, f'{ym} 월 누적')

    def load_all_setup_logs(self):
        content, oc, ac = self.log.read_setup_all()
        self._display_setup(content, oc, ac, '전체 누적')

    def export_logs_to_excel(self):
        try:
            ym = datetime.now().strftime('%Y-%m')
            path = self.log.export_month_to_csv(ym, self.cfg.get_setup_rate())
            messagebox.showinfo('완료', f'CSV 저장 완료!\n{path}')
            self.write_system_log(f'월별 CSV 내보내기: {path}')
        except Exception as e:
            messagebox.showerror('오류', f'저장 실패: {e}')

    def open_log_editor_popup(self):
        date_str = self._date_str_from_selected()
        pop = ctk.CTkToplevel(self)
        pop.title('SETUP 로그 편집기')
        pop.geometry('600x500')
        pop.transient(self)
        pop.focus_force()

        ctk.CTkLabel(
            pop, text=f'[{date_str}] 로그 편집',
            font=('Pretendard', 14, 'bold')
        ).pack(pady=10)
        txt = ctk.CTkTextbox(
            pop, font=('Pretendard', 13),
            fg_color='#1e1e1e', text_color='#f1c40f'
        )
        txt.pack(fill='both', expand=True, padx=10, pady=5)
        raw = self.log.read_setup_raw(date_str)
        txt.insert(
            '1.0',
            raw if raw else '해당 날짜의 로그가 없습니다. 작성 후 저장하면 생성됩니다.'
        )

        def save_log():
            try:
                self.log.write_setup_raw(date_str, txt.get('1.0', 'end-1c'))
                self.load_setup_log()
                messagebox.showinfo('저장 완료', '수정되었습니다.', parent=pop)
                self.write_system_log('로그 직접 수정 완료')
                pop.destroy()
            except Exception as e:
                messagebox.showerror('오류', f'저장 실패: {e}', parent=pop)

        ctk.CTkButton(
            pop, text='💾 저장하기', height=40,
            font=('Pretendard', 13, 'bold'),
            fg_color='#27ae60', hover_color='#1e8449',
            command=save_log
        ).pack(pady=10, padx=20, fill='x')

    # ──────────────────────────────────────────
    #  LMS 공통 드라이버 생성
    # ──────────────────────────────────────────
    def _create_lms_driver(self, name: str):
        lms_id, lms_pw = self.cfg.get_credentials()
        try:
            driver = SeleniumManager.create_incognito()
            SeleniumManager.lms_login(driver, lms_id, lms_pw)

            wait = WebDriverWait(driver, 10)
            driver.get('https://www.lmsone.com/wcms/member/memManage/memList.asp')

            search_box = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='keyword'], input[name='keyWord']")
                )
            )
            driver.execute_script(f"arguments[0].value = '{name}';", search_box)
            search_box.send_keys(Keys.ENTER)
            time.sleep(0.5)

            # 양쪽 공백만 제거하고 내부 공백은 그대로 비교 (공백 포함 이름 지원)
            name_norm = ' '.join(name.strip().split()).lower()
            target_link = None
            for link in driver.find_elements(By.CSS_SELECTOR, 'table tbody tr a'):
                link_norm = ' '.join(link.text.strip().split()).lower()
                if link_norm == name_norm or link_norm.startswith(name_norm + '('):
                    target_link = link
                    break

            if not target_link:
                self.write_system_log(f"'{name}' 학생 링크를 찾지 못했습니다.")
                SeleniumManager.safe_quit(driver)
                return None

            driver.execute_script('arguments[0].click();', target_link)
            wait.until(lambda d: len(d.window_handles) > 1)

            all_wins = driver.window_handles
            target_win = all_wins[-1]
            for w in all_wins:
                if w != target_win:
                    driver.switch_to.window(w)
                    driver.close()
            driver.switch_to.window(target_win)
            return driver

        except Exception as e:
            self.write_system_log(f'LMS 드라이버 로딩 실패: {e}')
            SeleniumManager.safe_quit(locals().get('driver'))
            return None

    # ──────────────────────────────────────────
    #  업무 팝업 (열기)
    # ──────────────────────────────────────────
    def open_work_popup(self, ui_row_idx: int, name: str):
        self.write_system_log(f'[{name}] 상세페이지 로딩 중...')
        self.show_toast_notification('LMS 로딩 중', f'[{name}] 정보를 불러옵니다...')
        threading.Thread(
            target=self._fetch_and_show,
            args=(ui_row_idx, name), daemon=True
        ).start()

    def _extract_lms_info(self, driver, name: str) -> dict:
        """driver가 위치한 LMS 학생 상세페이지에서 정보 추출."""
        info = {
            'id': '-', 'nm': '-', 'sch': '-', 'grd': '-',
            'p_nm': '-', 'hp': '-', 'p_hp': '-', 'history': ''
        }
        try:
            self.write_system_log(f'[{name}] 스마트 프레임 탐색 시작')
            found = False

            driver.switch_to.default_content()
            if driver.find_elements(By.ID, 'user_id'):
                found = True
            if not found:
                frames = (
                    driver.find_elements(By.TAG_NAME, 'iframe') +
                    driver.find_elements(By.TAG_NAME, 'frame')
                )
                for idx, frame in enumerate(frames):
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        if driver.find_elements(By.ID, 'user_id'):
                            found = True
                            self.write_system_log(f'[{name}] {idx+1}번 프레임에서 발견')
                            break
                    except Exception:
                        continue

            if not found:
                self.write_system_log(f'[{name}] 데이터 프레임을 찾지 못했습니다.')
            else:
                self.write_system_log(f'[{name}] 위치 확인, 데이터 추출 시작')

            def _get(by, selector, attr='value'):
                try:
                    el = driver.find_element(by, selector)
                    return el.get_attribute(attr) or ''
                except Exception as ex:
                    self.write_system_log(f'{selector} 추출 실패: {str(ex).splitlines()[0]}')
                    return '-'

            info['id'] = _get(By.ID, 'user_id')
            info['nm'] = _get(By.ID, 'user_nm')
            info['sch'] = _get(By.ID, 'school_nm')
            info['p_nm'] = _get(By.NAME, 'parents_nm')

            try:
                grd_el = driver.find_element(By.ID, 'school_year_cd')
                info['grd'] = (
                    Select(grd_el).first_selected_option.text
                    if grd_el.tag_name == 'select'
                    else grd_el.get_attribute('value')
                )
            except Exception as e:
                self.write_system_log(f'grd 추출 실패: {str(e).splitlines()[0]}')

            try:
                hp1_el = driver.find_element(By.ID, 'hp1')
                hp1 = (
                    Select(hp1_el).first_selected_option.text.strip()
                    if hp1_el.tag_name == 'select'
                    else hp1_el.get_attribute('value')
                )
                hp_inputs = driver.find_elements(
                    By.CSS_SELECTOR, "input[name='hp2'], input[id='hp2'], input[id='hp3']"
                )
                hp2 = hp_inputs[0].get_attribute('value') if hp_inputs else ''
                hp3 = hp_inputs[1].get_attribute('value') if len(hp_inputs) > 1 else ''
                info['hp'] = f'{hp1}-{hp2}-{hp3}'
            except Exception as e:
                self.write_system_log(f'hp 추출 실패: {str(e).splitlines()[0]}')

            try:
                p1_el = driver.find_element(By.ID, 'parents_hp1')
                p1 = (
                    Select(p1_el).first_selected_option.text.strip()
                    if p1_el.tag_name == 'select'
                    else p1_el.get_attribute('value')
                )
                p2 = driver.find_element(By.ID, 'parents_hp2').get_attribute('value')
                p3 = driver.find_element(By.ID, 'parents_hp3').get_attribute('value')
                info['p_hp'] = f'{p1}-{p2}-{p3}'
            except Exception as e:
                self.write_system_log(f'p_hp 추출 실패: {str(e).splitlines()[0]}')

            try:
                rows = driver.find_elements(By.CSS_SELECTOR, '#memMagTable tbody tr')
                history_text = ''
                for i in range(0, len(rows), 2):
                    if i + 1 < len(rows):
                        tds = rows[i].find_elements(By.TAG_NAME, 'td')
                        if tds:
                            header = tds[0].text.replace('M', '').replace('X', '').strip()
                            content = rows[i + 1].text.strip()
                            history_text += f'■ {header}\n{content}\n' + '-' * 50 + '\n'
                info['history'] = history_text or '등록된 상담 이력이 없습니다.'
            except Exception as e:
                info['history'] = '이력을 불러오지 못했습니다.'
                self.write_system_log(f'history 추출 실패: {str(e).splitlines()[0]}')

        except Exception:
            self.write_system_log(
                f'[{name}] 크롤링 에러:\n{traceback.format_exc()}'
            )
        return info

    def _fetch_and_show(self, ui_row_idx: int, name: str):
        # ── 캐시 히트: 팝업 즉시 표시 + 백그라운드로 브라우저만 준비 ──
        cached = self.lms_info_cache.get(name)
        if cached:
            self.write_system_log(f'[{name}] 캐시 데이터 사용 (즉시 표시)')
            self.after(0, lambda: self._build_work_popup_ui(ui_row_idx, name, None, cached))

            def _bg_open():
                driver = self._create_lms_driver(name)
                if driver:
                    self.work_drivers[ui_row_idx] = driver
                    self.write_system_log(f'[{name}] 브라우저 백그라운드 준비 완료')
                else:
                    self.write_system_log(f'[{name}] 브라우저 백그라운드 준비 실패 (저장 불가)')
            threading.Thread(target=_bg_open, daemon=True).start()
            return

        # ── 캐시 미스: 기존 방식으로 전체 크롤링 ──
        driver = self._create_lms_driver(name)
        if not driver:
            self.after(
                0, lambda: messagebox.showerror(
                    '오류', f'[{name}] 학생 정보를 불러올 수 없습니다.'
                )
            )
            return

        info = self._extract_lms_info(driver, name)
        # 크롤링 결과를 캐시에 저장
        self.lms_info_cache[name] = info
        self._save_lms_cache()

        self.write_system_log(f'[{name}] 데이터 수집 완료. 팝업 생성...')
        self.after(0, lambda: self._build_work_popup_ui(ui_row_idx, name, driver, info))

    def _build_work_popup_ui(self, ui_row_idx, name, driver, info):
        self.work_drivers[ui_row_idx] = driver

        pop = ctk.CTkToplevel(self)
        pop.title(f'업무 처리 - {name}')
        pop.geometry('1100x850')
        pop.transient(self)
        pop.focus_force()
        pop.protocol('WM_DELETE_WINDOW', lambda: self.close_work_popup(pop, ui_row_idx))

        is_topmost = self.cfg.getboolean('SETTINGS', 'popup_topmost', True)
        pop.attributes('-topmost', is_topmost)

        info_f = ctk.CTkFrame(pop, fg_color='#2b2b2b', corner_radius=10)
        info_f.pack(fill='x', padx=20, pady=(20, 10))

        row1 = ctk.CTkFrame(info_f, fg_color='transparent')
        row1.pack(fill='x', padx=15, pady=(15, 5))

        id_lbl = ctk.CTkLabel(
            row1, text=f"🆔 ID: {info['id']}",
            font=('Pretendard', 14, 'bold'), text_color='#3498db', cursor='hand2'
        )
        id_lbl.pack(side='left')
        id_lbl.bind(
            '<Button-1>',
            lambda e: [pyperclip.copy(info['id']),
                       self.write_system_log(f"ID 복사: {info['id']}")]
        )
        ctk.CTkLabel(
            row1,
            text=(f"   |   👤 {info['nm']}   |   "
                  f"🏫 {info['sch']} {info['grd']}   |   "
                  f"👨‍👩‍👧 {info['p_nm']}"),
            font=('Pretendard', 14, 'bold'), text_color='#3498db'
        ).pack(side='left')

        topmost_var = ctk.BooleanVar(value=is_topmost)
        ctk.CTkSwitch(
            row1, text='항상 위에',
            variable=topmost_var,
            command=lambda: pop.attributes('-topmost', topmost_var.get())
        ).pack(side='right')

        row2 = ctk.CTkFrame(info_f, fg_color='transparent')
        row2.pack(fill='x', padx=15, pady=(5, 15))

        n_col_val = ''
        if (len(self.current_data_cache) > ui_row_idx and
                len(self.current_data_cache[ui_row_idx]) > 10):
            n_col_val = str(self.current_data_cache[ui_row_idx][10]).strip()

        stu_color = '#e74c3c' if '아이' in n_col_val else '#2ecc71'
        par_color = '#e74c3c' if ('모' in n_col_val or '부' in n_col_val) else '#2ecc71'

        ctk.CTkLabel(
            row2, text=f"📱 학생폰: {info['hp']}",
            font=('Pretendard', 14, 'bold'), text_color=stu_color
        ).pack(side='left')
        ctk.CTkLabel(
            row2, text='   |   ',
            font=('Pretendard', 14, 'bold'), text_color='#2ecc71'
        ).pack(side='left')
        ctk.CTkLabel(
            row2, text=f"📞 학부모폰: {info['p_hp']}",
            font=('Pretendard', 14, 'bold'), text_color=par_color
        ).pack(side='left')
        if n_col_val:
            ctk.CTkLabel(
                row2, text=f'   [ 📌 N열: {n_col_val} ]',
                font=('Pretendard', 13, 'bold'), text_color='#f1c40f'
            ).pack(side='left', padx=15)

        main_f = ctk.CTkFrame(pop, fg_color='transparent')
        main_f.pack(fill='both', expand=True, padx=20, pady=10)

        left_f = ctk.CTkFrame(main_f)
        left_f.pack(side='left', fill='both', expand=True, padx=(0, 10))

        ctk.CTkLabel(
            left_f, text='[ 기존 특이사항 / 상담 이력 ]',
            font=('Pretendard', 13, 'bold')
        ).pack(anchor='w', pady=(10, 5), padx=10)
        history_box = ctk.CTkTextbox(
            left_f, height=150, font=('Pretendard', 12),
            fg_color='#1e1e1e', text_color='#aaaaaa'
        )
        history_box.pack(fill='x', padx=10, pady=(0, 10))
        history_box.insert('1.0', info['history'])
        history_box.configure(state='disabled')

        ctk.CTkLabel(
            left_f, text='[ 새 상담 내용 ]',
            font=('Pretendard', 13, 'bold')
        ).pack(anchor='w', pady=(10, 5), padx=10)
        t_var = ctk.StringVar(value='개통')
        rb_f = ctk.CTkFrame(left_f, fg_color='transparent')
        rb_f.pack(anchor='w', padx=10)

        as_f = ctk.CTkFrame(main_f, width=200)

        def on_radio_change():
            if t_var.get() == 'AS':
                as_f.pack(side='left', fill='y')
            else:
                as_f.pack_forget()

        ctk.CTkRadioButton(
            rb_f, text='개통완료', variable=t_var,
            value='개통', command=on_radio_change
        ).pack(side='left', padx=(0, 10))
        ctk.CTkRadioButton(
            rb_f, text='AS완료', variable=t_var,
            value='AS', command=on_radio_change
        ).pack(side='left', padx=10)

        memo = ctk.CTkTextbox(left_f, height=200, font=('Pretendard', 13))
        memo.pack(fill='both', expand=True, padx=10, pady=(10, 0))

        # 퀵카피
        quick_copy_f = ctk.CTkFrame(main_f, width=220)
        quick_copy_f.pack(side='right', fill='y', padx=(10, 0))
        ctk.CTkLabel(
            quick_copy_f, text='📋 퀵카피',
            font=('Pretendard', 12, 'bold')
        ).pack(pady=10)
        qc_scroll = ctk.CTkScrollableFrame(quick_copy_f)
        qc_scroll.pack(fill='both', expand=True, padx=5, pady=(0, 10))
        for i in range(1, 16):
            t = self.cfg.get('COPY_BUTTONS', f'btn_{i}_title', f'업무 {i}')
            c = self.cfg.get('COPY_BUTTONS', f'btn_{i}_text', '')
            if t or c:
                ctk.CTkButton(
                    qc_scroll, text=t, fg_color='#333333', height=32,
                    command=lambda txt=c: [
                        pyperclip.copy(txt),
                        self.write_system_log('퀵카피 복사됨')
                    ]
                ).pack(fill='x', pady=3)

        # AS 상용구
        ctk.CTkLabel(
            as_f, text='AS 상용구\n(더블클릭)',
            font=('Pretendard', 12, 'bold')
        ).pack(pady=10)
        as_list = Listbox(
            as_f, bg='#2b2b2b', fg='white',
            font=('Pretendard', 11), borderwidth=0,
            highlightthickness=0, selectbackground='#1f538d'
        )
        as_list.pack(fill='both', expand=True, padx=10, pady=(0, 10))
        for tmpl in self.cfg.get('AS_TEMPLATES', 'list', '').split(','):
            if tmpl.strip():
                as_list.insert('end', tmpl.strip())

        def add_text(event):
            try:
                sel = as_list.get(as_list.curselection())
                memo.insert('end', f'\n- {sel}')
            except Exception:
                pass

        as_list.bind('<Double-Button-1>', add_text)
        on_radio_change()

        # ── 저장 버튼 ──
        def save_lms():
            save_btn.configure(state='disabled', text='저장 중...')

            def bg_task():
                try:
                    # 캐시 히트로 열린 경우 백그라운드 브라우저 준비 대기 (최대 30초)
                    target_driver = self.work_drivers.get(ui_row_idx)
                    if not target_driver:
                        self.after(
                            0, lambda: save_btn.configure(
                                text='브라우저 준비 중...'
                            )
                        )
                        for _ in range(30):
                            time.sleep(1)
                            target_driver = self.work_drivers.get(ui_row_idx)
                            if target_driver:
                                break
                    if not target_driver:
                        self.after(
                            0, lambda: messagebox.showerror(
                                '오류', '브라우저 세션 준비에 실패했습니다.\n잠시 후 다시 시도해주세요.'
                            )
                        )
                        self.after(
                            0, lambda: save_btn.configure(
                                state='normal', text='LMS 저장 및 완료'
                            )
                        )
                        return

                    memo_raw = memo.get('1.0', 'end').strip()
                    category = t_var.get()
                    final_text = (
                        f'[{category}] {name} '
                        f"{datetime.now().strftime('%m/%d %H:%M')}"
                        f'[{category}]완료\n: {memo_raw}'
                    )

                    wait = WebDriverWait(target_driver, 10)
                    qna = wait.until(
                        EC.presence_of_element_located((By.ID, 'qna_content'))
                    )
                    qna.clear()
                    qna.send_keys(final_text)

                    for eid, val in [
                        ('incall_gb', '417'), ('qna_gb', 'etc'), ('call_gb', '903')
                    ]:
                        Select(target_driver.find_element(By.ID, eid)).select_by_value(val)

                    target_driver.find_element(
                        By.CSS_SELECTOR, "input.button2g[value='저장']"
                    ).click()
                    time.sleep(1.0)

                    try:
                        target_driver.switch_to.alert.accept()
                        time.sleep(0.5)
                    except Exception:
                        pass

                    # 로컬 로그
                    self.log.write_setup(category, name, memo_raw)
                    self.after(0, self.load_setup_log)

                    # 구글 시트
                    self.write_system_log(f'[{name}] 구글 시트 ㅇ 표기 중...')
                    ok = self.sheet_mgr.mark_complete(
                        self.row_map.get(ui_row_idx, -1),
                        self.selected_date, name
                    )
                    if ok:
                        self.write_system_log(f'[{name}] 시트 업데이트 성공')
                    else:
                        self.write_system_log(f'🚨 [{name}] 시트에서 위치를 찾지 못했습니다.')

                    def update_ui_and_close():
                        try:
                            # row_widgets[idx][3] = M열(개통시간) 라벨
                            widgets = self.row_widgets.get(ui_row_idx, [])
                            if len(widgets) > 3:
                                widgets[3].configure(text='ㅇ')
                            while len(self.current_data_cache[ui_row_idx]) <= 9:
                                self.current_data_cache[ui_row_idx].append('')
                            self.current_data_cache[ui_row_idx][9] = 'ㅇ'
                            self.flashing_rows.discard(ui_row_idx)
                            if ui_row_idx in self.row_widgets:
                                for w in self.row_widgets[ui_row_idx]:
                                    w.configure(fg_color=['#F9F9FA', '#343638'])
                        except Exception as e:
                            self.write_system_log(f'UI 업데이트 실패: {e}')

                        try:
                            try:
                                target_driver.switch_to.alert.accept()
                            except Exception:
                                pass
                            target_driver.quit()
                            del self.work_drivers[ui_row_idx]
                        except Exception as e:
                            self.write_system_log(f'브라우저 종료 실패: {e}')

                        pop.destroy()
                        self.write_system_log(f'{name} 업무 저장 완료')

                    self.after(0, update_ui_and_close)

                except Exception as e:
                    self.write_system_log(f'저장 중 오류: {e}')
                    self.after(
                        0, lambda: messagebox.showerror('에러', f'저장 오류:\n{e}')
                    )
                    self.after(
                        0, lambda: save_btn.configure(
                            state='normal', text='LMS 저장 및 완료'
                        )
                    )

            threading.Thread(target=bg_task, daemon=True).start()

        save_btn = ctk.CTkButton(
            left_f, text='LMS 저장 및 완료',
            fg_color='#28a745', height=50,
            font=('Pretendard', 15, 'bold'), command=save_lms
        )
        save_btn.pack(pady=15, padx=10, fill='x')

    def close_work_popup(self, pop, ui_row_idx):
        driver = self.work_drivers.get(ui_row_idx)
        if driver:
            try:
                driver.switch_to.alert.accept()
            except Exception:
                pass
            try:
                driver.quit()
            except Exception as e:
                self.write_system_log(f'브라우저 종료 에러: {e}')
            self.work_drivers.pop(ui_row_idx, None)
        pop.destroy()
        self.write_system_log('업무 팝업 종료')

    # ──────────────────────────────────────────
    #  입학식 OT 팝업
    # ──────────────────────────────────────────
    def open_admission_popup(self, time_str: str):
        """입학식 시간 기준으로 해당 시간 학생 전체를 하나의 팝업으로 표시."""
        if not time_str or not time_str.strip():
            messagebox.showwarning('알림', '입학식 시간 정보가 없습니다.')
            return

        # 같은 OT 시간 학생 수집
        students = []
        for idx, row in enumerate(self.current_data_cache):
            f_row = row + [''] * 20
            if str(f_row[11]).strip() != time_str:
                continue
            students.append({
                'name':    str(f_row[1]).strip(),
                'grade':   str(f_row[2]).strip(),
                'n_col':   str(f_row[10]).strip(),
                'time_str': time_str,
                'row_idx': idx,
            })

        if not students:
            messagebox.showwarning('알림', f'{time_str} 시간 학생이 없습니다.')
            return

        self.write_system_log(
            f'입학식 팝업: {time_str} ({len(students)}명) LMS 정보 로딩 중...'
        )
        self.show_toast_notification(
            '입학식 팝업 로딩',
            f'{time_str} OT — {len(students)}명 정보 수집 중...'
        )
        threading.Thread(
            target=self._fetch_admission_info_bg,
            args=(students,), daemon=True
        ).start()

    def _fetch_admission_info_bg(self, students: list):
        """각 학생 LMS 정보를 캐시 우선으로 수집 후 팝업 생성."""
        for s in students:
            name = s['name']
            cached = self.lms_info_cache.get(name)
            if cached:
                s['lms_info'] = cached
            else:
                driver = self._create_lms_driver(name)
                if driver:
                    info = self._extract_lms_info(driver, name)
                    SeleniumManager.safe_quit(driver)
                    self.lms_info_cache[name] = info
                    self._save_lms_cache()
                    s['lms_info'] = info
                else:
                    s['lms_info'] = {
                        'id': '-', 'nm': '-', 'sch': '-', 'grd': '-',
                        'p_nm': '-', 'hp': '-', 'p_hp': '-', 'history': ''
                    }
        self.after(0, lambda: self._build_admission_popup_ui(students))

    def _lms_write_note(self, driver, text: str) -> bool:
        """드라이버가 위치한 학생 상세페이지에 메모를 저장."""
        try:
            wait = WebDriverWait(driver, 10)
            driver.switch_to.default_content()

            # qna_content 를 현재 컨텍스트나 프레임에서 탐색
            qna = None
            try:
                qna = driver.find_element(By.ID, 'qna_content')
            except Exception:
                frames = (driver.find_elements(By.TAG_NAME, 'iframe') +
                          driver.find_elements(By.TAG_NAME, 'frame'))
                for frame in frames:
                    try:
                        driver.switch_to.default_content()
                        driver.switch_to.frame(frame)
                        qna = driver.find_element(By.ID, 'qna_content')
                        break
                    except Exception:
                        continue

            if not qna:
                return False

            qna.clear()
            qna.send_keys(text)

            for eid, val in [
                ('incall_gb', '417'), ('qna_gb', 'etc'), ('call_gb', '903')
            ]:
                Select(driver.find_element(By.ID, eid)).select_by_value(val)

            driver.find_element(
                By.CSS_SELECTOR, "input.button2g[value='저장']"
            ).click()
            time.sleep(1.0)
            try:
                driver.switch_to.alert.accept()
                time.sleep(0.5)
            except Exception:
                pass
            return True
        except Exception as e:
            self.write_system_log(f'LMS 메모 저장 오류: {e}')
            return False

    def _build_admission_popup_ui(self, students: list):
        time_str = students[0]['time_str'] if students else '-'

        pop = ctk.CTkToplevel(self)
        pop.title(f'입학식 업무 처리 — {time_str} OT ({len(students)}명)')
        pop.geometry('1020x780')
        pop.transient(self)
        pop.focus_force()
        pop.attributes('-topmost', self.cfg.getboolean('SETTINGS', 'popup_topmost', True))

        # ── 헤더 ──
        ctk.CTkLabel(
            pop,
            text=f'🎓  {time_str} OT  —  총 {len(students)}명',
            font=('Pretendard', 17, 'bold'), text_color='#f1c40f'
        ).pack(pady=(15, 5))

        # ── 학생 카드 스크롤 영역 ──
        scroll = ctk.CTkScrollableFrame(pop, fg_color='#1a1a1a')
        scroll.pack(fill='both', expand=True, padx=15, pady=5)

        today_str = f'{datetime.now().month}/{datetime.now().day}'
        student_vars_list = []

        CHECKS = [
            ('kakao',      '카톡등록'),
            ('level',      '레벨테스트'),
            ('note',       '노트'),
            ('form',       '폼등록'),
            ('tt_send',    '전체시간표'),
            ('schedule',   '시간표배정'),
        ]

        for s in students:
            lms = s.get('lms_info', {})

            card = ctk.CTkFrame(scroll, fg_color='#2b2b2b', corner_radius=8)
            card.pack(fill='x', pady=5, padx=2)

            # ── 정보 행 (LMS 저장 토글 포함) ──
            info_row = ctk.CTkFrame(card, fg_color='transparent')
            info_row.pack(fill='x', padx=12, pady=(12, 4))

            save_lms_v = ctk.BooleanVar(value=True)
            ctk.CTkCheckBox(
                info_row, text='', variable=save_lms_v,
                width=24, checkbox_width=20, checkbox_height=20
            ).pack(side='left', padx=(0, 4))
            ctk.CTkLabel(
                info_row,
                text=f'👤 {s["name"]}',
                font=('Pretendard', 14, 'bold'), text_color='#f1c40f'
            ).pack(side='left')
            ctk.CTkLabel(
                info_row, text=f'  {s["grade"]}',
                font=('Pretendard', 13)
            ).pack(side='left')
            ctk.CTkLabel(
                info_row, text=f'  📱 {lms.get("hp", "-")}',
                font=('Pretendard', 13)
            ).pack(side='left')
            ctk.CTkLabel(
                info_row, text=f'  📞 {lms.get("p_hp", "-")}',
                font=('Pretendard', 13)
            ).pack(side='left')
            n_txt = s['n_col']
            n_color = '#e74c3c' if n_txt else '#aaaaaa'
            ctk.CTkLabel(
                info_row, text=f'  📌 {n_txt}',
                font=('Pretendard', 13, 'bold'), text_color=n_color
            ).pack(side='left')
            ctk.CTkLabel(
                info_row, text='  ← LMS저장',
                font=('Pretendard', 10), text_color='#888888'
            ).pack(side='left')

            # ── 체크리스트 행 ──
            cb_row = ctk.CTkFrame(card, fg_color='transparent')
            cb_row.pack(fill='x', padx=12, pady=(0, 12))

            sv = {'student': s, 'save_lms': save_lms_v}
            for key, label in CHECKS:
                var = ctk.BooleanVar()
                sv[key] = var
                ctk.CTkCheckBox(
                    cb_row, text=label, variable=var, width=105
                ).pack(side='left', padx=4)

            # 첫 수업 일자 Entry
            ctk.CTkLabel(
                cb_row, text=' 첫수업:', font=('Pretendard', 12)
            ).pack(side='left', padx=(10, 2))
            first_class_v = ctk.StringVar(value=today_str)
            sv['first_class'] = first_class_v
            ctk.CTkEntry(
                cb_row, textvariable=first_class_v,
                width=65, height=28, font=('Pretendard', 12)
            ).pack(side='left')

            student_vars_list.append(sv)

        # ── 상태 레이블 ──
        status_lbl = ctk.CTkLabel(
            pop, text='', font=('Pretendard', 12), text_color='#2ecc71'
        )
        status_lbl.pack(pady=(4, 0))

        # ── 저장 버튼 ──
        def save_all():
            # UI 값 수집 (메인 스레드에서)
            def _ox(var): return 'O' if var.get() else 'X'
            save_data = []
            for sv in student_vars_list:
                save_data.append({
                    'student':     sv['student'],
                    'save_lms':    sv['save_lms'].get(),
                    'kakao':       _ox(sv['kakao']),
                    'level':       _ox(sv['level']),
                    'note':        _ox(sv['note']),
                    'first_class': sv['first_class'].get().strip() or today_str,
                    'form':        _ox(sv['form']),
                    'tt_send':     _ox(sv['tt_send']),
                    'schedule':    _ox(sv['schedule']),
                })
            save_btn.configure(state='disabled', text='처리 중...')
            threading.Thread(
                target=lambda: _bg_save(save_data), daemon=True
            ).start()

        def _bg_save(save_data: list):
            # 체크된 학생만 처리
            targets = [sd for sd in save_data if sd['save_lms']]
            skipped = len(save_data) - len(targets)
            total   = len(targets)

            self.after(0, lambda: self.write_adm_log(
                f'\n▶ {time_str} OT 처리 시작 '
                f'({total}명 처리 / {skipped}명 스킵)'
            ))

            if total == 0:
                self.after(0, lambda: status_lbl.configure(text='처리할 학생이 없습니다.'))
                self.after(0, lambda: save_btn.configure(state='normal', text='LMS 저장 및 완료'))
                return

            done = 0
            for sd in targets:
                s       = sd['student']
                name    = s['name']
                grade   = s['grade']
                kakao   = sd['kakao']
                level   = sd['level']
                note    = sd['note']
                fc      = sd['first_class']
                form    = sd['form']
                tt_send = sd['tt_send']
                sched   = sd['schedule']

                checklist_str = (
                    f'카톡{kakao}/레벨{level}/노트{note}/'
                    f'첫수업{fc}/폼{form}/시간표{tt_send}/배정{sched}'
                )
                lms_text = (
                    f'입학식 완료 : 카톡등록{kakao}/레벨테스트{level}/노트{note}/'
                    f'첫 수업 일자 {fc}/폼 등록{form}/'
                    f'전체 시간표 발송{tt_send}/시간표 배정{sched}'
                )

                self.after(0, lambda n=name, i=done: status_lbl.configure(
                    text=f'[{i + 1}/{total}] {n} 처리 중...'
                ))

                # ── 1. 입학식 관리 시트 행 추가 (없으면 추가, 있으면 스킵) ──
                try:
                    self.sheet_mgr.write_to_admission_sheet(
                        self.selected_date,
                        [[name, grade, time_str]]
                    )
                    self.write_system_log(f'[{name}] 입학식 시트 행 확인/추가')
                except Exception as e:
                    self.write_system_log(f'[{name}] 시트 행 추가 오류: {e}')

                # ── 2. LMS 저장 ──
                lms_ok = False
                try:
                    drv = self._create_lms_driver(name)
                    if drv:
                        lms_ok = self._lms_write_note(drv, lms_text)
                        SeleniumManager.safe_quit(drv)
                        self.write_system_log(
                            f'[{name}] 입학식 LMS {"저장 완료" if lms_ok else "저장 실패"}'
                        )
                    else:
                        self.write_system_log(f'[{name}] LMS 드라이버 생성 실패')
                except Exception as e:
                    self.write_system_log(f'[{name}] LMS 저장 오류: {e}')

                # ── 3. 입학식 관리 시트 체크리스트 업데이트 ──
                sheet_ok = False
                try:
                    sheet_ok = self.sheet_mgr.update_admission_checklist(
                        self.selected_date, name,
                        kakao, level, note, fc, form, tt_send, sched
                    )
                    self.write_system_log(
                        f'[{name}] 입학식 시트 '
                        f'{"업데이트 완료" if sheet_ok else "행 탐색 실패"}'
                    )
                except Exception as e:
                    self.write_system_log(f'[{name}] 시트 업데이트 오류: {e}')

                # ── 3-1. 신아라오티 시트 O열 'ㅇ' 표시 ──
                try:
                    ot_ok = self.sheet_mgr.mark_ot_complete(self.selected_date, name)
                    self.write_system_log(
                        f'[{name}] 신아라오티 O열 '
                        f'{"ㅇ 표시 완료" if ot_ok else "행 탐색 실패"}'
                    )
                except Exception as e:
                    self.write_system_log(f'[{name}] O열 표시 오류: {e}')

                # ── 4. 입학식 로그 파일 기록 ──
                try:
                    self.log.write_admission(
                        time_str, name, checklist_str, lms_ok, sheet_ok
                    )
                except Exception as e:
                    self.write_system_log(f'[{name}] 입학식 로그 기록 오류: {e}')

                # ── 5. 입학식 모니터 업데이트 ──
                lms_tag   = 'LMS✓' if lms_ok   else 'LMS✗'
                sheet_tag = '시트✓' if sheet_ok else '시트✗'
                log_line  = f'  {name} — {checklist_str} | {lms_tag} {sheet_tag}'
                self.after(0, lambda ll=log_line: self.write_adm_log(ll))
                done += 1

            self.after(0, lambda: self.write_adm_log(f'✅ 완료 ({total}명)\n'))
            self.after(0, lambda: status_lbl.configure(
                text=f'✅ {total}명 처리 완료!'
                     + (f' ({skipped}명 스킵)' if skipped else '')
            ))
            self.after(0, lambda: save_btn.configure(
                state='normal', text='LMS 저장 및 완료'
            ))

        save_btn = ctk.CTkButton(
            pop, text='LMS 저장 및 완료',
            font=('Pretendard', 14, 'bold'), height=46,
            fg_color='#1a7a4a', hover_color='#145c38',
            command=save_all
        )
        save_btn.pack(pady=(5, 15), padx=20, fill='x')

    # ──────────────────────────────────────────
    #  개별 강의 배정
    # ──────────────────────────────────────────
    def start_individual_assign(self, name: str):
        threading.Thread(
            target=self._run_individual_assign, args=(name,), daemon=True
        ).start()

    def _run_individual_assign(self, name: str):
        lms_id, lms_pw = self.cfg.get_credentials()
        driver = None
        try:
            self.write_system_log(f'[{name}] 개별 강의 배정 시작...')
            driver = SeleniumManager.create_incognito()
            wait = SeleniumManager.lms_login(driver, lms_id, lms_pw)
            short_wait = WebDriverWait(driver, 3)

            folder = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[@class='folder' and contains(text(), '방송수업관리')]")
                )
            )
            driver.execute_script('arguments[0].click();', folder)
            time.sleep(0.5)

            file_link = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[@class='file' and contains(text(), '정규방송-코칭')]")
                )
            )
            driver.execute_script('arguments[0].click();', file_link)

            class_link = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(text(), '개통방(서우석 선생님)')]")
                )
            )
            driver.execute_script('arguments[0].click();', class_link)

            search_btn = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@value='검색' and contains(@onclick, 'studentSearch.asp')]")
                )
            )
            main_window = driver.current_window_handle
            driver.execute_script('arguments[0].click();', search_btn)

            wait.until(lambda d: len(d.window_handles) > 1)
            for wh in driver.window_handles:
                if wh != main_window:
                    driver.switch_to.window(wh)
                    break

            kw_input = wait.until(EC.presence_of_element_located((By.NAME, 'keyWord')))
            kw_input.clear()
            kw_input.send_keys(name)

            srch = driver.find_element(
                By.XPATH, "//input[@type='button' and @value='검색' and contains(@class, 'srch')]"
            )
            driver.execute_script('arguments[0].click();', srch)
            time.sleep(1.2)

            checkboxes = driver.find_elements(By.NAME, 'member_seq')
            if not checkboxes:
                self.write_system_log(f'[{name}] 검색 결과 없음')
                return

            name_norm = ' '.join(name.strip().split()).lower()
            exact_cb = paren_cb = None
            for cb in checkboxes:
                try:
                    for td in cb.find_elements(By.XPATH, './ancestor::tr/td'):
                        td_norm = ' '.join(td.text.strip().split()).lower()
                        if td_norm == name_norm:
                            exact_cb = cb
                            break
                        elif td_norm.startswith(name_norm + '(') and not paren_cb:
                            paren_cb = cb
                    if exact_cb:
                        break
                except Exception:
                    pass

            target_cb = exact_cb or paren_cb or checkboxes[0]
            driver.execute_script('arguments[0].click();', target_cb)
            time.sleep(0.3)

            try:
                submit_btn = driver.find_element(
                    By.XPATH, "//input[@type='button' and @name='등록' and @value='등록']"
                )
                driver.execute_script('arguments[0].click();', submit_btn)
            except Exception:
                driver.execute_script('Submit();')

            for _ in range(2):
                try:
                    short_wait.until(EC.alert_is_present()).accept()
                    time.sleep(0.8)
                except Exception:
                    break

            self.write_system_log(f'[{name}] 강의방 배정 성공!')

        except Exception as e:
            self.write_system_log(f'[{name}] 개별 배정 에러: {e}')
        finally:
            SeleniumManager.safe_quit(driver)

    # ──────────────────────────────────────────
    #  카카오 채팅 검색 매크로
    # ──────────────────────────────────────────
    def start_kakao_search(self, name: str):
        threading.Thread(
            target=self._run_kakao_search, args=(name,), daemon=True
        ).start()

    def _run_kakao_search(self, name: str):
        driver = None
        try:
            self.write_system_log(f'[{name}] 카카오 검색 시작...')
            profile_dir = os.path.join(self.base_path, 'chrome_kakao_profile')
            driver = SeleniumManager.create_with_profile(profile_dir)
            wait = WebDriverWait(driver, 10)

            driver.get('https://business.kakao.com/_ADPZb/chats')

            try:
                wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn_opt')))
            except TimeoutException:
                self.write_system_log('카카오 로그인 대기 (최대 60초)')
                WebDriverWait(driver, 60).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, 'button.btn_opt'))
                )

            try:
                filter_btn = wait.until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//button[contains(@class, 'btn_opt')]")
                    )
                )
                if '진행 중 상담' in filter_btn.text:
                    driver.execute_script('arguments[0].click();', filter_btn)
                    time.sleep(0.5)
                    all_chat = wait.until(
                        EC.element_to_be_clickable((By.XPATH, "//*[text()='전체 상담']"))
                    )
                    driver.execute_script('arguments[0].click();', all_chat)
                    time.sleep(0.5)
            except Exception as e:
                self.write_system_log(f'카카오 필터 전환 실패(무시): {e}')

            search_input = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[contains(@class, 'tf_g') and @name='keyword']")
                )
            )
            search_input.clear()
            search_input.send_keys(name)

            search_btn = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//span[contains(@class, 'ico_search')]")
                )
            )
            driver.execute_script('arguments[0].click();', search_btn)
            time.sleep(1.5)

            elements = driver.find_elements(By.XPATH, f"//*[contains(text(), '{name}')]")
            clickable = []
            for el in elements:
                try:
                    if el.is_displayed() and el.tag_name not in ('script', 'style', 'input'):
                        clickable.append(el)
                except Exception:
                    pass

            if not clickable:
                self.after(
                    0, lambda: messagebox.showinfo('알림', f'카카오에 [{name}] 학생이 없습니다.')
                )
                return

            for item in clickable:
                try:
                    driver.execute_script('arguments[0].click();', item)
                    time.sleep(0.5)
                except Exception:
                    pass
            self.write_system_log(f'[{name}] 카카오 채팅방 열기 완료.')

        except Exception as e:
            self.write_system_log(f'카카오 검색 에러: {e}')
        # 참고: 카카오 드라이버는 세션 유지를 위해 quit 하지 않음

    # ──────────────────────────────────────────
    #  시간표 작성 팝업
    # ──────────────────────────────────────────
    def open_timetable_popup(self):
        import json
        tt_path = os.path.join(self.base_path, 'timetable_data.json')
        try:
            with open(tt_path, 'r', encoding='utf-8') as f:
                self.tt_data = json.load(f)
        except Exception as e:
            messagebox.showerror('에러', f'시간표 데이터 로드 실패: {e}')
            return

        pop = ctk.CTkToplevel(self)
        pop.title('수업 시간표 작성')
        pop.geometry('1200x900')
        pop.transient(self)
        pop.focus_force()

        self.selected_subjects: list = []
        self.final_selection: dict = {}
        self.grid_cells: dict = {}
        self.current_matches: dict = {}
        self.grade_var = ctk.StringVar(value='초4')

        left_panel = ctk.CTkFrame(pop, width=250)
        left_panel.pack(side='left', fill='y', padx=10, pady=10)
        right_panel = ctk.CTkFrame(pop)
        right_panel.pack(side='right', fill='both', expand=True, padx=10, pady=10)

        ctk.CTkLabel(left_panel, text='1. 학생 이름',
                     font=('Pretendard', 14, 'bold')).pack(pady=(10, 5))
        self.student_name_entry = ctk.CTkEntry(
            left_panel, placeholder_text='이름 입력', height=30
        )
        self.student_name_entry.pack(fill='x', padx=15, pady=5)

        ctk.CTkLabel(left_panel, text='2. 학년 선택',
                     font=('Pretendard', 14, 'bold')).pack(pady=(15, 5))
        grade_menu = ctk.CTkOptionMenu(
            left_panel,
            values=list(self.tt_data['subjects_by_grade'].keys()),
            variable=self.grade_var,
            command=lambda _: update_subject_list()
        )
        grade_menu.pack(pady=5)

        sub_frame = ctk.CTkScrollableFrame(left_panel, width=200, height=400)
        sub_frame.pack(fill='both', expand=True, padx=5, pady=(10, 0))

        table_f = ctk.CTkFrame(right_panel, fg_color='#1a1a1a')
        table_f.pack(fill='both', expand=True, padx=10, pady=10)

        days = ['월', '화', '수', '목', '금']
        base_times = set(
            t
            for g in self.tt_data['subjects_by_grade'].values()
            for s in g.values()
            for d in s.values()
            for t in d
        )
        all_times = sorted(list(base_times | {'21:40', '22:30', '23:20'}))

        def init_grid():
            for j, day in enumerate(['시간'] + days):
                ctk.CTkLabel(
                    table_f, text=day,
                    font=('Pretendard', 13, 'bold'),
                    width=140, fg_color='#333333'
                ).grid(row=0, column=j, padx=1, pady=1)
            for i, ts in enumerate(all_times):
                ctk.CTkLabel(
                    table_f, text=ts, width=80, fg_color='#2b2b2b'
                ).grid(row=i + 1, column=0, padx=1, pady=1)
                for j, day in enumerate(days):
                    btn = ctk.CTkButton(
                        table_f, text='', width=140, height=45,
                        fg_color='#111111', font=('Pretendard', 11),
                        state='disabled'
                    )
                    btn.grid(row=i + 1, column=j + 1, padx=1, pady=1)
                    btn.configure(command=(lambda d=day, t=ts: lambda: click_cell(d, t))())
                    self.grid_cells[(day, ts)] = btn

        def update_subject_list():
            for w in sub_frame.winfo_children():
                w.destroy()
            self.selected_subjects = []
            self.final_selection = {}
            grade = self.grade_var.get()
            all_subs = (
                list(self.tt_data['subjects_by_grade'][grade].keys()) +
                list(self.tt_data['english_timetable'].keys())
            )
            for sub in all_subs:
                ctk.CTkCheckBox(
                    sub_frame, text=sub,
                    command=lambda s=sub: toggle_subject(s)
                ).pack(anchor='w', pady=2, padx=5)
            update_timetable_view()

        def toggle_subject(sub):
            if sub in self.selected_subjects:
                self.selected_subjects.remove(sub)
            else:
                self.selected_subjects.append(sub)
            for k in [k for k, v in self.final_selection.items() if v == sub]:
                del self.final_selection[k]
            update_timetable_view()

        def update_timetable_view():
            grade = self.grade_var.get()
            subject_counts = {}
            for s in self.final_selection.values():
                subject_counts[s] = subject_counts.get(s, 0) + 1

            sel_sub_data = {}
            for sub in self.selected_subjects:
                sel_sub_data[sub] = (
                    self.tt_data['english_timetable'].get(sub) or
                    self.tt_data['subjects_by_grade'][grade].get(sub, {})
                )

            self.current_matches.clear()
            for day in days:
                for ts in all_times:
                    matches = []
                    key = (day, ts)
                    for sub in self.selected_subjects:
                        if (subject_counts.get(sub, 0) >= 2 and
                                self.final_selection.get(key) != sub):
                            continue
                        for dk, tl in sel_sub_data[sub].items():
                            if day in dk and ts in tl:
                                matches.append(sub)
                    self.current_matches[key] = matches
                    btn = self.grid_cells[key]
                    if matches:
                        is_sel = key in self.final_selection
                        color = '#e67e22' if is_sel else '#1f538d'
                        text = (
                            self.final_selection[key] if is_sel
                            else '\n'.join(matches)
                        )
                        btn.configure(text=text, fg_color=color, state='normal')
                    else:
                        btn.configure(text='', fg_color='#111111', state='disabled')

        def click_cell(day, ts):
            key = (day, ts)
            matches = self.current_matches.get(key, [])
            if not matches:
                return
            if key in self.final_selection:
                del self.final_selection[key]
                update_timetable_view()
                return
            if len(matches) > 1:
                menu = ctk.CTkToplevel(pop)
                menu.title('과목 선택')
                menu.geometry('200x250')
                menu.transient(pop)
                menu.focus_force()
                ctk.CTkLabel(menu, text='등록할 과목 선택',
                             font=('Pretendard', 12, 'bold')).pack(pady=10)
                for m in matches:
                    ctk.CTkButton(
                        menu, text=m,
                        command=lambda val=m: [
                            self.final_selection.update({key: val}),
                            menu.destroy(),
                            update_timetable_view()
                        ]
                    ).pack(pady=2, padx=10, fill='x')
            else:
                self.final_selection[key] = matches[0]
                update_timetable_view()

        def send_to_clipboard():
            if not self.final_selection:
                return
            summary = {}
            for (day, ts), sub in self.final_selection.items():
                summary.setdefault(sub, []).append((day, ts))
            result = f"[과목별 시간표 - {self.grade_var.get()}]\n"
            for sub, slots in summary.items():
                tg = {}
                for d, t in slots:
                    tg.setdefault(t, []).append(d)
                for t, dl in tg.items():
                    result += f'• {sub} {t} {", ".join(dl)}\n'
            pyperclip.copy(result)
            self.write_system_log(f'과목별 시간표 복사 ({len(self.final_selection)}개)')
            self.copy_btn.configure(text='✅ 복사 완료')
            self.after(2000, lambda: self.copy_btn.configure(text='📋 과목별 시간표 복사'))

        def send_full_timetable():
            if not self.selected_subjects:
                messagebox.showwarning('선택 없음', '과목을 체크해주세요.', parent=pop)
                return
            name = self.student_name_entry.get().strip()
            grade = self.grade_var.get()
            days_order = ['월', '화', '수', '목', '금']
            result = f"[전체 시간표 - {name or '학생'}]\n\n"
            for sub in self.selected_subjects:
                result += f'[{sub}]\n'
                sub_data = (
                    self.tt_data['english_timetable'].get(sub) or
                    self.tt_data['subjects_by_grade'].get(grade, {}).get(sub, {})
                )
                slots = []
                for day in days_order:
                    for dk, tl in sub_data.items():
                        if day in dk:
                            for t in tl:
                                slots.append((day, t))
                if not slots:
                    result += ' - 배정 가능한 시간이 없습니다.\n\n'
                    continue
                t2d = {}
                for d, t in slots:
                    if d not in t2d.setdefault(t, []):
                        t2d[t].append(d)
                for t in sorted(t2d):
                    dl = sorted(t2d[t], key=lambda x: days_order.index(x))
                    result += f' - {t} ({", ".join(dl)})\n'
                result += '\n'
            pyperclip.copy(result.strip())
            self.write_system_log(f'전체 시간표 복사 ({len(self.selected_subjects)}과목)')
            self.full_copy_btn.configure(text='✅ 전체 복사 완료')
            self.after(2000, lambda: self.full_copy_btn.configure(text='📝 전체 시간표 복사'))

        def assign_timetable():
            name = self.student_name_entry.get().strip()
            if not name:
                messagebox.showwarning('입력 오류', '학생 이름을 입력해주세요.', parent=pop)
                return
            is_weekly = self.chk_weekly_var.get()
            is_writing = self.chk_writing_var.get()
            self.write_system_log(f'[{name}] 시간표 배정 매크로 실행...')
            threading.Thread(
                target=self._run_timetable_assignment,
                args=(name, is_weekly, is_writing), daemon=True
            ).start()

        def view_timetable():
            name = self.student_name_entry.get().strip()
            if not name:
                messagebox.showwarning('입력 오류', '학생 이름을 입력해주세요.', parent=pop)
                return
            threading.Thread(
                target=self._run_view_timetable, args=(name,), daemon=True
            ).start()

        # 버튼 배치
        self.full_copy_btn = ctk.CTkButton(
            left_panel, text='📝 전체 시간표 복사',
            fg_color='#2980b9', hover_color='#1f618d',
            command=send_full_timetable
        )
        self.full_copy_btn.pack(side='bottom', fill='x', padx=10, pady=(5, 20))

        self.copy_btn = ctk.CTkButton(
            left_panel, text='📋 과목별 시간표 복사',
            fg_color='#27ae60', hover_color='#1e8449',
            command=send_to_clipboard
        )
        self.copy_btn.pack(side='bottom', fill='x', padx=10, pady=(5, 5))

        ctk.CTkButton(
            left_panel, text='👀 시간표 보기',
            fg_color='#f39c12', hover_color='#d68910',
            command=view_timetable
        ).pack(side='bottom', fill='x', padx=10, pady=10)

        ctk.CTkButton(
            left_panel, text='👨‍🏫 시간표 배정',
            fg_color='#8e44ad', command=assign_timetable
        ).pack(side='bottom', fill='x', padx=10, pady=2)

        chk_frame = ctk.CTkFrame(left_panel, fg_color='transparent')
        chk_frame.pack(side='bottom', fill='x', padx=10, pady=(15, 5))
        self.chk_weekly_var = ctk.BooleanVar(value=False)
        self.chk_writing_var = ctk.BooleanVar(value=False)
        ctk.CTkCheckBox(chk_frame, text='주간', variable=self.chk_weekly_var, width=60
                        ).pack(side='left', padx=(10, 5))
        ctk.CTkCheckBox(chk_frame, text='필기', variable=self.chk_writing_var, width=60
                        ).pack(side='left', padx=5)

        init_grid()
        update_subject_list()

    # ──────────────────────────────────────────
    #  시간표 배정 매크로
    # ──────────────────────────────────────────
    def _run_timetable_assignment(self, name: str, assign_weekly=False, assign_writing=False):
        driver = None
        try:
            driver = self._create_lms_driver(name)
            if not driver:
                return
            wait = WebDriverWait(driver, 10)
            self.write_system_log(f'[{name}] 배정 URL 분석 중...')

            current_url = driver.current_url
            mid_match = re.search(r'member_id=([^&]+)', current_url)
            mseq_match = re.search(r'member_seq=([^&]+)', current_url)
            if not (mid_match and mseq_match):
                self.write_system_log('오류: 회원 ID/SEQ 추출 실패')
                return

            parsed = urlparse(current_url)
            base = f'{parsed.scheme}://{parsed.netloc}'
            target_url = (
                f"{base}/wcms/member/memManage/tab/classSearch.asp"
                f"?member_id={mid_match.group(1)}&member_seq={mseq_match.group(1)}"
            )
            driver.get(target_url)
            wait.until(EC.presence_of_element_located((By.ID, 'key')))

            time_map = {
                '14:00': '6611', '15:00': '6610', '16:00': '6490', '16:30': '6614',
                '16:40': '6481', '16:50': '6482', '17:00': '6416', '17:25': '6567',
                '17:30': '6495', '17:50': '1750', '18:00': '6418', '18:20': '6477',
                '18:30': '6496', '19:00': '6420', '19:10': '6421', '19:30': '6497',
                '19:50': '6703', '20:00': '6422', '20:20': '6452', '20:30': '6498',
                '20:40': '6455', '20:50': '6453', '21:00': '6423', '21:40': '6424',
                '22:30': '6494', '23:20': '6701',
            }
            day_map = {
                '월': '6502', '화': '6503', '수': '6504', '목': '6505', '금': '6506'
            }

            for (day, ts), sub in self.final_selection.items():
                try:
                    self.write_system_log(f'[{sub}] 배정 ({day} {ts})')
                    Select(
                        wait.until(EC.presence_of_element_located((By.ID, 'key')))
                    ).select_by_value('tb1.onair_nm')
                    kw = driver.find_element(By.NAME, 'keyWord')
                    kw.clear()
                    kw.send_keys(sub)

                    t_val = time_map.get(ts, '')
                    if t_val:
                        Select(
                            driver.find_element(By.ID, 'sh_school_time')
                        ).select_by_value(t_val)

                    for cb in driver.find_elements(By.NAME, 'sh_week_gb'):
                        if cb.is_selected():
                            driver.execute_script('arguments[0].click();', cb)

                    d_val = day_map.get(day)
                    if d_val:
                        tgt = driver.find_element(
                            By.XPATH,
                            f"//input[@name='sh_week_gb' and @value='{d_val}']"
                        )
                        driver.execute_script('arguments[0].click();', tgt)

                    srch = driver.find_element(
                        By.XPATH,
                        "//input[@type='button' and @value='검색' and contains(@class, 'srch')]"
                    )
                    driver.execute_script('arguments[0].click();', srch)

                    checkbox = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.NAME, 'onair_seqs'))
                    )
                    driver.execute_script('arguments[0].click();', checkbox)

                    assign_btn = driver.find_element(
                        By.XPATH, "//input[@type='button' and @value='방송수업개별배정']"
                    )
                    driver.execute_script('arguments[0].click();', assign_btn)

                    for _ in range(2):
                        try:
                            WebDriverWait(driver, 2).until(EC.alert_is_present()).accept()
                            time.sleep(0.5)
                        except Exception:
                            break
                    self.write_system_log(f'[{sub}] 배정 성공')

                except Exception as e:
                    self.write_system_log(f'[{sub}] 배정 실패 (건너뜀): {e}')

            for extra in (['주간'] if assign_weekly else []) + (['필기'] if assign_writing else []):
                try:
                    self.write_system_log(f'[{extra}] 특수방 배정...')
                    Select(
                        wait.until(EC.presence_of_element_located((By.ID, 'key')))
                    ).select_by_value('tb1.onair_nm')
                    kw = driver.find_element(By.NAME, 'keyWord')
                    kw.clear()
                    kw.send_keys(extra)
                    try:
                        Select(driver.find_element(By.ID, 'sh_school_time')).select_by_index(0)
                    except Exception:
                        pass
                    for cb in driver.find_elements(By.NAME, 'sh_week_gb'):
                        if cb.is_selected():
                            driver.execute_script('arguments[0].click();', cb)

                    srch = driver.find_element(
                        By.XPATH,
                        "//input[@type='button' and @value='검색' and contains(@class, 'srch')]"
                    )
                    driver.execute_script('arguments[0].click();', srch)
                    checkbox = WebDriverWait(driver, 3).until(
                        EC.presence_of_element_located((By.NAME, 'onair_seqs'))
                    )
                    driver.execute_script('arguments[0].click();', checkbox)
                    assign_btn = driver.find_element(
                        By.XPATH, "//input[@type='button' and @value='방송수업개별배정']"
                    )
                    driver.execute_script('arguments[0].click();', assign_btn)
                    for _ in range(2):
                        try:
                            WebDriverWait(driver, 2).until(EC.alert_is_present()).accept()
                            time.sleep(0.5)
                        except Exception:
                            break
                    self.write_system_log(f'[{extra}] 방 배정 성공')
                except Exception as e:
                    self.write_system_log(f'[{extra}] 배정 실패: {e}')

            self.write_system_log(f'[{name}] 모든 배정 완료. 종료...')
            time.sleep(3)
        except Exception as e:
            self.write_system_log(f'시간표 배정 에러: {e}')
        finally:
            SeleniumManager.safe_quit(driver)

    def _run_view_timetable(self, name: str):
        driver = None
        try:
            driver = self._create_lms_driver(name)
            if not driver:
                return
            wait = WebDriverWait(driver, 10)
            tab_btn = wait.until(EC.element_to_be_clickable((By.ID, 'pro_tab4')))
            driver.execute_script('arguments[0].click();', tab_btn)
            time.sleep(1.0)

            all_wins = driver.window_handles
            last = all_wins[-1]
            for w in all_wins:
                if w != last:
                    driver.switch_to.window(w)
                    driver.close()
            driver.switch_to.window(last)
            self.write_system_log(f'[{name}] 시간표 조회 완료')
        except Exception as e:
            self.write_system_log(f'시간표 보기 에러: {e}')
            SeleniumManager.safe_quit(driver)

    # ──────────────────────────────────────────
    #  일괄 강의방 등록
    # ──────────────────────────────────────────
    def start_bulk_enroll(self):
        if self._bulk_enroll_running:
            messagebox.showwarning('경고', '이미 일괄 등록이 진행 중입니다.')
            return
        if not self.current_data_cache:
            messagebox.showwarning('경고', '불러온 학생 데이터가 없습니다.')
            return
        count = len(self.current_data_cache)
        if not messagebox.askyesno(
            '확인',
            f'현재 목록 학생 {count}명을 개통방에 일괄 등록하시겠습니까?\n'
            f'(기존 학생은 초기화됩니다)'
        ):
            return
        self._bulk_enroll_running = True
        threading.Thread(target=self._run_bulk_enroll, daemon=True).start()

    def _run_bulk_enroll(self):
        lms_id, lms_pw = self.cfg.get_credentials()
        driver = None
        try:
            self.write_system_log('일괄 등록 매크로 시작...')
            driver = SeleniumManager.create_incognito()
            wait = SeleniumManager.lms_login(driver, lms_id, lms_pw)
            short_wait = WebDriverWait(driver, 3)

            folder = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[@class='folder' and contains(text(), '방송수업관리')]")
                )
            )
            driver.execute_script('arguments[0].click();', folder)
            time.sleep(0.5)

            file_link = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//span[@class='file' and contains(text(), '정규방송-코칭')]")
                )
            )
            driver.execute_script('arguments[0].click();', file_link)

            class_link = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//a[contains(text(), '개통방(서우석 선생님)')]")
                )
            )
            driver.execute_script('arguments[0].click();', class_link)
            time.sleep(2.0)

            # 기존 학생 삭제
            try:
                self.write_system_log('기존 학생 전체 삭제 중...')
                driver.execute_script(
                    "var chk = document.getElementById('allSelect');"
                    "if(chk && !chk.checked) { chk.click(); }"
                )
                time.sleep(0.5)
                driver.execute_script(
                    "if(typeof delStudent === 'function') { delStudent(); } "
                    "else { var btn = document.querySelector('.button10g[value=\"학생삭제\"]'); "
                    "if(btn) btn.click(); }"
                )
                time.sleep(0.5)
                short_wait.until(EC.alert_is_present()).accept()
                time.sleep(0.5)
                short_wait.until(EC.alert_is_present()).accept()
                time.sleep(1.0)
                self.write_system_log('기존 학생 삭제 완료')
            except Exception as e:
                self.write_system_log(f'삭제 시도 (비어있음 또는 실패): {e}')

            search_btn = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//input[@value='검색' and contains(@onclick, 'studentSearch.asp')]")
                )
            )
            main_window = driver.current_window_handle
            driver.execute_script('arguments[0].click();', search_btn)

            wait.until(lambda d: len(d.window_handles) > 1)
            for wh in driver.window_handles:
                if wh != main_window:
                    driver.switch_to.window(wh)
                    break

            success_count = 0
            for row in self.current_data_cache:
                if len(row) < 2:
                    continue
                student_name = str(row[1]).strip()
                if not student_name or student_name == '-' or '학생명' in student_name:
                    continue

                try:
                    self.write_system_log(f'[{student_name}] 등록 중...')
                    kw = wait.until(EC.presence_of_element_located((By.NAME, 'keyWord')))
                    kw.clear()
                    kw.send_keys(student_name)

                    srch = driver.find_element(
                        By.XPATH,
                        "//input[@type='button' and @value='검색' and contains(@class, 'srch')]"
                    )
                    driver.execute_script('arguments[0].click();', srch)
                    time.sleep(1.2)

                    checkboxes = driver.find_elements(By.NAME, 'member_seq')
                    if not checkboxes:
                        self.write_system_log(f'[{student_name}] 검색 결과 없음')
                        continue

                    name_norm = ' '.join(student_name.strip().split()).lower()
                    exact_cb = paren_cb = None
                    for cb in checkboxes:
                        try:
                            for td in cb.find_elements(By.XPATH, './ancestor::tr/td'):
                                td_norm = ' '.join(td.text.strip().split()).lower()
                                if td_norm == name_norm:
                                    exact_cb = cb
                                    break
                                elif td_norm.startswith(name_norm + '(') and not paren_cb:
                                    paren_cb = cb
                            if exact_cb:
                                break
                        except Exception:
                            pass

                    target_cb = exact_cb or paren_cb or checkboxes[0]
                    driver.execute_script('arguments[0].click();', target_cb)
                    time.sleep(0.3)

                    try:
                        submit_btn = driver.find_element(
                            By.XPATH,
                            "//input[@type='button' and @name='등록' and @value='등록']"
                        )
                        driver.execute_script('arguments[0].click();', submit_btn)
                    except Exception:
                        driver.execute_script('Submit();')

                    for _ in range(2):
                        try:
                            short_wait.until(EC.alert_is_present()).accept()
                            time.sleep(0.8)
                        except Exception:
                            break

                    success_count += 1
                    self.write_system_log(f'[{student_name}] 등록 성공!')

                except Exception as e:
                    self.write_system_log(f'[{student_name}] 에러: {e}')

            driver.close()
            driver.switch_to.window(main_window)
            self.write_system_log(f'일괄 등록 완료! (성공: {success_count}명)')
            self.after(
                0, lambda: messagebox.showinfo(
                    '완료', f'일괄 등록 완료\n(성공: {success_count}건)'
                )
            )

            # 등록 완료 후 학생 정보 프리패치
            names = [
                str(row[1]).strip()
                for row in self.current_data_cache
                if len(row) >= 2 and str(row[1]).strip()
                and str(row[1]).strip() not in ('-', '학생명')
            ]
            if names:
                self.write_system_log(f'학생 정보 프리패치 시작 ({len(names)}명)...')
                threading.Thread(
                    target=self._prefetch_students_info,
                    args=(names,), daemon=True
                ).start()

        except Exception as e:
            self.write_system_log(f'일괄 등록 치명적 에러: {e}')
        finally:
            SeleniumManager.safe_quit(driver)
            self._bulk_enroll_running = False

    def _prefetch_students_info(self, names: list):
        """일괄 등록 후 학생 정보를 미리 크롤링해 캐시에 저장."""
        lms_id, lms_pw = self.cfg.get_credentials()
        driver = None
        success = 0
        skipped = 0
        try:
            driver = SeleniumManager.create_incognito()
            SeleniumManager.lms_login(driver, lms_id, lms_pw)
            wait = WebDriverWait(driver, 10)

            for name in names:
                if name in self.lms_info_cache:
                    skipped += 1
                    continue
                try:
                    self.write_system_log(f'[프리패치] {name} 크롤링 중...')
                    driver.get(
                        'https://www.lmsone.com/wcms/member/memManage/memList.asp'
                    )
                    search_box = wait.until(
                        EC.presence_of_element_located(
                            (By.CSS_SELECTOR,
                             "input[name='keyword'], input[name='keyWord']")
                        )
                    )
                    driver.execute_script(
                        f"arguments[0].value = '{name}';", search_box
                    )
                    search_box.send_keys(Keys.ENTER)
                    time.sleep(0.5)

                    name_norm = ' '.join(name.strip().split()).lower()
                    target_link = None
                    for link in driver.find_elements(
                        By.CSS_SELECTOR, 'table tbody tr a'
                    ):
                        link_norm = ' '.join(link.text.strip().split()).lower()
                        if (link_norm == name_norm
                                or link_norm.startswith(name_norm + '(')):
                            target_link = link
                            break

                    if not target_link:
                        self.write_system_log(f'[프리패치] {name} 검색 결과 없음')
                        continue

                    main_win = driver.current_window_handle
                    driver.execute_script('arguments[0].click();', target_link)
                    wait.until(lambda d: len(d.window_handles) > 1)

                    detail_win = [w for w in driver.window_handles if w != main_win][0]
                    driver.switch_to.window(detail_win)

                    info = self._extract_lms_info(driver, name)
                    self.lms_info_cache[name] = info
                    self._save_lms_cache()
                    success += 1
                    self.write_system_log(
                        f'[프리패치] {name} 완료 ({success}/{len(names) - skipped})'
                    )

                    driver.close()
                    driver.switch_to.window(main_win)

                except Exception as e:
                    self.write_system_log(f'[프리패치] {name} 실패: {e}')
                    try:
                        driver.switch_to.window(driver.window_handles[0])
                    except Exception:
                        pass

        except Exception as e:
            self.write_system_log(f'[프리패치] 치명적 오류: {e}')
        finally:
            SeleniumManager.safe_quit(driver)
            self.write_system_log(
                f'[프리패치] 완료 — 성공: {success}명 / 스킵(기존캐시): {skipped}명'
            )

    # ──────────────────────────────────────────
    #  달력 / 설정 팝업
    # ──────────────────────────────────────────
    def open_calendar(self):
        win = ctk.CTkToplevel(self)
        win.title('날짜 선택')
        win.geometry('300x350')
        win.transient(self)
        win.focus_force()
        cal = Calendar(win, selectmode='day', locale='ko_KR')
        cal.pack(pady=15, padx=15)

        def sel():
            d = cal.selection_get()
            self.selected_date = f'{d.month}/{d.day}'
            self.date_btn.configure(text=f'📅 {self.selected_date}')
            win.destroy()
            self.load_sheet_data_async()

        ctk.CTkButton(win, text='데이터 불러오기', command=sel).pack(pady=10)

    def open_settings_menu(self):
        pop = ctk.CTkToplevel(self)
        pop.title('환경 설정')
        pop.geometry('350x500')
        pop.transient(self)
        pop.focus_force()

        for label, cmd in [
            ('📝 AS 상용구 편집', self.popup_as_templates),
            ('⌨ 단축키 및 테마 설정', self.popup_hotkey_theme),
            ('📋 퀵카피 텍스트 편집', self.popup_copy_edit),
            ('🔑 구글 시트 및 LMS 계정', self.popup_accounts),
            ('💰 수당 단가 설정', self.popup_rates),
        ]:
            ctk.CTkButton(pop, text=label, height=45, command=cmd
                          ).pack(fill='x', padx=30, pady=8)

    def popup_as_templates(self):
        pop = ctk.CTkToplevel(self)
        pop.title('AS 상용구 관리')
        pop.geometry('400x450')
        pop.transient(self)
        pop.focus_force()

        ctk.CTkLabel(
            pop, text='상용구를 줄바꿈으로 구분하여 입력하세요.',
            font=('Pretendard', 12)
        ).pack(pady=10)
        txt = ctk.CTkTextbox(pop, width=350, height=300)
        txt.pack(pady=5)
        current = self.cfg.get('AS_TEMPLATES', 'list', '장비교체 완료,현장점검 완료')
        txt.insert('1.0', current.replace(',', '\n'))

        def save():
            try:
                raw = txt.get('1.0', 'end-1c').strip().split('\n')
                cleaned = ','.join(r.strip() for r in raw if r.strip())
                self.cfg.set('AS_TEMPLATES', 'list', cleaned)
                self.cfg.save()
                pop.destroy()
                self.write_system_log('AS 상용구 업데이트')
            except Exception as e:
                messagebox.showerror('저장 오류', f'{e}')

        ctk.CTkButton(pop, text='저장', command=save).pack(pady=15)

    def popup_hotkey_theme(self):
        pop = ctk.CTkToplevel(self)
        pop.title('단축키 및 테마')
        pop.geometry('380x400')
        pop.transient(self)
        pop.focus_force()

        ctk.CTkLabel(pop, text='카톡 단축키 (ex: F4)').pack(pady=(15, 5))
        e = ctk.CTkEntry(pop)
        e.insert(0, self.cfg.get('SETTINGS', 'hotkey', 'F4'))
        e.pack(pady=5)

        ctk.CTkLabel(pop, text='특이사항(K열) 너비 px').pack(pady=(10, 5))
        kw = int(self.cfg.get('SETTINGS', 'k_column_width', '350'))
        sld = ctk.CTkSlider(pop, from_=200, to=1000)
        sld.set(kw)
        sld.pack(padx=20, fill='x')

        topmost_var = ctk.BooleanVar(
            value=self.cfg.getboolean('SETTINGS', 'popup_topmost', True)
        )
        ctk.CTkSwitch(
            pop, text='업무 팝업 항상 위에',
            font=('Pretendard', 12, 'bold'), variable=topmost_var
        ).pack(pady=(20, 5))

        def toggle_theme():
            new = 'light' if ctk.get_appearance_mode() == 'Dark' else 'dark'
            ctk.set_appearance_mode(new)
            self.cfg.set('SETTINGS', 'appearance_mode', new)

        ctk.CTkButton(
            pop, text='🌓 테마 전환',
            fg_color='transparent', border_width=1, command=toggle_theme
        ).pack(pady=10)

        def save():
            self.cfg.set('SETTINGS', 'hotkey', e.get())
            self.cfg.set('SETTINGS', 'k_column_width', str(int(sld.get())))
            self.cfg.set('SETTINGS', 'popup_topmost', str(topmost_var.get()))
            self.cfg.save()
            keyboard.unhook_all()
            keyboard.add_hotkey(e.get(), self.run_kakao_macro, suppress=False)
            self.render_header()
            self.render_admission_header()
            self.render_grid(self.current_data_cache)
            self.render_admission_grid(self.current_data_cache)
            pop.destroy()
            self.write_system_log('환경설정 저장')

        ctk.CTkButton(pop, text='저장하기', command=save,
                      fg_color='#27ae60').pack(pady=15)

    def popup_copy_edit(self):
        pop = ctk.CTkToplevel(self)
        pop.title('퀵카피 편집')
        pop.geometry('600x700')
        pop.transient(self)
        pop.focus_force()

        s = ctk.CTkScrollableFrame(pop)
        s.pack(fill='both', expand=True, padx=10, pady=10)
        ents = []
        for i in range(1, 16):
            f = ctk.CTkFrame(s)
            f.pack(fill='x', pady=2)
            t_ent = ctk.CTkEntry(f, width=120)
            t_ent.insert(0, self.cfg.get('COPY_BUTTONS', f'btn_{i}_title', ''))
            t_ent.pack(side='left', padx=2)
            x_ent = ctk.CTkEntry(f, width=350)
            x_ent.insert(0, self.cfg.get('COPY_BUTTONS', f'btn_{i}_text', ''))
            x_ent.pack(side='left', padx=2)
            ents.append((i, t_ent, x_ent))

        def save():
            for i, t_e, x_e in ents:
                self.cfg.set('COPY_BUTTONS', f'btn_{i}_title', t_e.get())
                self.cfg.set('COPY_BUTTONS', f'btn_{i}_text', x_e.get())
            self.cfg.save()
            if self.qc_pop is not None and self.qc_pop.winfo_exists():
                self.qc_pop.destroy()
                self.open_quick_copy_window()
            pop.destroy()
            self.write_system_log('퀵카피 저장')

        ctk.CTkButton(pop, text='저장', command=save).pack(pady=10)

    def popup_accounts(self):
        pop = ctk.CTkToplevel(self)
        pop.title('계정 연동 설정')
        pop.geometry('450x480')
        pop.transient(self)
        pop.focus_force()

        if not self.cfg.is_keyring_available():
            ctk.CTkLabel(
                pop,
                text='⚠ keyring 미설치 — 계정이 평문으로 저장됩니다.\n'
                     'pip install keyring 을 실행하면 보안 저장됩니다.',
                text_color='#e67e22', font=('Pretendard', 11)
            ).pack(pady=(10, 0), padx=20)

        ctk.CTkLabel(pop, text='[ LMS 계정 ]').pack(pady=(15, 5))
        lms_id, lms_pw = self.cfg.get_credentials()
        id_e = ctk.CTkEntry(pop, placeholder_text='ID')
        id_e.insert(0, lms_id)
        id_e.pack()
        pw_e = ctk.CTkEntry(pop, placeholder_text='PW', show='*')
        pw_e.insert(0, lms_pw)
        pw_e.pack(pady=5)

        ctk.CTkLabel(pop, text='[ Google Sheet ]').pack(pady=(20, 5))
        sid_e = ctk.CTkEntry(pop, width=350, placeholder_text='Spreadsheet ID')
        sid_e.insert(0, self.cfg.get('MAIN_SHEET', 'SPREADSHEET_ID'))
        sid_e.pack()
        sn_e = ctk.CTkEntry(pop, placeholder_text='Sheet Name')
        sn_e.insert(0, self.cfg.get('MAIN_SHEET', 'SHEET_NAME'))
        sn_e.pack(pady=5)

        def save():
            self.cfg.set_credentials(id_e.get(), pw_e.get())
            self.cfg.set('MAIN_SHEET', 'SPREADSHEET_ID', sid_e.get())
            self.cfg.set('MAIN_SHEET', 'SHEET_NAME', sn_e.get())
            self.cfg.save()
            self.sheet_mgr.invalidate()  # 캐시 초기화
            pop.destroy()
            self.write_system_log('계정 및 시트 정보 갱신')

        ctk.CTkButton(pop, text='저장', command=save).pack(pady=20)

    def popup_rates(self):
        """수당 단가 설정 팝업 (신규)"""
        pop = ctk.CTkToplevel(self)
        pop.title('수당 단가 설정')
        pop.geometry('320x250')
        pop.transient(self)
        pop.focus_force()

        ctk.CTkLabel(pop, text='카카오 상담 1건 단가 (원)',
                     font=('Pretendard', 13)).pack(pady=(20, 5))
        kakao_e = ctk.CTkEntry(pop)
        kakao_e.insert(0, str(self.cfg.get_kakao_rate()))
        kakao_e.pack(pady=5)

        ctk.CTkLabel(pop, text='개통/AS 1건 단가 (원)',
                     font=('Pretendard', 13)).pack(pady=(15, 5))
        setup_e = ctk.CTkEntry(pop)
        setup_e.insert(0, str(self.cfg.get_setup_rate()))
        setup_e.pack(pady=5)

        def save():
            try:
                self.cfg.set('SETTINGS', 'kakao_rate', kakao_e.get())
                self.cfg.set('SETTINGS', 'setup_rate', setup_e.get())
                self.cfg.save()
                self.update_kakao_ui()
                self.load_setup_log()
                pop.destroy()
                self.write_system_log('수당 단가 저장')
            except Exception as e:
                messagebox.showerror('오류', f'저장 실패: {e}')

        ctk.CTkButton(pop, text='저장', fg_color='#27ae60', command=save).pack(pady=20)


# ─────────────────────────────────────────────
if __name__ == '__main__':
    app = AraonWorkstation()
    app.mainloop()
