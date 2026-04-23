# --- admission.py  ARAON 입학식쌤 전용 v1.0 ---
# 시트 열람 → 시간표 배정 → 체크 자동화

import sys
import os
import time
import threading
import json
import re
import traceback

import customtkinter as ctk
from tkinter import messagebox
from datetime import datetime

from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC

from araon_core import ConfigManager, LogManager, SheetManager, SeleniumManager


def _get_base_path() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ──────────────────────────────────────────────────
#  입학식 전용 시트 관리자
#  (기존 개통/AS 시트와 별도 시트를 사용하는 경우를 지원)
# ──────────────────────────────────────────────────
class AdmissionSheetManager:
    """
    입학식 전용 시트:
      - A열: 이름
      - B열: 학년
      - C열: 연락처
      - D열: 배정 과목들 (콤마 구분)
      - E열: 완료 여부 (✅ 또는 공백)
      나머지 열은 운영에 따라 유동적
    시트 구조가 다를 경우 아래 상수만 조정.
    """

    NAME_COL = 1       # A열 (1-indexed)
    GRADE_COL = 2      # B열
    CONTACT_COL = 3    # C열
    SUBJECT_COL = 4    # D열
    DONE_COL = 5       # E열 — 완료 체크

    DONE_MARK = '✅'

    def __init__(self, config_manager: ConfigManager):
        self.cfg = config_manager
        self._sheet = None

    def invalidate(self):
        self._sheet = None

    def _get_sheet(self):
        if self._sheet:
            return self._sheet
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        creds_path = os.path.join(
            self.cfg.base_path,
            self.cfg.get('DEFAULT', 'CREDENTIALS_FILE', 'credentials.json')
        )
        scopes = [
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive",
        ]
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, scopes)
        client = gspread.authorize(creds)
        # 입학식 전용 시트 ID / 이름
        sheet_id = self.cfg.get('ADMISSION', 'spreadsheet_id',
                                self.cfg.get('MAIN_SHEET', 'SPREADSHEET_ID'))
        sheet_name = self.cfg.get('ADMISSION', 'sheet_name', '입학식')
        self._sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        return self._sheet

    def load_students(self) -> list[dict]:
        """전체 학생 목록 로드. 완료 여부 포함."""
        sheet = self._get_sheet()
        all_rows = sheet.get_all_values()
        students = []
        for i, row in enumerate(all_rows[1:], start=2):  # 1행은 헤더
            row += [''] * 10
            name = row[self.NAME_COL - 1].strip()
            if not name:
                continue
            students.append({
                'row': i,
                'name': name,
                'grade': row[self.GRADE_COL - 1].strip(),
                'contact': row[self.CONTACT_COL - 1].strip(),
                'subjects': [
                    s.strip()
                    for s in row[self.SUBJECT_COL - 1].split(',')
                    if s.strip()
                ],
                'done': row[self.DONE_COL - 1].strip() == self.DONE_MARK,
            })
        return students

    def mark_done(self, sheet_row: int):
        """해당 행 완료 체크."""
        sheet = self._get_sheet()
        sheet.update_cell(sheet_row, self.DONE_COL, self.DONE_MARK)

    def unmark_done(self, sheet_row: int):
        """완료 체크 해제."""
        sheet = self._get_sheet()
        sheet.update_cell(sheet_row, self.DONE_COL, '')

    def update_subjects(self, sheet_row: int, subjects: list[str]):
        """과목 목록 업데이트."""
        sheet = self._get_sheet()
        sheet.update_cell(sheet_row, self.SUBJECT_COL, ', '.join(subjects))


# ──────────────────────────────────────────────────
#  메인 앱
# ──────────────────────────────────────────────────
class AdmissionApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.base_path = _get_base_path()
        self.cfg = ConfigManager(self.base_path)
        self.log = LogManager(self.base_path)
        self.adm_sheet = AdmissionSheetManager(self.cfg)

        # 시간표 데이터
        self.tt_data: dict = {}
        self._load_timetable_data()

        # 상태
        self.students: list[dict] = []
        self.student_widgets: dict[int, dict] = {}   # row → widget dict
        self._macro_running = False

        self.title('ARAON 입학식쌤 전용 v1.0')
        self.geometry('1400x900')
        ctk.set_appearance_mode(
            self.cfg.get('SETTINGS', 'appearance_mode', 'dark')
        )

        # 아이콘 적용 (메인창 + 이후 생성되는 모든 CTkToplevel 팝업)
        self._icon_path = self._resolve_icon_path()
        if self._icon_path:
            try:
                self.iconbitmap(self._icon_path)
            except Exception:
                pass
        self._patch_toplevel_icon()

        self._ensure_admission_config()
        self._build_ui()
        self.log.write_system('--- 입학식 프로그램 가동 ---')
        self.load_students_async()

    def _resolve_icon_path(self) -> str:
        """favicon.ico 경로 탐색. 개발/배포/PyInstaller bundle 순."""
        candidates = [
            os.path.join(self.base_path, 'favicon.ico'),
            os.path.join(self.base_path, '..', 'img', 'favicon.ico'),
            os.path.join(self.base_path, '..', 'favicon.ico'),
        ]
        if hasattr(sys, '_MEIPASS'):
            candidates.insert(0, os.path.join(sys._MEIPASS, 'favicon.ico'))
        for p in candidates:
            if os.path.exists(p):
                return os.path.abspath(p)
        return ''

    def _patch_toplevel_icon(self):
        """이후 생성되는 모든 CTkToplevel에 자동으로 아이콘 적용 (monkey patch)."""
        icon_path = self._icon_path
        if not icon_path:
            return
        orig_init = ctk.CTkToplevel.__init__

        def patched(self_, *args, **kwargs):
            orig_init(self_, *args, **kwargs)
            try:
                self_.after(100, lambda: self_.iconbitmap(icon_path))
            except Exception:
                pass

        ctk.CTkToplevel.__init__ = patched

    def _ensure_admission_config(self):
        if not self.cfg.config.has_section('ADMISSION'):
            self.cfg.config.add_section('ADMISSION')
            self.cfg.set('ADMISSION', 'spreadsheet_id', '')
            self.cfg.set('ADMISSION', 'sheet_name', '입학식')
            self.cfg.save()

    def _load_timetable_data(self):
        tt_path = os.path.join(self.base_path, 'timetable_data.json')
        if os.path.exists(tt_path):
            with open(tt_path, 'r', encoding='utf-8') as f:
                self.tt_data = json.load(f)
        else:
            self.tt_data = {'subjects_by_grade': {}, 'english_timetable': {}}

    def write_log(self, msg: str):
        self.log.write_system(f'[입학식] {msg}')
        def _ui():
            if hasattr(self, 'status_bar') and self.status_bar.winfo_exists():
                self.status_bar.configure(text=f'  ● {msg}')
        self.after(0, _ui)

    # ──────────────────────────────────────────
    #  UI 구성
    # ──────────────────────────────────────────
    def _build_ui(self):
        # 상단 네비
        nav = ctk.CTkFrame(self, height=65, fg_color='#1a1a1a', corner_radius=0)
        nav.pack(side='top', fill='x')

        ctk.CTkLabel(
            nav, text='🎓 입학식쌤 관리 시스템',
            font=('Pretendard', 18, 'bold'), text_color='#f1c40f'
        ).pack(side='left', padx=20)

        ctk.CTkButton(
            nav, text='🔄 새로고침', width=100, fg_color='#2c3e50',
            command=self.load_students_async
        ).pack(side='left', padx=5)

        ctk.CTkButton(
            nav, text='⚙ 시트 설정', width=100, fg_color='#444444',
            command=self.open_settings
        ).pack(side='left', padx=5)

        ctk.CTkButton(
            nav, text='📊 진행 현황', width=110, fg_color='#8e44ad',
            command=self.show_progress
        ).pack(side='left', padx=15)

        # 검색
        self.search_var = ctk.StringVar()
        self.search_var.trace('w', lambda *_: self._filter_students())
        search_entry = ctk.CTkEntry(
            nav, textvariable=self.search_var,
            placeholder_text='학생 이름 검색...', width=200
        )
        search_entry.pack(side='left', padx=10)

        # 필터 (미완료만 / 전체)
        self.show_pending_only = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            nav, text='미완료만 표시',
            variable=self.show_pending_only,
            command=self._filter_students
        ).pack(side='left', padx=10)

        # 컨테이너
        container = ctk.CTkFrame(self, fg_color='transparent')
        container.pack(fill='both', expand=True, padx=10, pady=10)

        # 왼쪽: 학생 목록
        list_frame = ctk.CTkFrame(container, fg_color='#2b2b2b', width=600)
        list_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))

        # 헤더
        hdr = ctk.CTkFrame(list_frame, fg_color='#1f538d', height=40)
        hdr.pack(fill='x', padx=10, pady=(10, 0))
        for text, width in [('이름', 100), ('학년', 70), ('연락처', 130),
                             ('배정과목', 200), ('완료', 50), ('액션', 150)]:
            ctk.CTkLabel(
                hdr, text=text, width=width, font=('Pretendard', 12, 'bold')
            ).pack(side='left', padx=2)

        self.scroll_frame = ctk.CTkScrollableFrame(list_frame, fg_color='#1e1e1e')
        self.scroll_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # 오른쪽: 시간표 패널
        right_frame = ctk.CTkFrame(container, width=450, fg_color='#212121')
        right_frame.pack(side='right', fill='both')
        right_frame.pack_propagate(False)

        ctk.CTkLabel(
            right_frame, text='📅 시간표 배정 패널',
            font=('Pretendard', 16, 'bold'), text_color='#3498db'
        ).pack(pady=(20, 10))

        self.selected_student_label = ctk.CTkLabel(
            right_frame, text='← 학생을 선택하면 여기에 시간표가 표시됩니다',
            font=('Pretendard', 13), text_color='#888888'
        )
        self.selected_student_label.pack(pady=10, padx=15)

        # 시간표 배정용 과목 체크박스 영역
        self.timetable_scroll = ctk.CTkScrollableFrame(right_frame, fg_color='#1a1a1a')
        self.timetable_scroll.pack(fill='both', expand=True, padx=10, pady=5)

        # 배정 버튼들
        btn_frame = ctk.CTkFrame(right_frame, fg_color='transparent')
        btn_frame.pack(fill='x', padx=10, pady=10)

        ctk.CTkButton(
            btn_frame, text='🤖 LMS 자동 배정',
            fg_color='#8e44ad', hover_color='#732d91', height=45,
            font=('Pretendard', 13, 'bold'),
            command=self.auto_assign_selected
        ).pack(fill='x', pady=3)

        ctk.CTkButton(
            btn_frame, text='✅ 수동 완료 체크',
            fg_color='#27ae60', hover_color='#1e8449', height=40,
            command=self.manual_mark_done
        ).pack(fill='x', pady=3)

        ctk.CTkButton(
            btn_frame, text='📋 시간표 클립보드 복사',
            fg_color='#2980b9', hover_color='#1f618d', height=35,
            command=self.copy_timetable_to_clipboard
        ).pack(fill='x', pady=3)

        # 현재 선택 학생
        self.selected_student: dict | None = None
        self.timetable_checkboxes: dict[str, ctk.BooleanVar] = {}

        # 상태바
        self.status_bar = ctk.CTkLabel(
            self, text='  ● 대기 중', height=30,
            fg_color='#111111', text_color='#00FF00',
            anchor='w'
        )
        self.status_bar.pack(side='bottom', fill='x')

    # ──────────────────────────────────────────
    #  학생 목록 로드
    # ──────────────────────────────────────────
    def load_students_async(self):
        threading.Thread(target=self._load_students, daemon=True).start()

    def _load_students(self):
        try:
            self.write_log('학생 데이터 로딩 중...')
            students = self.adm_sheet.load_students()
            self.students = students
            self.after(0, self._render_students)
            self.write_log(f'로드 완료 ({len(students)}명)')
        except Exception as e:
            self.write_log(f'로드 실패: {e}')
            self.after(
                0, lambda: messagebox.showerror(
                    '로드 오류',
                    f'학생 데이터를 불러올 수 없습니다.\n'
                    f'시트 설정을 확인해주세요.\n\n{e}'
                )
            )

    def _render_students(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()
        self.student_widgets = {}
        self._filter_students()

    def _filter_students(self):
        for w in self.scroll_frame.winfo_children():
            w.destroy()

        query = self.search_var.get().strip().lower()
        pending_only = self.show_pending_only.get()

        filtered = [
            s for s in self.students
            if (not query or query in s['name'].lower())
            and (not pending_only or not s['done'])
        ]

        for student in filtered:
            self._render_student_row(student)

        if not filtered:
            ctk.CTkLabel(
                self.scroll_frame,
                text='표시할 학생이 없습니다.',
                text_color='#888888'
            ).pack(pady=20)

    def _render_student_row(self, student: dict):
        bg = '#1a3a1a' if student['done'] else '#1e1e1e'
        row_f = ctk.CTkFrame(self.scroll_frame, fg_color=bg, corner_radius=5)
        row_f.pack(fill='x', pady=2, padx=2)

        ctk.CTkLabel(row_f, text=student['name'], width=100,
                     font=('Pretendard', 13, 'bold'),
                     text_color='#aaaaaa' if student['done'] else 'white'
                     ).pack(side='left', padx=4)
        ctk.CTkLabel(row_f, text=student['grade'], width=70,
                     text_color='#888888'
                     ).pack(side='left', padx=2)
        ctk.CTkLabel(row_f, text=student['contact'], width=130,
                     text_color='#3498db'
                     ).pack(side='left', padx=2)
        subjects_text = ', '.join(student['subjects']) if student['subjects'] else '미배정'
        ctk.CTkLabel(row_f, text=subjects_text, width=200,
                     text_color='#2ecc71' if student['subjects'] else '#e74c3c',
                     wraplength=190
                     ).pack(side='left', padx=2)

        done_icon = '✅' if student['done'] else '○'
        ctk.CTkLabel(row_f, text=done_icon, width=50).pack(side='left', padx=2)

        ctk.CTkButton(
            row_f, text='배정하기', width=80, height=28,
            fg_color='#8e44ad' if not student['done'] else '#444444',
            command=lambda s=student: self.select_student(s)
        ).pack(side='right', padx=4, pady=4)

        if student['done']:
            ctk.CTkButton(
                row_f, text='되돌리기', width=65, height=28,
                fg_color='#7f8c8d', hover_color='#636e72',
                command=lambda s=student: self.revert_done(s)
            ).pack(side='right', padx=2, pady=4)

    # ──────────────────────────────────────────
    #  학생 선택 → 시간표 패널
    # ──────────────────────────────────────────
    def select_student(self, student: dict):
        self.selected_student = student
        self.selected_student_label.configure(
            text=f"선택: {student['name']} ({student['grade']})",
            text_color='#f1c40f'
        )
        self._build_timetable_panel(student)

    def _build_timetable_panel(self, student: dict):
        for w in self.timetable_scroll.winfo_children():
            w.destroy()
        self.timetable_checkboxes = {}

        grade = student['grade']
        # 학년명 정규화 (예: "초등4" → "초4", "중학교1" → "중1" 등)
        grade_key = self._normalize_grade(grade)

        grade_subjects = (
            list(self.tt_data['subjects_by_grade'].get(grade_key, {}).keys())
            if grade_key else []
        )
        english_subjects = list(self.tt_data.get('english_timetable', {}).keys())
        all_subjects = grade_subjects + english_subjects

        if not all_subjects:
            ctk.CTkLabel(
                self.timetable_scroll,
                text=f'"{grade_key}" 학년 데이터가 없습니다.\n'
                     '시간표 JSON을 확인해주세요.',
                text_color='#e74c3c'
            ).pack(pady=10)
            return

        ctk.CTkLabel(
            self.timetable_scroll,
            text='배정할 과목 선택:',
            font=('Pretendard', 13, 'bold'), text_color='#3498db'
        ).pack(anchor='w', padx=10, pady=(10, 5))

        for sub in all_subjects:
            var = ctk.BooleanVar(value=sub in student['subjects'])
            self.timetable_checkboxes[sub] = var
            ctk.CTkCheckBox(
                self.timetable_scroll, text=sub, variable=var
            ).pack(anchor='w', padx=15, pady=2)

    def _normalize_grade(self, grade: str) -> str:
        """학년 문자열을 시간표 JSON 키 형식으로 변환."""
        mapping = {
            '초등': '초', '초등학교': '초',
            '중학교': '중', '중': '중',
            '고등학교': '고등', '고등': '고등',
        }
        for k, v in mapping.items():
            if grade.startswith(k):
                rest = grade[len(k):].strip()
                num = re.search(r'\d', rest)
                if num:
                    return v + num.group()
                return v + rest
        # 이미 올바른 형식이면 그대로
        return grade

    # ──────────────────────────────────────────
    #  자동 LMS 배정
    # ──────────────────────────────────────────
    def auto_assign_selected(self):
        if not self.selected_student:
            messagebox.showwarning('선택 없음', '학생을 먼저 선택해주세요.')
            return
        if self._macro_running:
            messagebox.showwarning('진행 중', '매크로가 이미 실행 중입니다.')
            return

        selected_subs = [
            sub for sub, var in self.timetable_checkboxes.items() if var.get()
        ]
        if not selected_subs:
            messagebox.showwarning('과목 없음', '배정할 과목을 하나 이상 선택해주세요.')
            return

        name = self.selected_student['name']
        grade_key = self._normalize_grade(self.selected_student['grade'])

        # 시간표에서 배정 가능한 슬롯 자동 선택 (첫 번째 옵션)
        final_selection = {}
        for sub in selected_subs:
            sub_data = (
                self.tt_data.get('english_timetable', {}).get(sub) or
                self.tt_data.get('subjects_by_grade', {}).get(grade_key, {}).get(sub, {})
            )
            if not sub_data:
                continue
            # 첫 번째 가능한 시간대 자동 선택
            for day_key, time_list in sub_data.items():
                if time_list:
                    day = day_key[0] if len(day_key) == 2 else day_key  # '월수' → '월'
                    final_selection[(day, time_list[0])] = sub
                    break

        if not final_selection:
            messagebox.showwarning('배정 불가', '배정 가능한 시간대가 없습니다.')
            return

        self._macro_running = True
        threading.Thread(
            target=self._run_lms_assignment,
            args=(name, selected_subs, final_selection),
            daemon=True
        ).start()

    def _run_lms_assignment(self, name: str, subjects: list[str], final_selection: dict):
        driver = None
        try:
            self.write_log(f'[{name}] LMS 자동 배정 시작...')
            lms_id, lms_pw = self.cfg.get_credentials()

            driver = SeleniumManager.create_incognito()
            wait = SeleniumManager.lms_login(driver, lms_id, lms_pw)

            # 학생 상세 페이지 이동
            driver.get('https://www.lmsone.com/wcms/member/memManage/memList.asp')
            search_box = wait.until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, "input[name='keyword'], input[name='keyWord']")
                )
            )
            driver.execute_script(f"arguments[0].value = '{name}';", search_box)
            search_box.send_keys(Keys.ENTER)
            time.sleep(0.5)

            name_lower = name.strip().lower()
            target_link = None
            for link in driver.find_elements(By.CSS_SELECTOR, 'table tbody tr a'):
                lt = link.text.replace(' ', '').lower()
                if lt == name_lower or lt.startswith(name_lower + '('):
                    target_link = link
                    break

            if not target_link:
                self.write_log(f'[{name}] 학생 링크를 찾지 못했습니다.')
                return

            driver.execute_script('arguments[0].click();', target_link)
            wait.until(lambda d: len(d.window_handles) > 1)
            all_wins = driver.window_handles
            target_win = all_wins[-1]
            for w in all_wins:
                if w != target_win:
                    driver.switch_to.window(w)
                    driver.close()
            driver.switch_to.window(target_win)

            # 배정 페이지로 이동
            current_url = driver.current_url
            mid_m = re.search(r'member_id=([^&]+)', current_url)
            mseq_m = re.search(r'member_seq=([^&]+)', current_url)
            if not (mid_m and mseq_m):
                self.write_log(f'[{name}] 회원 ID/SEQ 추출 실패')
                return

            from urllib.parse import urlparse
            parsed = urlparse(current_url)
            base = f'{parsed.scheme}://{parsed.netloc}'
            assign_url = (
                f"{base}/wcms/member/memManage/tab/classSearch.asp"
                f"?member_id={mid_m.group(1)}&member_seq={mseq_m.group(1)}"
            )
            driver.get(assign_url)
            wait.until(EC.presence_of_element_located((By.ID, 'key')))

            time_map = {
                '16:40': '6481', '17:30': '6495', '18:20': '6477',
                '19:10': '6421', '20:00': '6422', '20:50': '6453',
                '21:40': '6424', '22:30': '6494', '23:20': '6701',
            }
            day_map = {
                '월': '6502', '화': '6503', '수': '6504', '목': '6505', '금': '6506'
            }

            success = 0
            for (day, ts), sub in final_selection.items():
                try:
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

                    success += 1
                    self.write_log(f'[{sub}] 배정 성공')

                except Exception as e:
                    self.write_log(f'[{sub}] 배정 실패: {e}')

            # 구글 시트에 완료 체크 + 과목 업데이트
            self.write_log(f'[{name}] 시트 완료 체크 중...')
            self.adm_sheet.update_subjects(self.selected_student['row'], subjects)
            self.adm_sheet.mark_done(self.selected_student['row'])

            # 로컬 상태 업데이트
            self.selected_student['done'] = True
            self.selected_student['subjects'] = subjects

            self.write_log(f'[{name}] 완료! (배정 {success}/{len(final_selection)}건)')
            self.after(0, self._render_students)
            self.after(
                0, lambda: messagebox.showinfo(
                    '완료', f'[{name}] LMS 배정 및 시트 체크 완료!\n'
                            f'({success}/{len(final_selection)}건 성공)'
                )
            )

        except Exception as e:
            self.write_log(f'[{name}] 배정 에러: {e}\n{traceback.format_exc()}')
            self.after(
                0, lambda: messagebox.showerror('오류', f'배정 중 오류 발생:\n{e}')
            )
        finally:
            SeleniumManager.safe_quit(driver)
            self._macro_running = False

    # ──────────────────────────────────────────
    #  수동 완료 / 되돌리기
    # ──────────────────────────────────────────
    def manual_mark_done(self):
        if not self.selected_student:
            messagebox.showwarning('선택 없음', '학생을 먼저 선택해주세요.')
            return
        name = self.selected_student['name']

        selected_subs = [
            sub for sub, var in self.timetable_checkboxes.items() if var.get()
        ]

        def _task():
            try:
                if selected_subs:
                    self.adm_sheet.update_subjects(
                        self.selected_student['row'], selected_subs
                    )
                    self.selected_student['subjects'] = selected_subs
                self.adm_sheet.mark_done(self.selected_student['row'])
                self.selected_student['done'] = True
                self.write_log(f'[{name}] 수동 완료 체크')
                self.after(0, self._render_students)
            except Exception as e:
                self.write_log(f'완료 체크 실패: {e}')
                self.after(
                    0, lambda: messagebox.showerror('오류', f'시트 업데이트 실패:\n{e}')
                )

        threading.Thread(target=_task, daemon=True).start()

    def revert_done(self, student: dict):
        if not messagebox.askyesno(
            '되돌리기', f'[{student["name"]}] 완료 체크를 해제하시겠습니까?'
        ):
            return

        def _task():
            try:
                self.adm_sheet.unmark_done(student['row'])
                student['done'] = False
                self.write_log(f"[{student['name']}] 완료 체크 해제")
                self.after(0, self._render_students)
            except Exception as e:
                self.write_log(f'되돌리기 실패: {e}')

        threading.Thread(target=_task, daemon=True).start()

    # ──────────────────────────────────────────
    #  클립보드 복사
    # ──────────────────────────────────────────
    def copy_timetable_to_clipboard(self):
        if not self.selected_student:
            messagebox.showwarning('선택 없음', '학생을 먼저 선택해주세요.')
            return

        import pyperclip
        student = self.selected_student
        grade_key = self._normalize_grade(student['grade'])
        selected_subs = [
            sub for sub, var in self.timetable_checkboxes.items() if var.get()
        ]

        if not selected_subs:
            messagebox.showwarning('과목 없음', '과목을 체크해주세요.')
            return

        result = f"[시간표 안내 - {student['name']} ({student['grade']})]\n\n"
        days_order = ['월', '화', '수', '목', '금']

        for sub in selected_subs:
            sub_data = (
                self.tt_data.get('english_timetable', {}).get(sub) or
                self.tt_data.get('subjects_by_grade', {}).get(grade_key, {}).get(sub, {})
            )
            result += f'[{sub}]\n'
            if not sub_data:
                result += ' - 배정 가능한 시간이 없습니다.\n\n'
                continue

            slots = []
            for day in days_order:
                for dk, tl in sub_data.items():
                    if day in dk:
                        for t in tl:
                            slots.append((day, t))

            t2d = {}
            for d, t in slots:
                if d not in t2d.setdefault(t, []):
                    t2d[t].append(d)

            for t in sorted(t2d):
                dl = sorted(t2d[t], key=lambda x: days_order.index(x))
                result += f' - {t} ({", ".join(dl)})\n'
            result += '\n'

        pyperclip.copy(result.strip())
        self.write_log(f"[{student['name']}] 시간표 클립보드 복사")
        messagebox.showinfo('복사 완료', '시간표가 클립보드에 복사되었습니다.')

    # ──────────────────────────────────────────
    #  진행 현황 팝업
    # ──────────────────────────────────────────
    def show_progress(self):
        total = len(self.students)
        done = sum(1 for s in self.students if s['done'])
        pending = total - done

        pop = ctk.CTkToplevel(self)
        pop.title('진행 현황')
        pop.geometry('350x300')
        pop.transient(self)
        pop.focus_force()

        pct = (done / total * 100) if total > 0 else 0

        ctk.CTkLabel(
            pop, text='📊 입학식 처리 현황',
            font=('Pretendard', 16, 'bold')
        ).pack(pady=20)

        ctk.CTkProgressBar(pop, width=280).pack(pady=5)
        bar = ctk.CTkProgressBar(pop, width=280)
        bar.set(pct / 100)
        bar.pack(pady=5)

        ctk.CTkLabel(
            pop, text=f'전체: {total}명',
            font=('Pretendard', 14)
        ).pack(pady=5)
        ctk.CTkLabel(
            pop, text=f'✅ 완료: {done}명',
            font=('Pretendard', 14), text_color='#2ecc71'
        ).pack(pady=2)
        ctk.CTkLabel(
            pop, text=f'○ 미처리: {pending}명',
            font=('Pretendard', 14), text_color='#e74c3c'
        ).pack(pady=2)
        ctk.CTkLabel(
            pop, text=f'진행률: {pct:.1f}%',
            font=('Pretendard', 16, 'bold'), text_color='#f1c40f'
        ).pack(pady=10)

    # ──────────────────────────────────────────
    #  설정
    # ──────────────────────────────────────────
    def open_settings(self):
        pop = ctk.CTkToplevel(self)
        pop.title('입학식 시트 설정')
        pop.geometry('450x350')
        pop.transient(self)
        pop.focus_force()

        ctk.CTkLabel(pop, text='[ 입학식 전용 구글 시트 ]',
                     font=('Pretendard', 14, 'bold')).pack(pady=(20, 10))

        ctk.CTkLabel(pop, text='Spreadsheet ID (비워두면 기본 시트 사용)').pack()
        sid_e = ctk.CTkEntry(pop, width=380)
        sid_e.insert(0, self.cfg.get('ADMISSION', 'spreadsheet_id', ''))
        sid_e.pack(pady=5)

        ctk.CTkLabel(pop, text='Sheet 이름').pack(pady=(10, 0))
        sn_e = ctk.CTkEntry(pop, width=200)
        sn_e.insert(0, self.cfg.get('ADMISSION', 'sheet_name', '입학식'))
        sn_e.pack(pady=5)

        ctk.CTkLabel(
            pop,
            text='※ ID를 비워두면 개통/AS 프로그램과 같은 시트 파일 사용\n'
                 '   (시트 탭 이름만 다르게 설정)',
            font=('Pretendard', 11), text_color='#888888'
        ).pack(pady=5)

        def save():
            self.cfg.set('ADMISSION', 'spreadsheet_id', sid_e.get())
            self.cfg.set('ADMISSION', 'sheet_name', sn_e.get())
            self.cfg.save()
            self.adm_sheet.invalidate()
            pop.destroy()
            self.write_log('설정 저장 완료')
            self.load_students_async()

        ctk.CTkButton(pop, text='저장', fg_color='#27ae60', command=save).pack(pady=20)


# ──────────────────────────────────────────────────
if __name__ == '__main__':
    app = AdmissionApp()
    app.mainloop()
