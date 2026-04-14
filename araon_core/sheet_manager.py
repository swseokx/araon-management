"""
araon_core/sheet_manager.py
구글 시트 연동 — 인증 캐싱 + 배치 API 호출
"""

import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials


_SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]


class SheetManager:
    """
    인증 객체와 시트 객체를 캐싱하여 매번 재인증하지 않음.
    설정이 바뀌면 invalidate()를 호출해 캐시를 초기화.
    """

    def __init__(self, config_manager):
        self.cfg = config_manager
        self._client = None
        self._sheet = None
        self._admission_sheet = None

    def invalidate(self):
        """설정 변경 후 캐시 초기화."""
        self._client = None
        self._sheet = None
        self._admission_sheet = None

    def _get_client(self):
        if self._client is not None:
            return self._client
        creds_path = os.path.join(
            self.cfg.base_path,
            self.cfg.get('DEFAULT', 'CREDENTIALS_FILE', 'credentials.json'),
        )
        creds = ServiceAccountCredentials.from_json_keyfile_name(creds_path, _SCOPES)
        self._client = gspread.authorize(creds)
        return self._client

    def _get_sheet(self):
        if self._sheet is not None:
            return self._sheet
        client = self._get_client()
        spreadsheet_id = self.cfg.get('MAIN_SHEET', 'SPREADSHEET_ID')
        sheet_name = self.cfg.get('MAIN_SHEET', 'SHEET_NAME', 'Sheet1')
        self._sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        return self._sheet

    def _get_admission_sheet(self):
        """입학식 관리 시트 (별도 스프레드시트)."""
        if self._admission_sheet is not None:
            return self._admission_sheet
        client = self._get_client()
        sheet_id = self.cfg.get('ADMISSION', 'spreadsheet_id',
                                self.cfg.get('MAIN_SHEET', 'SPREADSHEET_ID'))
        sheet_name = self.cfg.get('ADMISSION', 'sheet_name', '입학식 관리')
        self._admission_sheet = client.open_by_key(sheet_id).worksheet(sheet_name)
        return self._admission_sheet

    def load_day_data(self, selected_date: str) -> tuple[list[list], dict]:
        """
        선택 날짜 데이터 로드.
        반환: (filtered_rows, row_map)
          filtered_rows: [[col0..col15], ...]  (D열부터 S열, 0-indexed)
          row_map: {ui_index: actual_sheet_row_number}
        """
        sheet = self._get_sheet()

        # B열 전체 + D~S 범위를 배치 API로 한 번에 가져옴
        batch = sheet.batch_get(['B:B', 'D:S'])
        b_col_raw = batch[0]        # [[val], [val], ...]
        ds_raw = batch[1]           # [[d, e, ..., s], ...]

        b_col = [row[0] if row else '' for row in b_col_raw]

        # 선택 날짜 시작 행 탐색
        start_row = next(
            (i + 1 for i, val in enumerate(b_col) if selected_date in val), -1
        )
        if start_row == -1:
            return [], {}

        # 종료 행 탐색
        # B열은 날짜 헤더 행에서만 값이 있고 학생 행은 비어있어
        # gspread가 B:B를 마지막 날짜 헤더까지만 반환할 수 있음.
        # 그래서 기본값을 len(b_col)이 아닌 len(ds_raw)로 설정해야
        # 학생 행이 누락되지 않음.
        end_row = len(ds_raw)
        for i in range(start_row, len(b_col)):
            val = b_col[i].strip()
            if val and '/' in val and selected_date not in val:
                end_row = i   # 다음 날짜의 0-indexed 위치 (ds_raw와 동일 기준)
                break

        # D~S 슬라이스 (Python slice는 범위 초과를 자동으로 안전하게 처리)
        raw_data = ds_raw[start_row - 1: end_row]

        filtered_data = []
        row_map = {}
        for i, row in enumerate(raw_data):
            if len(row) < 2:
                continue
            name_raw = str(row[1])
            if not name_raw.strip() or '학생명' in name_raw or '---' in name_raw:
                continue
            row = list(row)
            row[1] = name_raw.strip()   # 양쪽 공백만 제거, 내부 공백 유지
            filtered_data.append(row)
            row_map[len(filtered_data) - 1] = start_row + i

        return filtered_data, row_map

    def mark_complete(self, actual_row: int, selected_date: str, name: str) -> bool:
        """
        해당 학생 행의 M열(13번째 열)에 'ㅇ' 기록.
        실시간으로 정확한 행을 재탐색하여 안전하게 업데이트.
        """
        sheet = self._get_sheet()
        # B열·E열만 다시 읽어 정확한 행 재탐색
        batch = sheet.batch_get(['B:B', 'E:E'])
        b_col = [row[0] if row else '' for row in batch[0]]
        e_col = [row[0] if row else '' for row in batch[1]]

        start_r = next((i for i, val in enumerate(b_col) if selected_date in val), -1)
        if start_r == -1:
            return False

        target_row = -1
        clean_name = name.replace(' ', '').strip()
        for i in range(start_r, max(len(b_col), len(e_col))):
            if i < len(b_col):
                bval = b_col[i].strip()
                if bval and '/' in bval and selected_date not in bval:
                    break
            if i < len(e_col):
                cell_name = str(e_col[i]).replace(' ', '').strip()
                if cell_name == clean_name:
                    target_row = i + 1
                    break

        if target_row == -1:
            return False

        sheet.update_cell(target_row, 13, 'ㅇ')
        return True

    def write_to_admission_sheet(self, date_str: str, rows: list[list]) -> tuple[int, int]:
        """
        입학식 등록: 선택 날짜 학생 데이터를 입학식 관리 시트에 추가.

        date_str: "4/11" 형식의 날짜
        rows: [[학생명, 학년, 입학식시간], ...] 리스트

        입학식 관리 시트 열 구조 (A~M):
          A=날짜, B=학생명, C=학년, D=입학식시간,
          E=입학식안내문자, F=카톡등록, G=레벨테스트, H=노트,
          I=첫수업일자, J=폼등록, K=전체시간표발송, L=시간표배정, M=최종확인

        반환: (추가된 수, 중복 스킵된 수)
        """
        sheet = self._get_admission_sheet()

        existing = sheet.batch_get(['A:A', 'B:B'])
        a_col = [r[0] if r else '' for r in existing[0]]
        b_col = [r[0] if r else '' for r in existing[1]]

        existing_keys = set()
        for a_val, b_val in zip(a_col, b_col):
            key = f"{str(a_val).strip()}_{str(b_val).replace(' ', '').strip()}"
            existing_keys.add(key)

        added = 0
        skipped = 0
        new_rows = []
        for row in rows:
            name  = str(row[0]).replace(' ', '').strip() if len(row) > 0 else ''
            grade = str(row[1]).strip() if len(row) > 1 else ''
            time_ = str(row[2]).strip() if len(row) > 2 else ''

            key = f"{date_str}_{name}"
            if key in existing_keys:
                skipped += 1
                continue

            # A=날짜, B=학생명, C=학년, D=입학식시간, E=안내문자, F~L=체크리스트, M=최종확인
            new_rows.append([date_str, name, grade, time_,
                              '', '', '', '', '', '', '', '', ''])
            existing_keys.add(key)
            added += 1

        if new_rows:
            sheet.append_rows(new_rows, value_input_option='USER_ENTERED')

        return added, skipped

    def update_admission_checklist(
        self, date_str: str, name: str,
        kakao: str, level_test: str, note: str, first_class: str,
        form_reg: str, timetable_send: str, schedule: str
    ) -> bool:
        """
        입학식 관리 시트에서 날짜+학생명으로 행을 찾아 체크리스트(F~M) 업데이트.
        kakao/level_test/note/form_reg/timetable_send/schedule: 'O' 또는 'X'
        first_class: 날짜 문자열 (예: '4/14')
        """
        sheet = self._get_admission_sheet()
        existing = sheet.batch_get(['A:A', 'B:B'])
        a_col = [r[0] if r else '' for r in existing[0]]
        b_col = [r[0] if r else '' for r in existing[1]]

        clean_name = name.replace(' ', '').strip()
        target_row = -1
        for i, (a_val, b_val) in enumerate(zip(a_col, b_col)):
            if (str(a_val).strip() == date_str and
                    str(b_val).replace(' ', '').strip() == clean_name):
                target_row = i + 1
                break

        if target_row == -1:
            return False

        # F=카톡, G=레벨, H=노트, I=첫수업일자, J=폼, K=전체시간표, L=시간표배정
        # M(최종확인)은 자동 기입하지 않음
        sheet.update(
            f'F{target_row}:L{target_row}',
            [[kakao, level_test, note, first_class,
              form_reg, timetable_send, schedule]],
            value_input_option='USER_ENTERED'
        )
        return True

    def mark_ot_complete(self, selected_date: str, name: str) -> bool:
        """
        신아라오티 시트에서 해당 학생의 O열(15번째 열)을 'ㅇ'으로 표시.
        selected_date: '4/14' 형식
        """
        sheet = self._get_sheet()
        batch = sheet.batch_get(['B:B', 'E:E'])
        b_col = [row[0] if row else '' for row in batch[0]]
        e_col = [row[0] if row else '' for row in batch[1]]

        start_r = next(
            (i for i, val in enumerate(b_col) if selected_date in val), -1
        )
        if start_r == -1:
            return False

        clean_name = name.strip()
        target_row = -1
        for i in range(start_r, max(len(b_col), len(e_col))):
            if i < len(b_col):
                bval = b_col[i].strip()
                if bval and '/' in bval and selected_date not in bval:
                    break
            if i < len(e_col):
                cell_name = str(e_col[i]).strip()
                if cell_name == clean_name:
                    target_row = i + 1
                    break

        if target_row == -1:
            return False

        # O열 = 15번째 열 (A=1 기준)
        sheet.update_cell(target_row, 15, 'ㅇ')
        return True

    def get_sheet_direct(self):
        """업무 팝업 등 직접 sheet 객체가 필요한 경우 반환."""
        return self._get_sheet()

    def col_values(self, col: int) -> list[str]:
        return self._get_sheet().col_values(col)
