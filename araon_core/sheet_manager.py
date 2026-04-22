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
        if not os.path.exists(creds_path):
            raise FileNotFoundError(
                f'credentials.json 파일이 없습니다.\n'
                f'프로그램 폴더에 credentials.json 을 넣어주세요.\n'
                f'경로: {creds_path}'
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

    @staticmethod
    def _date_full_and_dow(date_str: str) -> tuple[str, str]:
        """'4/22' → ('4/22 (수)', '(수)'). 실패 시 원본 반환."""
        from datetime import date as _date
        try:
            m_, d_ = map(int, date_str.split('/'))
            year = _date.today().year
            dt = _date(year, m_, d_)
            dow = ['월', '화', '수', '목', '금', '토', '일'][dt.weekday()]
            return f'{m_}/{d_} ({dow})', f'({dow})'
        except Exception:
            return date_str, ''

    def write_to_admission_sheet(self, date_str: str, rows: list[list]) -> tuple[int, int]:
        """
        입학식 등록: 학생 데이터를 입학식 관리 시트에 추가.

        동작:
          - C열(학생명)에 같은 이름이 이미 있으면 행 추가를 스킵(체크리스트는 별도 업데이트).
          - 새로 추가할 때 A열 날짜는 해당 날짜가 이미 시트에 있으면 생략(한 번만 표시).

        date_str: "4/22" 형식
        rows: [[학생명, 학년(무시), 입학식시간], ...]

        시트 스키마:
          A=날짜("4/22 (수)"), B=요일("(수)"), C=학생명, D=(미사용), E=입학식시간,
          F=입학식안내문자, G=카톡등록, H=레벨테스트, I=노트,
          J=첫수업일자, K=폼등록, L=전체시간표발송, M=시간표배정

        반환: (추가된 수, 중복 스킵된 수)
        """
        sheet = self._get_admission_sheet()
        date_full, dow_marker = self._date_full_and_dow(date_str)

        existing = sheet.batch_get(['A:A', 'C:C'])
        a_col = [r[0] if r else '' for r in existing[0]]
        c_col = [r[0] if r else '' for r in existing[1]]

        existing_names = {
            str(c).replace(' ', '').strip()
            for c in c_col if str(c).strip()
        }
        date_already_present = any(
            str(a).strip() == date_full for a in a_col
        )

        added = 0
        skipped = 0
        new_rows = []
        for row in rows:
            name = str(row[0]).replace(' ', '').strip() if len(row) > 0 else ''
            time_ = str(row[2]).strip() if len(row) > 2 else ''
            if not name:
                continue
            if name in existing_names:
                skipped += 1
                continue

            # 첫 번째로 이 날짜가 추가될 때만 A 채우고, 이후는 공란
            a_val = '' if date_already_present else date_full
            if not date_already_present:
                date_already_present = True

            # A, B, C, D(미사용), E(시간), F~M(체크리스트 공란)
            new_rows.append([
                a_val, dow_marker, name, '', time_,
                '', '', '', '', '', '', '', ''
            ])
            existing_names.add(name)
            added += 1

        if new_rows:
            sheet.append_rows(new_rows, value_input_option='USER_ENTERED')

        return added, skipped

    def update_admission_checklist(
        self, name: str,
        notice_msg: str,
        kakao: str, level_test: str, note: str, first_class: str,
        form_reg: str, timetable_send: str, schedule: str
    ) -> bool:
        """
        입학식 관리 시트에서 C열(학생명)으로 행을 찾아 체크리스트(F~M) 업데이트.
        notice_msg/kakao/level_test/note/form_reg/timetable_send/schedule: 'O' 또는 'X'
        first_class: 날짜 문자열 (예: '4/14')
        """
        sheet = self._get_admission_sheet()
        c_col = sheet.col_values(3)

        clean_name = name.replace(' ', '').strip()
        target_row = -1
        for i, val in enumerate(c_col):
            if str(val).replace(' ', '').strip() == clean_name:
                target_row = i + 1
                break

        if target_row == -1:
            return False

        # F=안내문자, G=카톡, H=레벨, I=노트, J=첫수업, K=폼, L=시간표발송, M=시간표배정
        sheet.update(
            f'F{target_row}:M{target_row}',
            [[notice_msg, kakao, level_test, note, first_class,
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

    def get_admission_checklist_by_names(self, names) -> dict[str, dict]:
        """
        입학식 관리 시트에서 학생명으로 체크리스트를 반환.
        날짜 필터링 없음 — 이름이 일치하는 가장 마지막(아래쪽) 행 기준.

        신 스키마: C열(idx2)=학생명, F~M(idx5~12)=체크리스트
        구 스키마: B열(idx1)=학생명, E~L(idx4~11)=체크리스트
        → 두 스키마 모두 자동 감지.

        반환: {학생명(공백제거): {'notice':..., ...}}
        """
        sheet = self._get_admission_sheet()
        all_rows = sheet.get_all_values()

        wanted = {str(n).replace(' ', '').strip() for n in names if n}
        result: dict[str, dict] = {}
        for row in all_rows:
            def _g(i, r=row):
                return r[i].strip() if i < len(r) else ''

            # 신 스키마 우선 (C열), 구 스키마 폴백 (B열)
            c_name = (row[2] if len(row) > 2 else '').replace(' ', '').strip()
            b_name = (row[1] if len(row) > 1 else '').replace(' ', '').strip()

            if c_name and c_name in wanted:
                matched_name = c_name
                offset = 5   # F=idx5 시작 (신 스키마)
            elif b_name and b_name in wanted:
                matched_name = b_name
                offset = 4   # E=idx4 시작 (구 스키마)
            else:
                continue

            # 같은 이름 여러 행 → 마지막 행이 최종값
            result[matched_name] = {
                'notice':      _g(offset),      # F(신) / E(구)
                'kakao':       _g(offset + 1),  # G / F
                'level':       _g(offset + 2),  # H / G
                'note':        _g(offset + 3),  # I / H
                'first_class': _g(offset + 4),  # J / I
                'form':        _g(offset + 5),  # K / J
                'tt_send':     _g(offset + 6),  # L / K
                'schedule':    _g(offset + 7),  # M / L
            }
        return result

    def get_admission_checklists(self, date_str: str) -> dict[str, dict]:
        """[DEPRECATED] 날짜 기반 조회 — 호환용 shim. 새 코드는 by_names 사용."""
        sheet = self._get_admission_sheet()
        all_rows = sheet.get_all_values()
        date_full, _ = self._date_full_and_dow(date_str)

        result: dict[str, dict] = {}
        for row in all_rows:
            row_date = row[0].strip() if len(row) > 0 else ''
            # 새 스키마(A="4/22 (수)") / 구 스키마(A="4/22") 모두 대응
            if row_date != date_str and row_date != date_full:
                continue
            c_name = (row[2] if len(row) > 2 else '').replace(' ', '').strip()
            if not c_name:
                continue

            def _g(i, r=row):
                return r[i].strip() if i < len(r) else ''
            result[c_name] = {
                'notice':      _g(5),
                'kakao':       _g(6),
                'level':       _g(7),
                'note':        _g(8),
                'first_class': _g(9),
                'form':        _g(10),
                'tt_send':     _g(11),
                'schedule':    _g(12),
            }
        return result

    def load_first_class_list(self) -> list[dict]:
        """
        '첫수업명단' 시트에서 학생 목록을 반환.
        E열 = 학생명, R열 = 첫수업일.
        반환: [{'name': ..., 'first_class': ...}, ...]  (헤더/빈 행 제외)
        """
        client = self._get_client()
        spreadsheet_id = self.cfg.get('MAIN_SHEET', 'SPREADSHEET_ID')
        sheet = client.open_by_key(spreadsheet_id).worksheet('첫수업명단')

        batch = sheet.batch_get(['E:E', 'R:R'])
        e_col = [row[0] if row else '' for row in batch[0]]
        r_col = [row[0] if row else '' for row in batch[1]]

        max_len = max(len(e_col), len(r_col))
        result: list[dict] = []
        for i in range(max_len):
            name = (e_col[i] if i < len(e_col) else '').strip()
            date = (r_col[i] if i < len(r_col) else '').strip()
            if not name:
                continue
            # 헤더/구분선 필터
            if ('학생명' in name or name in ('이름', '성명', '학생이름')
                    or '---' in name):
                continue
            result.append({'name': name, 'first_class': date})
        return result

    def get_sheet_direct(self):
        """업무 팝업 등 직접 sheet 객체가 필요한 경우 반환."""
        return self._get_sheet()

    def col_values(self, col: int) -> list[str]:
        return self._get_sheet().col_values(col)
