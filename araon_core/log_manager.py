"""
araon_core/log_manager.py
통합 로그 관리 — 시스템 로그 + 셋업 로그를 단일 클래스로 처리
"""

import os
import csv
import re
from datetime import datetime


class LogManager:
    """
    로그 파일 구조:
      log/YYYY-MM-DD.log          — 시스템(디버그) 로그
      setup_log/setup_YYYY-MM-DD.log — 개통/AS 실적 로그
    """

    def __init__(self, base_path: str):
        self.base_path = base_path
        for folder in ('log', 'setup_log', 'admission_log'):
            os.makedirs(os.path.join(base_path, folder), exist_ok=True)

    # ------------------------------------------------------------------
    # 시스템 로그
    # ------------------------------------------------------------------
    def write_system(self, msg: str):
        now = datetime.now()
        log_msg = f"[{now.strftime('%H:%M:%S')}] {msg}"
        log_path = os.path.join(
            self.base_path, 'log', f"{now.strftime('%Y-%m-%d')}.log"
        )
        try:
            with open(log_path, 'a+', encoding='utf-8') as f:
                f.write(log_msg + '\n')
        except OSError:
            pass
        print(log_msg)
        return log_msg  # UI 표시용으로 반환

    def read_system_today(self) -> str:
        log_path = os.path.join(
            self.base_path, 'log', f"{datetime.now().strftime('%Y-%m-%d')}.log"
        )
        if not os.path.exists(log_path):
            return '오늘 기록된 로그 파일이 없습니다.'
        try:
            with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            return content if content.strip() else '파일은 존재하지만 내용이 비어있습니다.'
        except OSError as e:
            return f'로그 파일을 읽는 중 오류: {e}'

    # ------------------------------------------------------------------
    # 셋업(개통/AS) 로그
    # ------------------------------------------------------------------
    def write_setup(self, category: str, name: str, memo: str):
        """
        category: '개통' | 'AS'
        """
        now = datetime.now()
        memo_single = memo.replace('\n', ' ').strip()
        entry = (
            f"[{now.strftime('%H:%M:%S')}][{category}] {name} 완료"
            f" | {memo_single}\n"
        )
        log_path = os.path.join(
            self.base_path, 'setup_log', f"setup_{now.strftime('%Y-%m-%d')}.log"
        )
        try:
            with open(log_path, 'a+', encoding='utf-8') as f:
                f.write(entry)
        except OSError as e:
            self.write_system(f'셋업 로그 쓰기 실패: {e}')

    def read_setup(self, date_str: str) -> tuple[str, int, int]:
        """date_str: 'YYYY-MM-DD' → (content, open_count, as_count)"""
        log_path = os.path.join(
            self.base_path, 'setup_log', f"setup_{date_str}.log"
        )
        return self._parse_setup_file(log_path)

    def read_setup_month(self, year_month: str) -> tuple[str, int, int]:
        """year_month: 'YYYY-MM'"""
        folder = os.path.join(self.base_path, 'setup_log')
        files = sorted(
            f for f in os.listdir(folder)
            if f.startswith(f'setup_{year_month}') and f.endswith('.log')
        )
        combined, open_cnt, as_cnt = '', 0, 0
        for file in files:
            date_str = file.replace('setup_', '').replace('.log', '')
            combined += f'\n--- [{date_str}] ---\n'
            content, oc, ac = self._parse_setup_file(os.path.join(folder, file))
            combined += content
            open_cnt += oc
            as_cnt += ac
        return combined, open_cnt, as_cnt

    def read_setup_all(self) -> tuple[str, int, int]:
        folder = os.path.join(self.base_path, 'setup_log')
        files = sorted(
            f for f in os.listdir(folder)
            if f.startswith('setup_') and f.endswith('.log')
        )
        combined, open_cnt, as_cnt = '', 0, 0
        for file in files:
            date_str = file.replace('setup_', '').replace('.log', '')
            combined += f'\n--- [{date_str}] ---\n'
            content, oc, ac = self._parse_setup_file(os.path.join(folder, file))
            combined += content
            open_cnt += oc
            as_cnt += ac
        return combined, open_cnt, as_cnt

    def _parse_setup_file(self, path: str) -> tuple[str, int, int]:
        if not os.path.exists(path):
            return '', 0, 0
        content, open_cnt, as_cnt = '', 0, 0
        try:
            with open(path, 'r', encoding='utf-8') as f:
                for line in f:
                    display = (
                        line.replace(' | ', '\n  └특이사항: ')
                        if ' | ' in line else line
                    )
                    content += display
                    if '[개통]' in line:
                        open_cnt += 1
                    if '[AS]' in line:
                        as_cnt += 1
        except OSError:
            pass
        return content, open_cnt, as_cnt

    def read_setup_raw(self, date_str: str) -> str:
        """로그 편집기용 원문 반환"""
        log_path = os.path.join(
            self.base_path, 'setup_log', f"setup_{date_str}.log"
        )
        if not os.path.exists(log_path):
            return ''
        with open(log_path, 'r', encoding='utf-8') as f:
            return f.read()

    def write_setup_raw(self, date_str: str, content: str):
        """로그 편집기 저장"""
        log_path = os.path.join(
            self.base_path, 'setup_log', f"setup_{date_str}.log"
        )
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def get_setup_log_path(self, date_str: str) -> str:
        return os.path.join(self.base_path, 'setup_log', f"setup_{date_str}.log")

    # ------------------------------------------------------------------
    # 입학식 로그
    # ------------------------------------------------------------------
    def write_admission(self, ot_time: str, name: str, checklist_str: str,
                        lms_ok: bool, sheet_ok: bool):
        """
        입학식 OT 처리 결과를 admission_log/admission_YYYY-MM-DD.log 에 기록.
        checklist_str: 예) '카톡O/레벨O/노트X/첫수업4/14/폼O/시간표O/배정X'
        """
        now = datetime.now()
        lms_tag   = 'LMS✓' if lms_ok   else 'LMS✗'
        sheet_tag = '시트✓' if sheet_ok else '시트✗'
        entry = (
            f"[{now.strftime('%H:%M:%S')}][입학식][{ot_time}] {name} "
            f"| {checklist_str} | {lms_tag} {sheet_tag}\n"
        )
        log_path = os.path.join(
            self.base_path, 'admission_log',
            f"admission_{now.strftime('%Y-%m-%d')}.log"
        )
        try:
            with open(log_path, 'a+', encoding='utf-8') as f:
                f.write(entry)
        except OSError as e:
            self.write_system(f'입학식 로그 쓰기 실패: {e}')

    def read_admission_today(self) -> str:
        log_path = os.path.join(
            self.base_path, 'admission_log',
            f"admission_{datetime.now().strftime('%Y-%m-%d')}.log"
        )
        if not os.path.exists(log_path):
            return ''
        try:
            with open(log_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read()
        except OSError:
            return ''

    # ------------------------------------------------------------------
    # CSV 내보내기
    # ------------------------------------------------------------------
    def export_month_to_csv(self, year_month: str, setup_rate: int) -> str:
        """월별 실적 CSV 내보내기. 저장 경로 반환."""
        folder = os.path.join(self.base_path, 'setup_log')
        files = sorted(
            f for f in os.listdir(folder)
            if f.startswith(f'setup_{year_month}') and f.endswith('.log')
        )
        export_path = os.path.join(self.base_path, f"{year_month}_정산기록.csv")
        open_cnt, as_cnt = 0, 0
        rows = [['날짜', '분류', '이름', '특이사항']]

        for file in files:
            date_str = file.replace('setup_', '').replace('.log', '')
            with open(os.path.join(folder, file), 'r', encoding='utf-8') as f:
                for line in f:
                    if not line.strip() or line.startswith('---'):
                        continue
                    parts = line.split(' | ')
                    memo = parts[1].strip() if len(parts) > 1 else ''
                    main_part = parts[0]
                    category = (
                        '개통' if '[개통]' in main_part
                        else 'AS' if '[AS]' in main_part
                        else '기타'
                    )
                    if category == '개통':
                        open_cnt += 1
                    elif category == 'AS':
                        as_cnt += 1
                    name_match = re.search(r'\] \[(.*?)\] (.*?) 완료', main_part)
                    name = name_match.group(2).strip() if name_match else '이름모름'
                    rows.append([date_str, category, name, memo])

        total = open_cnt + as_cnt
        money = total * setup_rate
        rows.append([])
        rows.append([f'개통 {open_cnt}건', f'AS {as_cnt}건', '', ''])
        rows.append([f'총합 {money:,}원', f'(계산식: (개통+AS)*{setup_rate})', '', ''])

        with open(export_path, 'w', encoding='utf-8-sig', newline='') as f:
            csv.writer(f).writerows(rows)

        return export_path
