"""
ARAON Launcher (GitHub Releases 자동업데이트)
=============================================
사용자는 main.exe 대신 이 파일(ARAON실행.exe)을 실행하세요.

동작 순서:
  1. settings.ini 에서 GitHub repo / token 읽기
  2. GitHub API 로 최신 릴리즈 버전 확인
  3. 신버전이면 다운로드 + 압축 해제 (main.exe 가 꺼진 상태 → 자유롭게 교체)
  4. main.exe 실행 후 런처 종료
"""

import configparser
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

# PyInstaller exe 로 실행 중이면 sys.executable 이 .exe 경로
if getattr(sys, 'frozen', False):
    BASE = Path(sys.executable).resolve().parent
else:
    BASE = Path(__file__).resolve().parent

sys.path.insert(0, str(BASE))


# ── 아이콘 찾기 ──────────────────────────────────────────────────────────────
def _find_icon() -> str:
    """favicon.ico 경로 탐색. 없으면 빈 문자열."""
    candidates = [
        BASE / 'favicon.ico',
        BASE / 'img' / 'favicon.ico',
        BASE / 'bin' / 'favicon.ico',
        BASE.parent / 'img' / 'favicon.ico',
    ]
    if hasattr(sys, '_MEIPASS'):
        candidates.insert(0, Path(sys._MEIPASS) / 'favicon.ico')
    for p in candidates:
        if p.exists():
            return str(p.resolve())
    return ''


_ICON_PATH = _find_icon()


def _apply_icon(win: tk.Tk) -> None:
    if _ICON_PATH:
        try:
            win.iconbitmap(_ICON_PATH)
        except Exception:
            pass


# ── 헬퍼 ──────────────────────────────────────────────────────────────────────

def _ini_path() -> Path:
    """settings.ini 경로 탐색."""
    p = BASE / 'bin' / 'settings.ini'
    return p if p.exists() else BASE / 'settings.ini'


def _read_update_config() -> tuple[str, str]:
    """settings.ini 에서 (repo, token) 반환."""
    cfg = configparser.ConfigParser()
    cfg.read(_ini_path(), encoding='utf-8')
    repo  = cfg.get('UPDATE', 'repo',  fallback='').strip()
    token = cfg.get('UPDATE', 'token', fallback='').strip()
    return repo, token


def _notes_already_shown(version: str) -> bool:
    """이 버전의 패치노트를 이미 본 적 있으면 True."""
    cfg = configparser.ConfigParser()
    cfg.read(_ini_path(), encoding='utf-8')
    return cfg.get('UPDATE', 'notes_shown_for', fallback='').strip() == version


def _mark_notes_shown(version: str):
    """패치노트를 봤다고 settings.ini 에 기록."""
    ini = _ini_path()
    cfg = configparser.ConfigParser()
    cfg.read(ini, encoding='utf-8')
    if not cfg.has_section('UPDATE'):
        cfg.add_section('UPDATE')
    cfg.set('UPDATE', 'notes_shown_for', version)
    try:
        with open(ini, 'w', encoding='utf-8') as f:
            cfg.write(f)
    except Exception:
        pass


def _launch_main():
    """main 실행 후 런처 종료."""
    if getattr(sys, 'frozen', False):
        # 배포: bin/main.exe
        main_exe = BASE / 'bin' / 'main.exe'
        if not main_exe.exists():
            main_exe = BASE / 'main.exe'       # 구버전 호환
        subprocess.Popen([str(main_exe)])
    else:
        subprocess.Popen([sys.executable, str(BASE / 'main.py')])
    sys.exit(0)


# ── 업데이트 UI ───────────────────────────────────────────────────────────────

def _show_update_ui(info: dict, token: str):
    from araon_core import updater as upd

    ver     = info.get('version', '?')
    notes   = info.get('notes', '').strip()
    current = upd.local_version()

    # 패치노트를 이번에 처음 보는지 확인 후 표시 여부 결정
    show_notes = bool(notes) and not _notes_already_shown(ver)
    if show_notes:
        _mark_notes_shown(ver)

    # 노트 표시 여부에 따라 창 높이 조정
    win_h = 420 if show_notes else 220
    root = tk.Tk()
    root.title('ARAON Management — 업데이트')
    root.geometry(f'500x{win_h}')
    root.resizable(False, show_notes)   # 노트 있을 때만 세로 리사이즈 허용
    root.attributes('-topmost', True)
    root.configure(bg='#0f172a')
    _apply_icon(root)

    # ── 헤더 ──────────────────────────────────────────────────
    tk.Label(
        root,
        text=f'새 버전 v{ver} 이 있습니다  (현재 v{current})',
        bg='#0f172a', fg='#fbbf24',
        font=('맑은 고딕', 12, 'bold'),
    ).pack(pady=(20, 4))

    # ── 패치노트 (최초 1회만) ──────────────────────────────────
    if show_notes:
        notes_frame = tk.Frame(root, bg='#0f172a')
        notes_frame.pack(fill='both', expand=True, padx=18, pady=(2, 6))

        tk.Label(
            notes_frame, text='📋 이번 업데이트 내용',
            bg='#0f172a', fg='#94a3b8',
            font=('맑은 고딕', 9, 'bold'),
            anchor='w',
        ).pack(anchor='w', pady=(0, 3))

        txt_frame = tk.Frame(notes_frame, bg='#1e293b', bd=1, relief='flat')
        txt_frame.pack(fill='both', expand=True)

        scrollbar = tk.Scrollbar(txt_frame)
        scrollbar.pack(side='right', fill='y')

        txt = tk.Text(
            txt_frame,
            bg='#1e293b', fg='#cbd5e1',
            font=('맑은 고딕', 9),
            wrap='word',
            relief='flat',
            bd=0,
            padx=10, pady=8,
            yscrollcommand=scrollbar.set,
            state='normal',
            cursor='arrow',
        )
        txt.insert('1.0', notes)
        txt.config(state='disabled')
        txt.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=txt.yview)

    # ── 진행바 ────────────────────────────────────────────────
    style = ttk.Style(root)
    style.theme_use('clam')
    style.configure(
        'G.Horizontal.TProgressbar',
        troughcolor='#1e293b', background='#10b981',
        bordercolor='#0f172a', lightcolor='#10b981', darkcolor='#10b981',
    )
    bar = ttk.Progressbar(
        root, length=450, mode='determinate',
        style='G.Horizontal.TProgressbar',
    )
    bar.pack(pady=6, padx=24)

    status_var = tk.StringVar(value='업데이트하려면 아래 버튼을 누르세요.')
    tk.Label(
        root, textvariable=status_var,
        bg='#0f172a', fg='#94a3b8', font=('맑은 고딕', 9),
    ).pack()

    # ── 버튼 ──────────────────────────────────────────────────
    btn_frame = tk.Frame(root, bg='#0f172a')
    btn_frame.pack(pady=12)

    def do_update():
        update_btn.config(state='disabled')
        skip_btn.config(state='disabled')
        status_var.set('다운로드 중...')

        def _progress(ratio: float):
            root.after(0, lambda r=ratio: (
                bar.config(value=int(r * 100)),
                status_var.set(f'{int(r * 100)}%'),
            ))

        def _bg():
            try:
                upd.apply_update(info['download_url'], token, _progress)
                root.after(0, _done)
            except Exception as e:
                root.after(0, lambda err=e: _on_error(err))

        def _done():
            bar.config(value=100)
            status_var.set('✅ 완료! 잠시 후 앱이 실행됩니다...')
            root.after(1500, lambda: (_safe_destroy(root), _launch_main()))

        def _on_error(err):
            messagebox.showerror(
                '업데이트 오류',
                f'업데이트에 실패했습니다:\n{err}\n\n기존 버전으로 실행합니다.',
                parent=root,
            )
            _safe_destroy(root)
            _launch_main()

        threading.Thread(target=_bg, daemon=True).start()

    def do_skip():
        _safe_destroy(root)
        _launch_main()

    def _safe_destroy(win):
        try:
            win.destroy()
        except Exception:
            pass

    update_btn = tk.Button(
        btn_frame, text='  업데이트 후 실행  ',
        command=do_update,
        bg='#10b981', fg='white',
        font=('맑은 고딕', 10, 'bold'),
        relief='flat', padx=6, pady=8, cursor='hand2',
        activebackground='#059669', activeforeground='white',
        borderwidth=0,
    )
    update_btn.pack(side='left', padx=8)

    skip_btn = tk.Button(
        btn_frame, text='  이 버전으로 실행  ',
        command=do_skip,
        bg='#334155', fg='#e2e8f0',
        font=('맑은 고딕', 10),
        relief='flat', padx=6, pady=8, cursor='hand2',
        activebackground='#475569', activeforeground='white',
        borderwidth=0,
    )
    skip_btn.pack(side='left', padx=8)

    root.protocol('WM_DELETE_WINDOW', do_skip)
    root.mainloop()


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    repo, token = _read_update_config()

    if not repo:
        _launch_main()
        return

    # 버전 확인 중 로딩 창
    loading = tk.Tk()
    loading.title('ARAON Management')
    loading.geometry('280x60')
    loading.resizable(False, False)
    loading.attributes('-topmost', True)
    loading.configure(bg='#0f172a')
    _apply_icon(loading)
    tk.Label(
        loading, text='업데이트 확인 중...',
        bg='#0f172a', fg='#94a3b8', font=('맑은 고딕', 10)
    ).pack(expand=True)

    info_box: list = [None]
    done = [False]          # 중복 호출 방지 플래그

    def _check():
        from araon_core import updater as upd
        try:
            info_box[0] = upd.check_update(repo, token)
        except Exception:
            info_box[0] = None
        try:
            loading.after(0, _after_check)
        except Exception:
            pass            # 이미 destroy 된 경우

    def _after_check():
        if done[0]:
            return
        done[0] = True
        try:
            loading.destroy()
        except Exception:
            pass
        info = info_box[0]
        if info:
            _show_update_ui(info, token)
        else:
            _launch_main()

    threading.Thread(target=_check, daemon=True).start()
    # 안전망: 6초 후에도 응답이 없으면 업데이트 확인 스킵하고 앱 실행
    loading.after(6000, _after_check)
    loading.mainloop()


if __name__ == '__main__':
    main()
