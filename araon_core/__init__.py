# Copyright (c) 2026 swseokx. All rights reserved.

"""
araon_core 패키지.

하이브리드 로딩 전략:
- main.exe 빌드: 모든 의존성(gspread, selenium 등) 포함됨 → eager import 성공.
  PyInstaller 정적 분석이 서브모듈을 감지해 번들에 포함.
- launcher.exe 빌드: gspread 등 제외 → sheet_manager import 실패(ImportError).
  except 블록으로 넘어가 필요한 건 __getattr__ 로 lazy import.
"""

from importlib import import_module

__all__ = ['ConfigManager', 'LogManager', 'SheetManager', 'SeleniumManager']

# eager 시도 — 성공 시 PyInstaller 정적 분석이 서브모듈을 잡아서 번들에 포함됨
try:
    from .config_manager import ConfigManager          # noqa: F401
    from .log_manager import LogManager                # noqa: F401
    from .sheet_manager import SheetManager            # noqa: F401
    from .selenium_manager import SeleniumManager      # noqa: F401
except ImportError:
    # launcher 등 일부 의존성 제외 빌드에선 실패할 수 있음 → lazy fallback
    pass


_LAZY_MAP = {
    'ConfigManager':   '.config_manager',
    'LogManager':      '.log_manager',
    'SheetManager':    '.sheet_manager',
    'SeleniumManager': '.selenium_manager',
}


def __getattr__(name):
    """eager import 가 실패한 환경에서만 호출됨. 서브모듈을 필요한 시점에 로드."""
    mod_name = _LAZY_MAP.get(name)
    if mod_name is None:
        raise AttributeError(f'module {__name__!r} has no attribute {name!r}')
    mod = import_module(mod_name, __name__)
    obj = getattr(mod, name)
    globals()[name] = obj   # 캐시
    return obj
