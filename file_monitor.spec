# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for file_monitor.py
# 실행: pyinstaller --noconfirm file_monitor.spec

import sys
from PyInstaller.utils.hooks import collect_all, collect_data_files

# tkinterdnd2: 패키지 전체(모듈 + tkdnd DLL 등) 수집
tmp_tkdnd = collect_all('tkinterdnd2')
datas = list(tmp_tkdnd[0])
binaries = list(tmp_tkdnd[1])
hiddenimports = list(tmp_tkdnd[2])

# customtkinter: 데이터 파일(.json, .otf 등) 수집
datas += collect_data_files('customtkinter')

# pystray: Windows 백엔드가 동적 로딩되므로 collect_all로 전체 수집
tmp_pystray = collect_all('pystray')
datas    += list(tmp_pystray[0])
binaries += list(tmp_pystray[1])
hiddenimports += list(tmp_pystray[2])

# 추가 hidden imports (동적 import 대비)
hiddenimports += [
    # pystray Windows 백엔드 (명시적 지정)
    'pystray._win32',
    # watchdog
    'watchdog',
    'watchdog.observers',
    'watchdog.observers.polling',
    'watchdog.events',
    # 표준 라이브러리
    'queue',
    # PIL
    'PIL',
    'PIL.Image',
    'PIL.ImageDraw',
    'PIL.ImageTk',
]

# pyhwpx: DLL + 서브모듈 + 데이터파일 전체 수집 (hiddenimports만으로는 DLL/데이터 누락)
try:
    import pyhwpx
    tmp_pyhwpx = collect_all('pyhwpx')
    datas    += list(tmp_pyhwpx[0])
    binaries += list(tmp_pyhwpx[1])
    hiddenimports += list(tmp_pyhwpx[2])
except ImportError:
    pass

try:
    import win32com.client
    hiddenimports += [
        'win32com.client', 'win32api', 'win32con', 'win32gui', 'pythoncom',
    ]
except ImportError:
    pass

# pyhwpx가 의존하는 패키지 (pyhwpx/core.py에서 직접 import)
hiddenimports += ['pyperclip', 'pandas', 'numpy']

# ─── 불필요한 대형 패키지 제외 (파일 크기 절감) ───────────────────────
# 주의: pandas/numpy는 pyhwpx 의존성이므로 제외하지 않음
EXCLUDES = [
    # ML / 데이터과학
    'torch', 'torchvision', 'torchaudio',
    'tensorflow', 'keras',
    'scipy', 'sklearn', 'skimage',
    'pyarrow', 'fsspec',
    'numpy.random._examples',
    'matplotlib', 'seaborn', 'plotly',
    'sympy', 'statsmodels',
    'cv2', 'imageio',
    # Jupyter / IPython
    'IPython', 'ipykernel', 'ipywidgets',
    'jupyter', 'notebook', 'nbformat', 'nbconvert',
    'traitlets', 'zmq',
    # Qt (customtkinter는 Tk 기반이므로 불필요)
    'PyQt5', 'PyQt6', 'PySide2', 'PySide6',
    # 데이터베이스
    'sqlalchemy', 'psycopg2', 'MySQLdb', 'pysqlite2',
    # 기타
    'lxml', 'openpyxl', 'xlrd', 'xlwt',
    'cryptography', 'OpenSSL',
    'jedi', 'parso',
    'pygments',
    'docutils',
    'pydantic',
    'transformers', 'tokenizers', 'huggingface_hub',
    'aiohttp', 'aiofiles',
    'pytest', 'unittest',
]

a = Analysis(
    ['file_monitor.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=EXCLUDES,
    noarchive=False,
    optimize=1,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='FileMonitor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,   # GUI 앱이므로 콘솔 창 숨김
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)
