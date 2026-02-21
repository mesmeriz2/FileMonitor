import os
import sys
import json
import re
import threading
import time
import subprocess
import logging
from datetime import datetime
from typing import Optional, Callable, Tuple
import tkinter.messagebox as messagebox
import tkinter.filedialog as filedialog

import customtkinter as ctk

# 모듈 레벨 로거 설정 (debug_mode는 설정 파일 로드 후 업데이트 가능)
logger = logging.getLogger('FileMonitor')
if not logger.handlers:
    _log_handler = logging.StreamHandler()
    _log_handler.setFormatter(logging.Formatter('%(levelname)s [FileMonitor]: %(message)s'))
    logger.addHandler(_log_handler)
logger.setLevel(logging.WARNING)  # 기본값: WARNING 이상만 출력 (debug_mode=True시 DEBUG로 변경)

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    TKDND_AVAILABLE = True
    logger.debug("tkinterdnd2 import 성공")

    # tkinterdnd2 경로를 환경 변수에 추가 (PyInstaller 호환)
    try:
        import tkinterdnd2

        # PyInstaller 환경 확인
        if getattr(sys, 'frozen', False):
            # PyInstaller 환경: _MEIPASS 경로 사용
            tkdnd_lib_path = os.path.join(sys._MEIPASS, 'tkinterdnd2', 'tkdnd')
        else:
            # 개발 환경: 설치된 패키지 경로 사용
            tkdnd_lib_path = os.path.dirname(tkinterdnd2.__file__)

        if tkdnd_lib_path not in os.environ.get('PATH', ''):
            os.environ['PATH'] = tkdnd_lib_path + os.pathsep + os.environ.get('PATH', '')
        logger.debug("tkinterdnd2 라이브러리 경로: %s", tkdnd_lib_path)
    except Exception as e:
        logger.warning("tkinterdnd2 경로 설정 오류: %s", e)

except Exception as e:
    logger.debug("tkinterdnd2 import 실패: %s (%s)", e, type(e).__name__)

    # Python 3.13 + tix 오류 확인
    python_version = sys.version_info
    if python_version.major == 3 and python_version.minor >= 13 and 'tix' in str(e):
        logger.warning(
            "Python %d.%d에서 tkinter.tix 모듈이 제거되어 tkinterdnd2가 작동하지 않습니다. "
            "'파일 선택' 버튼으로 모든 기능을 사용할 수 있습니다.",
            python_version.major, python_version.minor
        )
    else:
        logger.debug("tkinterdnd2 사용 불가 - '파일 선택' 버튼으로 모든 기능을 사용할 수 있습니다.")

    DND_FILES = None
    TkinterDnD = None
    TKDND_AVAILABLE = False

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw
import queue

# pyhwpx import (PyInstaller가 감지할 수 있도록 상단에서 import)
try:
    import pyhwpx
    PYHWPX_AVAILABLE = True
except ImportError:
    PYHWPX_AVAILABLE = False
    pyhwpx = None

# PyInstaller 호환 경로 설정
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(BASE_DIR, "config.json")

# 날짜 헤더 관련 정규식 패턴 (참조 코드에서 가져옴)
EXISTING_PREFIX_PATTERN = re.compile(r'^\d{6}[\s_\-]')
LONG_DATE_PREFIX_PATTERN = re.compile(r'^(\d{8})([\s_\-]*)(.*)')
PERIOD_DATE_PREFIX_PATTERN = re.compile(r'^(\d{4})([.\-_])(\d{2})\2(\d{2})([\s_\-]*)(.*)')
SHORT_PERIOD_DATE_PREFIX_PATTERN = re.compile(r'^(\d{2})([.\-_])(\d{2})\2(\d{2})([\s_\-]*)(.*)')
SIX_DIGIT_PREFIX_PATTERN = re.compile(r'^(\d{6})([\s_\-]+)(.*)')

ALLOWED_EXTENSIONS = {'.pdf', '.hwp', '.hwpx', '.hwpm', '.doc', '.docx',
                      '.ppt', '.pptx', '.xls', '.xlsx', '.txt', '.zip'}

# 재시도 설정
MAX_FILE_RENAME_RETRIES = 10
MAX_HWP_INIT_RETRIES = 5
MAX_HWP_CHECK_RETRIES = 3
MAX_PDF_WAIT_RETRIES = 10
MAX_HWPX_WAIT_RETRIES = 10

# 대기 시간 (초)
FILE_ACCESS_WAIT = 0.2
FILE_LOCK_WAIT = 0.3
HWP_QUIT_WAIT = 0.5
QUEUE_EMPTY_WAIT = 0.1
FILE_SIZE_CHECK_WAIT = 0.1
PDF_CONVERSION_WAIT = 0.3

# 타임아웃
QUEUE_GET_TIMEOUT = 1.0
FILE_READY_TIMEOUT = 5.0

# 파일 안정화
FILE_STABLE_COUNT = 3

# 처리 완료 파일 타임아웃
PROCESSED_FILE_TIMEOUT = 30.0
PROCESSING_FILE_TIMEOUT = 10.0


def parse_dnd_files(drop_text: str) -> list:
    """드롭된 파일 경로 문자열 파싱"""
    if not drop_text:
        return []
    
    parts = re.findall(r'\{[^}]*\}|[^\s]+', drop_text)
    filepaths = []
    for part in parts:
        item = part.strip()
        if item.startswith("{") and item.endswith("}"):
            item = item[1:-1]
        if item:
            filepaths.append(item)
    return filepaths


class ConfigManager:
    """설정 파일 관리 클래스"""
    
    DEFAULT_CONFIG = {
        "monitor_folder": "",
        "extensions": [".hwp", ".hwpx", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"],
        "pdf_output_folder": "",  # 빈 문자열이면 원본 파일과 같은 폴더에 저장
        "hwpx_converter_path": r"C:\Program Files (x86)\Hnc\HwpxConverter\HwpxConverter.exe",
        "hancom_pdf_printer": "Hancom PDF",  # 한컴 PDF 프린터 이름 (설정에서 변경 가능)
        "save_logs": False,
        "log_file_path": "monitor_log.txt",
        "window_geometry": "800x600",
        "theme": "dark",
        "debug_mode": False,
        "auto_convert_pdf": True
    }
    
    def __init__(self, config_path: str = CONFIG_PATH):
        self.config_path = config_path
        self.config = self.load_config()
    
    def load_config(self) -> dict:
        """설정 파일 로드"""
        try:
            if os.path.exists(self.config_path):
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    # 기본값으로 누락된 키 채우기
                    for key, value in self.DEFAULT_CONFIG.items():
                        if key not in config:
                            config[key] = value
                    return config
        except Exception as e:
            logger.error("설정 파일 로드 오류: %s", e)

        # 기본 설정 반환
        return self.DEFAULT_CONFIG.copy()
    
    def save_config(self):
        """설정 파일 저장"""
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error("설정 파일 저장 오류: %s", e)
    
    def get(self, key: str, default=None):
        """설정 값 가져오기"""
        return self.config.get(key, default)
    
    def set(self, key: str, value):
        """설정 값 설정 (단일 항목; 저장 포함)"""
        self.config[key] = value
        self.save_config()

    def batch_update(self, updates: dict):
        """여러 설정을 한 번에 업데이트하고 저장 (파일 I/O 1회)"""
        self.config.update(updates)
        self.save_config()


class DateHeaderProcessor:
    """날짜 헤더 처리 클래스 (참조 코드 기반)"""
    
    @staticmethod
    def _rename_with_retry(old_filepath: str, new_filepath: str) -> Tuple[Optional[str], Optional[str]]:
        """파일 이름 변경을 재시도와 함께 수행
        
        Args:
            old_filepath: 원본 파일 경로
            new_filepath: 새 파일 경로
            
        Returns:
            (새 파일명, 에러 메시지) 튜플
        """
        retry_count = 0
        while retry_count < MAX_FILE_RENAME_RETRIES:
            try:
                # 파일이 사용 가능한지 확인
                try:
                    with open(old_filepath, 'rb'):
                        pass
                except (IOError, PermissionError, OSError):
                    time.sleep(FILE_ACCESS_WAIT)
                    retry_count += 1
                    continue
                
                # 파일 이름 변경 시도
                os.rename(old_filepath, new_filepath)
                return os.path.basename(new_filepath), None
                
            except (OSError, IOError, PermissionError) as e:
                retry_count += 1
                if retry_count >= MAX_FILE_RENAME_RETRIES:
                    return None, f"파일 이름 변경 실패: {str(e)}"
                time.sleep(FILE_LOCK_WAIT)
            except Exception as e:
                return None, f"예상치 못한 오류: {str(e)}"
        
        return None, "파일 이름 변경 실패: 최대 재시도 횟수 초과"
    
    @staticmethod
    def shorten_date_prefix(filename: str) -> Optional[str]:
        """기존 날짜 헤더를 YYMMDD 형식으로 변환 (6자리 뒤 _, -는 공백으로 정규화)"""
        long_match = LONG_DATE_PREFIX_PATTERN.match(filename)
        if long_match:
            full_date = long_match.group(1)
            rest = long_match.group(3).lstrip(' \t_-')
            short_date = full_date[2:]
            return f"{short_date} {rest}"

        period_match = PERIOD_DATE_PREFIX_PATTERN.match(filename)
        if period_match:
            year = period_match.group(1)
            month = period_match.group(3)
            day = period_match.group(4)
            rest = period_match.group(6).lstrip(' \t_-')
            short_date = f"{year[2:]}{month}{day}"
            return f"{short_date} {rest}"

        short_period_match = SHORT_PERIOD_DATE_PREFIX_PATTERN.match(filename)
        if short_period_match:
            short_date = (
                short_period_match.group(1)
                + short_period_match.group(3)
                + short_period_match.group(4)
            )
            rest = short_period_match.group(6).lstrip(' \t_-')
            return f"{short_date} {rest}"

        six_digit_match = SIX_DIGIT_PREFIX_PATTERN.match(filename)
        if six_digit_match:
            rest = six_digit_match.group(3).lstrip(' \t_-')
            return f"{six_digit_match.group(1)} {rest}"

        return None
    
    @staticmethod
    def get_preferred_date(filepath: str) -> str:
        """파일의 생성/수정 시간 중 최신값을 YYMMDD 형식으로 반환"""
        created = os.path.getctime(filepath)
        modified = os.path.getmtime(filepath)
        best_time = max(created, modified)
        return datetime.fromtimestamp(best_time).strftime("%y%m%d")
    
    @staticmethod
    def add_date_prefix(filepath: str, filename: str) -> str:
        """파일명 앞에 날짜 접두사 추가"""
        date = DateHeaderProcessor.get_preferred_date(filepath)
        return f"{date} {filename}"
    
    @staticmethod
    def rename_file_with_date(filepath: str) -> Tuple[Optional[str], Optional[str]]:
        """파일에 날짜 접두사 추가 (참조 코드 기반)"""
        filename = os.path.basename(filepath)
        
        if not os.path.isfile(filepath):
            return None, "파일이 아닙니다"
        
        ext = os.path.splitext(filename)[1].lower()
        if ext not in ALLOWED_EXTENSIONS:
            return None, f"지원하지 않는 확장자: {ext}"
        
        # 이미 날짜 접두사가 있는 경우 통일된 형식으로 변환
        if EXISTING_PREFIX_PATTERN.match(filename):
            new_filename = DateHeaderProcessor.shorten_date_prefix(filename)
            if new_filename and new_filename != filename:
                new_filepath = os.path.join(os.path.dirname(filepath), new_filename)
                return DateHeaderProcessor._rename_with_retry(filepath, new_filepath)
            return None, "이미 날짜 접두사가 있습니다"
        
        # 날짜 접두사 추가
        new_filename = DateHeaderProcessor.shorten_date_prefix(filename)
        if not new_filename:
            new_filename = DateHeaderProcessor.add_date_prefix(filepath, filename)
        
        if new_filename == filename:
            return None, "변경사항이 없습니다"
        
        new_filepath = os.path.join(os.path.dirname(filepath), new_filename)
        return DateHeaderProcessor._rename_with_retry(filepath, new_filepath)


class HWPXConverter:
    """HWP → HWPX 변환 클래스"""
    
    @staticmethod
    def convert_hwp_to_hwpx(filepath: str, converter_path: str, log_callback: Optional[Callable] = None) -> Tuple[bool, Optional[str]]:
        """HWP 파일을 HWPX로 변환하고, 성공 시 원본 삭제
        
        Args:
            filepath: 변환할 HWP 파일 경로
            converter_path: HWPX 변환기 실행 파일 경로
            log_callback: 로그 콜백 함수
            
        Returns:
            (성공 여부, 결과 메시지) 튜플
        """
        if not os.path.exists(filepath):
            return False, f"파일을 찾을 수 없습니다: {filepath}"
        
        if not filepath.lower().endswith('.hwp'):
            return False, f"HWP 파일이 아닙니다: {filepath}"
        
        # HWPX 변환기 경로 확인
        if not os.path.exists(converter_path):
            return False, f"HWPX 변환기를 찾을 수 없습니다: {converter_path}"
        
        filename = os.path.basename(filepath)
        name_wo_ext = os.path.splitext(filename)[0]
        hwpx_path = os.path.join(os.path.dirname(filepath), name_wo_ext + ".hwpx")
        
        try:
            # 변환 실행
            result = subprocess.run([converter_path, filepath], capture_output=True, text=True, timeout=60)
            
            # 변환 성공 여부 확인 (파일이 생성될 때까지 대기)
            retry_count = 0
            while retry_count < MAX_HWPX_WAIT_RETRIES:
                if os.path.exists(hwpx_path):
                    # 파일이 완전히 생성되었는지 확인
                    try:
                        file_size = os.path.getsize(hwpx_path)
                        time.sleep(FILE_SIZE_CHECK_WAIT)
                        if file_size == os.path.getsize(hwpx_path):
                            # 변환 성공 시 원본 삭제
                            try:
                                os.remove(filepath)
                                return True, f"{filename} → {name_wo_ext}.hwpx"
                            except (OSError, PermissionError) as e:
                                return False, f"원본 파일 삭제 실패 ({filepath}): {str(e)}"
                    except (OSError, IOError) as e:
                        pass
                
                time.sleep(FILE_ACCESS_WAIT)
                retry_count += 1

            # 루프 종료: 최대 재시도 초과 → 변환 실패
            return False, f"변환 실패: HWPX 파일이 생성되지 않았습니다 ({hwpx_path})"
                
        except subprocess.TimeoutExpired:
            return False, f"변환 시간 초과: {filepath}"
        except FileNotFoundError:
            return False, f"변환기 실행 파일을 찾을 수 없습니다: {converter_path}"
        except (OSError, IOError) as e:
            return False, f"파일 시스템 오류: {str(e)}"
        except Exception as e:
            return False, f"예상치 못한 변환 오류 ({filepath}): {str(e)}"


class PDFConverterQueue:
    """PDF 변환 작업 큐 관리 클래스 (순차 처리)"""
    
    def __init__(self, log_callback: Optional[Callable] = None, stats_callback: Optional[Callable] = None, config: Optional['ConfigManager'] = None):
        self.queue = queue.Queue()
        self.log_callback = log_callback
        self.stats_callback = stats_callback  # 통계 업데이트 콜백
        self.config = config  # 설정 참조 (프린터 이름 등 런타임 조회용)
        self.is_processing = False
        self.processing_thread = None
        self.lock = threading.Lock()  # 동시 접근 방지
        self.com_initialized = False
        self.hwp_available = False
    
    def add_task(self, filepath: str, output_dir: Optional[str], filename: str) -> None:
        """PDF 변환 작업을 큐에 추가
        
        Args:
            filepath: 변환할 파일 경로
            output_dir: PDF 출력 디렉토리 (None이면 원본 폴더)
            filename: 파일명 (로깅용)
        """
        self.queue.put((filepath, output_dir, filename))
        self._start_processing()
    
    def _start_processing(self):
        """처리 스레드 시작 (이미 실행 중이면 무시)"""
        with self.lock:
            if not self.is_processing:
                self.is_processing = True
                self.processing_thread = threading.Thread(target=self._process_queue, daemon=True)
                self.processing_thread.start()
    
    def _initialize_com(self) -> bool:
        """COM 초기화

        Returns:
            초기화 성공 여부
        """
        try:
            import pythoncom
            try:
                pythoncom.CoInitialize()
                return True
            except pythoncom.com_error as e:
                # 이미 초기화된 경우:
                #   CO_E_ALREADYINITIALIZED = -2147221008 (0x80040110) - 같은 스레드에서 재초기화
                #   RPC_E_CHANGED_MODE = -2147417850 (0x80010106) - 다른 스레드 모델로 이미 초기화
                # 두 경우 모두 COM은 정상 사용 가능
                ALREADY_INITIALIZED_CODES = {-2147221008, -2147417850}
                error_code = e.args[0] if hasattr(e, 'args') and e.args else None
                if error_code in ALREADY_INITIALIZED_CODES:
                    return True
                # 기타 COM 오류: 실제 초기화 실패
                if self.log_callback:
                    self.log_callback(f"COM 초기화 실패 (code={error_code}): {e}", "warning")
                return False
            except Exception:
                # COM 관련이 아닌 예외: 실패로 처리
                return False
        except ImportError:
            return False  # pythoncom이 없으면 False
        except Exception:
            return False
    
    def _check_hwp_available(self) -> bool:
        """한컴오피스 사용 가능 여부 확인
        
        Returns:
            한컴오피스 사용 가능 여부
        """
        try:
            import win32com.client
            if not self.com_initialized:
                return False
            
            try:
                test_hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                test_hwp.Quit()
                time.sleep(FILE_ACCESS_WAIT)
                return True
            except Exception as e:
                if self.log_callback:
                    self.log_callback(f"한컴오피스 확인 실패: {str(e)}", "warning")
                return False
        except ImportError:
            return False
    
    def _cleanup_com(self):
        """COM 정리"""
        if self.com_initialized:
            try:
                import pythoncom
                pythoncom.CoUninitialize()
            except:
                pass
    
    def _process_queue(self):
        """큐의 작업을 순차적으로 처리"""
        # COM 초기화
        self.com_initialized = self._initialize_com()
        
        # 한컴오피스 설치 확인
        self.hwp_available = self._check_hwp_available()
        
        try:
            while True:
                try:
                    # 큐에서 작업 가져오기 (타임아웃 없이 대기)
                    filepath, output_dir, filename = self.queue.get(timeout=QUEUE_GET_TIMEOUT)
                    
                    if self.log_callback:
                        self.log_callback(f"PDF 변환 시작: {filename}", "info")
                    
                    # 한컴오피스 사용 가능 여부 확인
                    if not self.hwp_available:
                        if self.log_callback:
                            self.log_callback(f"PDF 변환 실패 ({filename}): 한컴오피스가 설치되지 않았거나 COM 접근이 불가능합니다", "error")
                        if self.stats_callback:
                            self.stats_callback("failed")
                        self.queue.task_done()
                        continue
                    
                    # PDF 변환 실행 (순차 처리 보장, 설정에서 프린터 이름 조회)
                    printer_name = self.config.get("hancom_pdf_printer", "Hancom PDF") if self.config else "Hancom PDF"
                    success, result = PDFConverter.convert_hwp_to_pdf(filepath, output_dir, skip_check=True, printer_name=printer_name)
                    
                    # 변환 후 추가 대기 (한컴오피스 완전 종료 보장)
                    time.sleep(PDF_CONVERSION_WAIT)
                    
                    if success:
                        if self.log_callback:
                            output_location = output_dir if output_dir else "원본 폴더"
                            self.log_callback(f"PDF 변환 완료: {result} ({output_location})", "success")
                        # 통계 업데이트
                        if self.stats_callback:
                            self.stats_callback("success")
                    else:
                        if self.log_callback:
                            self.log_callback(f"PDF 변환 실패 ({filename}): {result}", "error")
                        # 통계 업데이트
                        if self.stats_callback:
                            self.stats_callback("failed")
                    
                    # 작업 완료 표시
                    self.queue.task_done()
                    
                except queue.Empty:
                    # 큐가 비어있으면 잠시 대기 후 다시 확인
                    time.sleep(QUEUE_EMPTY_WAIT)
                    # lock 내에서 원자적으로 확인 후 처리 종료 (race condition 방지)
                    with self.lock:
                        if self.queue.empty():
                            self.is_processing = False
                            break
                    # 큐에 새 작업이 추가된 경우 계속 처리
                except Exception as e:
                    if self.log_callback:
                        self.log_callback(f"PDF 변환 큐 처리 오류: {str(e)}", "error")
                    if self.stats_callback:
                        self.stats_callback("failed")
                    try:
                        self.queue.task_done()
                    except:
                        pass
        finally:
            # 스레드 종료 시 COM 정리
            self._cleanup_com()


class PDFConverter:
    """PDF 변환 클래스 (참조 코드 기반)"""
    
    @staticmethod
    def convert_hwp_to_pdf(filepath: str, output_dir: Optional[str] = None, skip_check: bool = False, printer_name: str = "Hancom PDF") -> Tuple[bool, Optional[str]]:
        """HWP/HWPX 파일을 PDF로 변환"""
        if not PYHWPX_AVAILABLE or pyhwpx is None:
            return False, "pyhwpx 라이브러리가 설치되지 않았습니다"
        
        # 한컴오피스 설치 확인 (skip_check가 False일 때만 수행)
        if not skip_check:
            try:
                import win32com.client
                check_retry_count = 0
                while check_retry_count < MAX_HWP_CHECK_RETRIES:
                    try:
                        test_hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                        test_hwp.Quit()
                        time.sleep(0.2)  # 종료 대기
                        break
                    except Exception as e:
                        check_retry_count += 1
                        if check_retry_count >= MAX_HWP_CHECK_RETRIES:
                            return False, f"한컴오피스가 설치되지 않았거나 COM 접근이 불가능합니다: {str(e)}"
                        time.sleep(HWP_QUIT_WAIT)  # 재시도 전 대기
            except ImportError:
                return False, "win32com.client를 import할 수 없습니다"
        
        if output_dir is None:
            output_dir = os.path.dirname(filepath)
        else:
            # 출력 폴더가 존재하지 않으면 생성
            if not os.path.exists(output_dir):
                try:
                    os.makedirs(output_dir, exist_ok=True)
                except Exception as e:
                    return False, f"출력 폴더 생성 실패: {str(e)}"
        
        filename = os.path.basename(filepath)
        output_filename = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(output_dir, output_filename)
        
        hwp = None
        try:
            # 한컴오피스 인스턴스 생성 (재시도 로직)
            # 주의: COM 초기화는 _process_queue 스레드에서 이미 수행됨 (중복 초기화 불필요)
            init_retry_count = 0
            while init_retry_count < MAX_HWP_INIT_RETRIES:
                try:
                    hwp = pyhwpx.Hwp(new=True, visible=False)
                    break
                except Exception as e:
                    init_retry_count += 1
                    if init_retry_count >= MAX_HWP_INIT_RETRIES:
                        return False, f"한컴오피스 인스턴스 생성 실패: {str(e)}"
                    time.sleep(HWP_QUIT_WAIT)  # 이전 인스턴스가 완전히 종료될 때까지 대기
            
            if not hwp:
                return False, "한컴오피스 인스턴스를 생성할 수 없습니다"
            
            hwp.Open(filepath)
            
            # PDF 변환 액션 생성
            action = hwp.CreateAction("Print")
            pset = action.CreateSet()
            action.GetDefault(pset)
            
            # PDF 프린터 설정
            pset.SetItem("PrintMethod", 0)
            pset.SetItem("PrinterName", printer_name)
            pset.SetItem("FileName", output_path)
            pset.SetItem("SaveToFile", True)
            
            # 변환 실행
            action.Execute(pset)
            
            # 결과 파일 존재 확인 (재시도 로직)
            retry_count = 0
            while retry_count < MAX_PDF_WAIT_RETRIES:
                if os.path.exists(output_path):
                    # 파일이 완전히 생성되었는지 확인 (파일 크기가 안정화될 때까지 대기)
                    try:
                        file_size = os.path.getsize(output_path)
                        time.sleep(FILE_SIZE_CHECK_WAIT)  # 짧은 대기
                        if file_size == os.path.getsize(output_path):
                            # 파일 크기가 변하지 않으면 완전히 생성된 것으로 간주
                            return True, output_filename
                    except (OSError, IOError):
                        pass
                
                time.sleep(0.2)
                retry_count += 1
            
            # 재시도 후에도 파일이 없으면 실패
            if os.path.exists(output_path):
                return True, output_filename
            else:
                return False, "PDF 파일이 생성되지 않았습니다"
                
        except Exception as e:
            return False, str(e)
        finally:
            if hwp:
                try:
                    hwp.Quit()
                    # 한컴오피스가 완전히 종료될 때까지 대기 (중요!)
                    time.sleep(HWP_QUIT_WAIT)
                except Exception as e:
                    # 종료 중 오류는 무시하되, 다음 변환을 위해 대기
                    time.sleep(HWP_QUIT_WAIT)
                    pass
                # 주의: COM 정리는 _process_queue 스레드 종료 시 수행됨 (중복 정리 불필요)


class FileMonitorHandler(FileSystemEventHandler):
    """파일 시스템 이벤트 핸들러"""
    
    # 크롬 등 브라우저의 임시 다운로드 파일 확장자
    TEMP_EXTENSIONS = {'.crdownload', '.tmp', '.part', '.download'}
    
    def __init__(self, callback: Callable, extensions: list):
        super().__init__()
        self.callback = callback
        self.extensions = [ext.lower() for ext in extensions]
        self.processing_files = set()  # 중복 처리 방지
        self.processed_files = set()  # 처리 완료된 파일 (재감지 방지)
        self._files_lock = threading.Lock()  # processing_files 스레드 안전성 보장
    
    def _should_process_file(self, filepath: str) -> bool:
        """파일을 처리해야 하는지 확인"""
        if not os.path.exists(filepath):
            return False
        
        filename = os.path.basename(filepath)
        
        # Office 임시 파일 무시 (~$로 시작하는 파일)
        if filename.startswith('~$'):
            return False
        
        # 임시 파일 확장자 무시
        ext = os.path.splitext(filepath)[1].lower()
        if ext in self.TEMP_EXTENSIONS:
            return False
        
        # 확장자 확인
        if ext not in self.extensions:
            return False
        
        # 이미 날짜 접두사가 있는 파일은 처리하지 않음 (재감지 방지)
        if EXISTING_PREFIX_PATTERN.match(filename):
            return False
        
        # 중복 처리 방지 및 처리 완료 파일 확인 (락 없이 빠른 경로 확인)
        with self._files_lock:
            if filepath in self.processing_files:
                return False
            if filepath in self.processed_files:
                return False

        return True

    def _discard_processed(self, filepath: str):
        """처리 완료 목록에서 안전하게 제거"""
        with self._files_lock:
            self.processed_files.discard(filepath)

    def _discard_processing(self, filepath: str):
        """처리 중 목록에서 안전하게 제거"""
        with self._files_lock:
            self.processing_files.discard(filepath)

    def _wait_for_file_ready(self, filepath: str, max_wait_seconds: float = FILE_READY_TIMEOUT) -> bool:
        """파일이 완전히 생성되고 안정화될 때까지 대기"""
        start_time = time.time()
        last_size = -1
        stable_count = 0
        
        while time.time() - start_time < max_wait_seconds:
            if not os.path.exists(filepath):
                time.sleep(0.1)
                continue
            
            try:
                # 파일 크기 확인
                current_size = os.path.getsize(filepath)
                
                # 파일이 열려있는지 확인
                try:
                    with open(filepath, 'rb'):
                        pass
                except (IOError, PermissionError, OSError):
                    time.sleep(0.2)
                    continue
                
                # 파일 크기가 안정화되었는지 확인
                if current_size == last_size:
                    stable_count += 1
                    if stable_count >= FILE_STABLE_COUNT:
                        return True
                else:
                    stable_count = 0
                    last_size = current_size
                
                time.sleep(0.1)
            except (OSError, IOError):
                time.sleep(0.1)
                continue
        
        # 최대 대기 시간 내에 파일이 준비되지 않았지만 존재하면 처리 시도
        return os.path.exists(filepath)
    
    def _process_file(self, filepath: str):
        """파일 처리 (공통 로직)"""
        if not self._should_process_file(filepath):
            return

        # 파일이 완전히 준비될 때까지 대기
        if not self._wait_for_file_ready(filepath):
            return

        ext = os.path.splitext(filepath)[1].lower()

        # 원자적 check-then-add로 중복 처리 방지 (스레드 안전)
        with self._files_lock:
            if filepath in self.processing_files:
                return
            self.processing_files.add(filepath)

        # 콜백 호출 (별도 스레드에서)
        if self.callback:
            def callback_wrapper():
                try:
                    # 콜백 실행
                    self.callback(filepath, ext)
                    # 처리 완료된 파일로 표시 (재감지 방지)
                    with self._files_lock:
                        self.processed_files.add(filepath)
                    # 일정 시간 후 처리 완료 목록에서 제거 (파일명 변경 후 재감지 방지 시간)
                    threading.Timer(PROCESSED_FILE_TIMEOUT, lambda: self._discard_processed(filepath)).start()
                finally:
                    # 처리 중 목록에서 제거
                    with self._files_lock:
                        self.processing_files.discard(filepath)

            threading.Thread(target=callback_wrapper, daemon=True).start()
        else:
            # 콜백이 없으면 처리 중 목록에서만 제거
            threading.Timer(PROCESSING_FILE_TIMEOUT, lambda: self._discard_processing(filepath)).start()
    
    def on_created(self, event):
        """파일 생성 이벤트 처리"""
        if event.is_directory:
            return
        
        filepath = event.src_path
        self._process_file(filepath)
    
    def on_moved(self, event):
        """파일 이동/이름 변경 이벤트 처리 (크롬 다운로드 완료 시 발생)"""
        if event.is_directory:
            return
        
        # 크롬 등은 임시 파일(.crdownload)을 최종 파일명으로 변경
        # event.dest_path가 최종 파일 경로
        filepath = event.dest_path
        self._process_file(filepath)


class FileMonitor:
    """파일 모니터링 클래스"""
    
    def __init__(self, config: ConfigManager, log_callback: Optional[Callable] = None):
        self.config = config
        self.log_callback = log_callback
        self.observer: Optional[Observer] = None
        self.is_monitoring = False
        self.stats = {"success": 0, "failed": 0}
        # PDF 변환 큐 초기화 (순차 처리)
        self.pdf_queue = PDFConverterQueue(
            log_callback=log_callback,
            stats_callback=self._update_stats,
            config=config
        )
    
    def _update_stats(self, result: str):
        """통계 업데이트 (PDF 변환 큐에서 호출)"""
        if result == "success":
            self.stats["success"] += 1
        elif result == "failed":
            self.stats["failed"] += 1
    
    def start_monitoring(self, folder_path: str) -> bool:
        """모니터링 시작
        
        Args:
            folder_path: 모니터링할 폴더 경로
            
        Returns:
            모니터링 시작 성공 여부
        """
        if self.is_monitoring:
            self.stop_monitoring()
        
        if not os.path.exists(folder_path):
            if self.log_callback:
                self.log_callback(f"오류: 폴더를 찾을 수 없습니다: {folder_path}", "error")
            return False
        
        try:
            self.event_handler = FileMonitorHandler(
                callback=self.process_file,
                extensions=self.config.get("extensions", [])
            )
            
            self.observer = Observer()
            self.observer.schedule(self.event_handler, folder_path, recursive=False)
            self.observer.start()
            self.is_monitoring = True
            
            if self.log_callback:
                self.log_callback(f"모니터링 시작: {folder_path}", "info")
            
            return True
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"모니터링 시작 실패: {str(e)}", "error")
            return False
    
    def stop_monitoring(self):
        """모니터링 중지"""
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
        self.is_monitoring = False
        if self.log_callback:
            self.log_callback("모니터링 중지", "info")
    
    def process_existing_files(self, folder_path: str):
        """기존 파일들을 1회 처리 (모니터링 없이)"""
        if not os.path.exists(folder_path):
            if self.log_callback:
                self.log_callback(f"오류: 폴더를 찾을 수 없습니다: {folder_path}", "error")
            return
        
        extensions = self.config.get("extensions", [])
        extensions_lower = [ext.lower() for ext in extensions]
        
        if self.log_callback:
            self.log_callback(f"기존 파일 처리 시작: {folder_path}", "info")
        
        # 폴더의 모든 파일 스캔
        try:
            files = os.listdir(folder_path)
            target_files = []
            
            for filename in files:
                filepath = os.path.join(folder_path, filename)
                if not os.path.isfile(filepath):
                    continue
                
                ext = os.path.splitext(filename)[1].lower()
                if ext in extensions_lower:
                    target_files.append(filepath)
            
            if not target_files:
                if self.log_callback:
                    self.log_callback("처리할 파일이 없습니다.", "info")
                return
            
            if self.log_callback:
                self.log_callback(f"총 {len(target_files)}개 파일 처리 시작", "info")
            
            # 각 파일 처리
            for filepath in target_files:
                ext = os.path.splitext(filepath)[1].lower()
                self.process_file(filepath, ext)
            
            if self.log_callback:
                self.log_callback(f"기존 파일 처리 완료: {len(target_files)}개 파일", "success")
                
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"기존 파일 처리 오류: {str(e)}", "error")
    
    def process_file(self, filepath: str, ext: str):
        """파일 처리 (날짜 헤더 추가 + HWP→HWPX 변환)"""
        filename = os.path.basename(filepath)
        
        # 이미 날짜 접두사가 있는 파일은 처리하지 않음 (재감지 방지)
        if EXISTING_PREFIX_PATTERN.match(filename):
            return
        
        if self.log_callback:
            self.log_callback(f"파일 감지: {filename}", "info")
        
        # 날짜 헤더 추가
        try:
            new_filename, error = DateHeaderProcessor.rename_file_with_date(filepath)
            if error:
                # "이미 날짜 접두사가 있습니다" 오류는 조용히 무시 (재감지 방지)
                if "이미 날짜 접두사가 있습니다" not in error:
                    if self.log_callback:
                        self.log_callback(f"날짜 헤더 추가 실패 ({filename}): {error}", "warning")
            elif new_filename:
                if self.log_callback:
                    self.log_callback(f"날짜 헤더 추가 완료: {filename} → {new_filename}", "success")
                # 파일명이 변경되었으므로 경로 업데이트
                new_filepath = os.path.join(os.path.dirname(filepath), new_filename)
                # 변경된 파일 경로를 처리 완료 목록에 추가하여 재감지 방지
                if hasattr(self, 'event_handler') and self.event_handler:
                    self.event_handler.processed_files.add(new_filepath)
                    # 일정 시간 후 처리 완료 목록에서 제거
                    threading.Timer(PROCESSED_FILE_TIMEOUT, lambda: self.event_handler.processed_files.discard(new_filepath)).start()
                filepath = new_filepath
                filename = new_filename
                self.stats["success"] += 1
        except Exception as e:
            if self.log_callback:
                self.log_callback(f"날짜 헤더 추가 오류 ({filename}): {str(e)}", "error")
            self.stats["failed"] += 1
        
        # HWP → HWPX 변환
        if ext.lower() == '.hwp':
            try:
                if self.log_callback:
                    self.log_callback(f"HWPX 변환 시작: {filename}", "info")
                
                converter_path = self.config.get("hwpx_converter_path", "")
                if not converter_path:
                    if self.log_callback:
                        self.log_callback(f"HWPX 변환 실패 ({filename}): 변환기 경로가 설정되지 않았습니다", "error")
                    self.stats["failed"] += 1
                    return
                
                success, result = HWPXConverter.convert_hwp_to_hwpx(filepath, converter_path, self.log_callback)
                
                if success:
                    if self.log_callback:
                        self.log_callback(f"HWPX 변환 완료: {result}", "success")
                    self.stats["success"] += 1
                else:
                    if self.log_callback:
                        self.log_callback(f"HWPX 변환 실패 ({filename}): {result}", "error")
                    self.stats["failed"] += 1
            except Exception as e:
                if self.log_callback:
                    self.log_callback(f"HWPX 변환 오류 ({filename}): {str(e)}", "error")
                self.stats["failed"] += 1


class LogQueue:
    """로그 큐 클래스 (스레드 안전)"""
    
    def __init__(self):
        self.queue = queue.Queue()
    
    def put(self, message: str, level: str = "info"):
        """로그 추가"""
        self.queue.put((message, level, datetime.now()))
    
    def get_all(self):
        """모든 로그 가져오기 (큐 비우기)"""
        logs = []
        while not self.queue.empty():
            try:
                logs.append(self.queue.get_nowait())
            except queue.Empty:
                break
        return logs


if TKDND_AVAILABLE:
    class DnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
        """드래그 앤 드롭을 지원하는 CTk 루트"""
        
        def __init__(self, *args, **kwargs):
            logger.debug("DnDCTk 초기화 시작")

            # CTk 초기화
            ctk.CTk.__init__(self, *args, **kwargs)
            logger.debug("CTk 초기화 완료")

            # TkinterDnD.DnDWrapper 초기화
            try:
                TkinterDnD.DnDWrapper.__init__(self)
                logger.debug("TkinterDnD.DnDWrapper 초기화 완료")
            except Exception as e:
                logger.error("TkinterDnD.DnDWrapper 초기화 오류: %s", e)

            # 간단한 tkdnd 패키지 확인 (상세 로드는 _ensure_tkdnd_loaded에서)
            try:
                self.TkdndVersion = self.tk.call('package', 'require', 'tkdnd')
                logger.debug("tkdnd 버전 %s 초기 로드 성공", self.TkdndVersion)
            except Exception as e:
                logger.debug("tkdnd 초기 로드 실패 (나중에 재시도): %s", e)
else:
    class DnDCTk(ctk.CTk):
        """드래그 앤 드롭 비활성 CTk 루트"""
        
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)


class MonitorApp(DnDCTk):
    """메인 애플리케이션 클래스"""
    
    def __init__(self):
        super().__init__()
        
        # 설정
        self.config_manager = ConfigManager()
        ctk.set_appearance_mode(self.config_manager.get("theme", "dark"))
        ctk.set_default_color_theme("blue")
        
        # 상태
        self.monitor = None
        self.log_queue = LogQueue()
        self.tray_icon = None
        self.tray_thread = None
        self._log_timer_id = None  # update_logs 타이머 ID (종료 시 cancel 용)

        # UI 초기화
        self.setup_ui()
        self.setup_tray()

        # 로그 업데이트 타이머
        self._log_timer_id = self.after(100, self.update_logs)
    
    def setup_ui(self):
        """UI 설정"""
        self.title("파일 모니터링 및 자동 처리")
        geometry = self.config_manager.get("window_geometry", "800x600")
        self.geometry(geometry)
        self.minsize(600, 400)
        
        # 메인 컨테이너
        main_container = ctk.CTkFrame(self)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # 상단: 상태 표시 영역
        status_frame = ctk.CTkFrame(main_container)
        status_frame.pack(fill="x", pady=(0, 10))
        
        # 첫 번째 줄: 상태 및 폴더 경로
        status_info_frame = ctk.CTkFrame(status_frame)
        status_info_frame.pack(fill="x", padx=5, pady=(5, 0))
        
        # 상태 표시
        self.status_label = ctk.CTkLabel(
            status_info_frame,
            text="● 중지됨",
            text_color="gray",
            font=ctk.CTkFont(size=16, weight="bold")
        )
        self.status_label.pack(side="left", padx=10, pady=5)

        # 처리 통계 인라인 표시
        self.stats_label = ctk.CTkLabel(
            status_info_frame, text="", font=ctk.CTkFont(size=12)
        )
        self.stats_label.pack(side="left", padx=10, pady=5)

        # 폴더 경로 표시 (줄바꿈 가능하도록 설정)
        self.folder_label = ctk.CTkLabel(
            status_info_frame,
            text="폴더: 미설정",
            font=ctk.CTkFont(size=12),
            anchor="w",
            justify="left"
        )
        self.folder_label.pack(side="left", padx=10, pady=5, fill="x", expand=True)
        
        # 두 번째 줄: 버튼 영역
        button_frame = ctk.CTkFrame(status_frame)
        button_frame.pack(fill="x", padx=5, pady=(5, 5))
        
        # 버튼 영역 (오른쪽 정렬, 순서: 설정 - PDF 변환 - 1회 실행 - 시작)
        button_width = 100
        
        # 시작/중지 버튼
        self.toggle_button = ctk.CTkButton(
            button_frame,
            text="시작",
            command=self.toggle_monitoring,
            width=button_width
        )
        self.toggle_button.pack(side="right", padx=5, pady=5)
        
        # 1회 실행 버튼
        self.once_button = ctk.CTkButton(
            button_frame,
            text="1회 실행",
            command=self.process_existing_files_once,
            width=button_width,
            fg_color="gray",
            hover_color="darkgray"
        )
        self.once_button.pack(side="right", padx=5, pady=5)
        
        # PDF 변환 버튼
        self.pdf_button = ctk.CTkButton(
            button_frame,
            text="PDF 변환",
            command=self.process_pdf_conversion_once,
            width=button_width,
            fg_color="purple",
            hover_color="darkviolet"
        )
        self.pdf_button.pack(side="right", padx=5, pady=5)
        
        # 설정 버튼
        settings_button = ctk.CTkButton(
            button_frame,
            text="⚙️ 설정",
            command=self.open_settings,
            width=button_width
        )
        settings_button.pack(side="right", padx=5, pady=5)

        # 드롭 영역
        self.drop_frame = ctk.CTkFrame(
            main_container,
            border_width=2,
            border_color="gray50",
            fg_color=("gray90", "gray20")
        )
        self.drop_frame.pack(fill="x", pady=(0, 10))

        # 드롭 영역 레이블과 버튼을 담을 컨테이너
        drop_content_frame = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        drop_content_frame.pack(fill="x", padx=10, pady=20)
        
        self.drop_label = ctk.CTkLabel(
            drop_content_frame,
            text="파일 선택 버튼을 사용하거나 파일을 드롭하세요",
            font=ctk.CTkFont(size=12, weight="bold")
        )
        self.drop_label.pack(side="left", fill="x", expand=True)
        
        # 파일 선택 버튼 추가 (drag-drop 대체)
        self.select_files_button = ctk.CTkButton(
            drop_content_frame,
            text="📁 파일 선택",
            command=self.select_files_for_pdf,
            width=120,
            fg_color="#2fa572",
            hover_color="#28a868"
        )
        self.select_files_button.pack(side="right", padx=(10, 0))
        
        # 드롭 타겟 설정 시도 (사용 가능한 경우)
        self.setup_drop_target()
        
        # 하단: 로그 패널
        log_frame = ctk.CTkFrame(main_container)
        log_frame.pack(fill="both", expand=True)
        
        # 로그 헤더
        log_header = ctk.CTkFrame(log_frame)
        log_header.pack(fill="x")
        
        log_title = ctk.CTkLabel(
            log_header,
            text="📋 로그",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        log_title.pack(side="left", padx=10, pady=5)
        
        self.log_toggle_button = ctk.CTkButton(
            log_header,
            text="접기",
            command=self.toggle_log_panel,
            width=60,
            height=25
        )
        self.log_toggle_button.pack(side="right", padx=10, pady=5)

        # 자동 스크롤 체크박스
        self.auto_scroll_var = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            log_header, text="자동 스크롤", variable=self.auto_scroll_var,
            width=90, height=25
        ).pack(side="right", padx=5, pady=5)

        # 지우기 버튼
        ctk.CTkButton(
            log_header, text="지우기", command=self.clear_log,
            width=55, height=25, fg_color="gray40", hover_color="gray30"
        ).pack(side="right", padx=3, pady=5)

        # 저장 버튼
        ctk.CTkButton(
            log_header, text="저장", command=self.save_log_to_file,
            width=55, height=25
        ).pack(side="right", padx=3, pady=5)

        # 로그 텍스트 박스
        self.log_textbox = ctk.CTkTextbox(
            log_frame,
            font=ctk.CTkFont(size=11),
            wrap="word"
        )
        self.log_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.log_panel_visible = True
        
        # 로그 컬러 태그 설정
        self.log_textbox.tag_config("info", foreground="#a0a0a0")  # 회색
        self.log_textbox.tag_config("success", foreground="#4caf50")  # 녹색
        self.log_textbox.tag_config("warning", foreground="#ff9800")  # 주황색
        self.log_textbox.tag_config("error", foreground="#f44336")  # 빨간색
        
        # 초기 로그
        self.add_log("애플리케이션이 시작되었습니다.", "info")
    
    def setup_tray(self):
        """시스템 트레이 설정"""
        try:
            # 트레이 아이콘 이미지 생성
            image = Image.new('RGB', (64, 64), color='#1a73e8')
            draw = ImageDraw.Draw(image)
            # 폴더 아이콘 모양 그리기
            draw.rectangle([20, 20, 44, 44], fill='white', outline='#1a73e8', width=2)
            draw.rectangle([20, 20, 32, 28], fill='#1a73e8')
            
            # 트레이 메뉴 (메인 스레드에서 실행되도록 래핑)
            menu = pystray.Menu(
                pystray.MenuItem("창 표시", lambda: self.after(0, self.show_window)),
                pystray.MenuItem("시작/중지", lambda: self.after(0, self.toggle_monitoring)),
                pystray.MenuItem("설정", lambda: self.after(0, self.open_settings)),
                pystray.Menu.SEPARATOR,
                pystray.MenuItem("종료", lambda: self.after(0, self.quit_app))
            )
            
            self.tray_icon = pystray.Icon("FileMonitor", image, "파일 모니터링", menu)
            
            # 트레이 스레드 시작
            self.tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            self.tray_thread.start()
        except Exception as e:
            logger.error("시스템 트레이 설정 오류: %s", e)
            self.tray_icon = None
    
    def _find_tkdnd_paths(self) -> list:
        """tkdnd 라이브러리 경로 후보 찾기"""
        if not TKDND_AVAILABLE:
            return []
        
        try:
            import tkinterdnd2
            base_dir = os.path.dirname(tkinterdnd2.__file__)
            if not os.path.isdir(base_dir):
                return []
            
            candidates = []
            for name in os.listdir(base_dir):
                if name.lower().startswith("tkdnd"):
                    path = os.path.join(base_dir, name)
                    if os.path.isdir(path):
                        candidates.append(path)
                        # 플랫폼별 경로 추가
                        platform_dir = self._get_tkdnd_platform_dir(path)
                        if platform_dir:
                            candidates.append(platform_dir)
            return candidates
        except Exception:
            return []
    
    def _get_tkdnd_platform_dir(self, tkdnd_root: str) -> Optional[str]:
        """현재 플랫폼에 맞는 tkdnd 하위 경로 반환"""
        try:
            import platform
            system = sys.platform.lower()
            machine = platform.machine().lower()
            is_64bit = sys.maxsize > 2**32
            
            if system.startswith("win"):
                if "arm" in machine:
                    name = "win-arm64"
                elif is_64bit:
                    name = "win-x64"
                else:
                    name = "win-x86"
            elif system.startswith("linux"):
                name = "linux-arm64" if "arm" in machine else "linux-x64"
            elif system.startswith("darwin"):
                name = "osx-arm64" if "arm" in machine else "osx-x64"
            else:
                return None
            
            candidate = os.path.join(tkdnd_root, name)
            return candidate if os.path.isdir(candidate) else None
        except Exception:
            return None
    
    def _ensure_tkdnd_loaded(self) -> bool:
        """tkdnd 패키지 로드 시도 (플랫폼 맞춤 버전)"""
        if not TKDND_AVAILABLE:
            logger.warning("tkinterdnd2 라이브러리를 사용할 수 없습니다.")
            return False

        try:
            import tkinterdnd2

            tkdnd_base_path = os.path.dirname(tkinterdnd2.__file__)
            logger.debug("tkinterdnd2 설치 경로: %s", tkdnd_base_path)

            # tkdnd 폴더 찾기
            tkdnd_root = None
            for item in os.listdir(tkdnd_base_path):
                if item.lower().startswith('tkdnd'):
                    full_path = os.path.join(tkdnd_base_path, item)
                    if os.path.isdir(full_path):
                        tkdnd_root = full_path
                        logger.debug("tkdnd 폴더 발견: %s", item)
                        break

            if not tkdnd_root:
                logger.error("tkdnd 폴더를 찾을 수 없습니다.")
                return False

            # _get_tkdnd_platform_dir로 플랫폼별 경로 결정 (중복 로직 제거)
            platform_path = self._get_tkdnd_platform_dir(tkdnd_root)
            if not platform_path:
                logger.error("현재 플랫폼에 맞는 tkdnd 경로를 찾을 수 없습니다.")
                return False

            logger.debug("tkdnd 플랫폼 경로: %s", platform_path)

            # Tcl auto_path에 플랫폼별 경로만 추가 (루트 경로는 추가하지 않음!)
            try:
                self.tk.call("lappend", "auto_path", platform_path)
            except Exception as e:
                logger.error("Tcl auto_path 추가 실패 (%s): %s", platform_path, e)
                return False

            # 방법 1: 일반 로드
            try:
                version = self.tk.eval("package require tkdnd")
                logger.debug("tkdnd 버전 %s 로드 성공 (방법 1)", version)
                return True
            except Exception as e:
                logger.debug("tkdnd 로드 방법 1 실패: %s", e)

            # 방법 2: pkgIndex.tcl을 올바른 컨텍스트에서 로드
            try:
                pkg_index_path = os.path.join(platform_path, "pkgIndex.tcl")
                if os.path.exists(pkg_index_path):
                    tcl_platform_path = platform_path.replace('\\', '/')
                    self.tk.eval(f'set dir "{tcl_platform_path}"')
                    tcl_pkg_index = pkg_index_path.replace('\\', '/')
                    self.tk.eval(f'source "{tcl_pkg_index}"')
                    version = self.tk.eval("package require tkdnd")
                    logger.debug("tkdnd 버전 %s 로드 성공 (방법 2: pkgIndex.tcl)", version)
                    return True
            except Exception as e:
                logger.debug("tkdnd 로드 방법 2 실패: %s", e)

            # 방법 3: DLL 직접 로드
            try:
                dll_path = os.path.join(platform_path, "libtkdnd2.9.4.dll")
                if os.path.exists(dll_path):
                    tcl_dll_path = dll_path.replace('\\', '/')
                    self.tk.eval(f'load "{tcl_dll_path}" tkdnd')
                    logger.debug("tkdnd 로드 성공 (방법 3: DLL 직접 로드)")
                    return True
            except Exception as e:
                logger.debug("tkdnd 로드 방법 3 실패: %s", e)

            logger.error("tkdnd 모든 로드 방법 실패")
            return False

        except Exception as e:
            logger.error("_ensure_tkdnd_loaded 전체 오류: %s", e)
            return False

    def _try_load_tkdnd_from_path(self, path: str) -> bool:
        """특정 경로에서 tkdnd 패키지 로드 시도"""
        try:
            self.tk.call("lappend", "auto_path", path)
            try:
                self.tk.eval("package require tkdnd")
                return True
            except Exception:
                pkg_index = os.path.join(path, "pkgIndex.tcl")
                if os.path.exists(pkg_index):
                    self.tk.eval(f"source {{{pkg_index}}}")
                    self.tk.eval("package require tkdnd")
                    return True
        except Exception:
            return False
        
        return False
    
    def setup_drop_target(self):
        """드롭 영역 등록"""
        logger.debug("setup_drop_target 시작")

        if not TKDND_AVAILABLE:
            logger.warning("tkinterdnd2를 import할 수 없습니다.")
            self.drop_label.configure(text="파일 선택 버튼을 사용하세요 (드롭 기능 비활성)")
            return

        try:
            if not self._ensure_tkdnd_loaded():
                self.drop_label.configure(text="파일 선택 버튼을 사용하세요 (드롭 기능 비활성)")
                self.add_log("드롭 기능이 비활성화되었습니다. tkdnd 패키지를 로드할 수 없습니다.", "warning")
                return

            # 루트 윈도우 전체를 드롭 타겟으로 등록
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self.handle_drop)
            self.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            self.dnd_bind("<<DragLeave>>", self._on_drag_leave)

            self.drop_label.configure(text="HWP/HWPX 파일을 여기로 드롭 (창 전체)")
            self.add_log("드롭 기능이 활성화되었습니다.", "success")
            logger.debug("드롭 기능 활성화 완료")

        except Exception as e:
            # 실패 시 명확한 오류 메시지
            self.drop_label.configure(text="파일 선택 버튼을 사용하세요 (드롭 기능 비활성)")
            self.add_log(f"드롭 기능 초기화 실패: {str(e)}", "error")
            logger.error("드롭 초기화 실패: %s", e)
    
    def handle_drop(self, event):
        """드롭된 파일을 PDF 변환 큐에 추가"""
        filepaths = parse_dnd_files(getattr(event, "data", ""))
        self._on_drag_leave(event)
        if not filepaths:
            self.add_log("드롭된 파일이 없습니다.", "warning")
            return

        self._process_files_for_pdf(filepaths, source="드롭")
    
    def select_files_for_pdf(self):
        """파일 선택 다이얼로그를 열어 PDF 변환할 파일 선택"""
        filepaths = filedialog.askopenfilenames(
            title="PDF로 변환할 파일 선택",
            filetypes=[
                ("한글 파일", "*.hwp *.hwpx"),
                ("모든 파일", "*.*")
            ]
        )
        
        if not filepaths:
            return
        
        self._process_files_for_pdf(list(filepaths), source="선택")
    
    def _process_files_for_pdf(self, filepaths: list, source: str = ""):
        """파일 목록을 PDF 변환 큐에 추가 (공통 로직)"""
        if not self.monitor:
            self.monitor = FileMonitor(self.config_manager, self.add_log)
        
        pdf_output_folder = self.config_manager.get("pdf_output_folder", "").strip()
        output_dir = pdf_output_folder if pdf_output_folder else None
        queued = 0
        skipped = 0
        
        for path in filepaths:
            if not os.path.exists(path):
                self.add_log(f"파일을 찾을 수 없습니다: {path}", "warning")
                skipped += 1
                continue
            
            if os.path.isdir(path):
                self.add_log(f"폴더는 지원하지 않습니다: {path}", "warning")
                skipped += 1
                continue
            
            ext = os.path.splitext(path)[1].lower()
            if ext not in ['.hwp', '.hwpx']:
                self.add_log(f"지원하지 않는 파일 형식: {os.path.basename(path)}", "warning")
                skipped += 1
                continue
            
            filename = os.path.basename(path)
            self.monitor.pdf_queue.add_task(path, output_dir, filename)
            queued += 1
        
        if queued:
            output_location = output_dir if output_dir else "원본 폴더"
            source_text = f"{source} " if source else ""
            self.add_log(f"{source_text}PDF 변환 작업 {queued}개가 큐에 추가되었습니다. ({output_location})", "success")
        elif skipped:
            self.add_log(f"{source}된 파일 중 변환 가능한 HWP/HWPX 파일이 없습니다.", "warning")
    
    def _on_drag_enter(self, event):
        """드래그 진입 시 드롭 영역 색상 변경"""
        self.drop_frame.configure(
            fg_color=("lightblue", "#1a3a5c"),
            border_color="royalblue"
        )

    def _on_drag_leave(self, event=None):
        """드래그 이탈 시 드롭 영역 색상 복원"""
        self.drop_frame.configure(
            fg_color=("gray90", "gray20"),
            border_color="gray50"
        )

    def clear_log(self):
        """로그 내용 지우기"""
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")

    def save_log_to_file(self):
        """로그를 파일로 저장"""
        path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("텍스트 파일", "*.txt"), ("모든 파일", "*.*")],
            title="로그 저장"
        )
        if path:
            content = self.log_textbox.get("1.0", "end")
            with open(path, "w", encoding="utf-8") as f:
                f.write(content)
            self.add_log(f"로그 저장됨: {path}", "success")

    def show_window(self, icon=None, item=None):
        """창 표시"""
        self.deiconify()
        self.lift()
        self.focus_force()
    
    def toggle_log_panel(self):
        """로그 패널 접기/펼치기"""
        if self.log_panel_visible:
            self.log_textbox.pack_forget()
            self.log_toggle_button.configure(text="펼치기")
            self.log_panel_visible = False
        else:
            self.log_textbox.pack(fill="both", expand=True, padx=10, pady=(0, 10))
            self.log_toggle_button.configure(text="접기")
            self.log_panel_visible = True
    
    def add_log(self, message: str, level: str = "info"):
        """로그 추가"""
        self.log_queue.put(message, level)
        
        # 파일 로그 저장 (설정된 경우)
        if self.config_manager.get("save_logs", False):
            try:
                log_file = self.config_manager.get("log_file_path", "monitor_log.txt")
                log_path = os.path.join(BASE_DIR, log_file)
                with open(log_path, 'a', encoding='utf-8') as f:
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    f.write(f"[{timestamp}] [{level.upper()}] {message}\n")
            except Exception as e:
                logger.error("로그 파일 저장 오류: %s", e)
    
    MAX_LOG_LINES = 1000  # 로그 최대 라인 수

    def update_logs(self):
        """로그 업데이트 (주기적 호출)"""
        try:
            logs = self.log_queue.get_all()
            for message, level, timestamp in logs:
                timestamp_str = timestamp.strftime("%H:%M:%S")
                full_message = f"[{timestamp_str}] {message}\n"
                self.log_textbox.insert("end", full_message, level)

            if logs and self.auto_scroll_var.get():
                self.log_textbox.see("end")

            # 최대 라인 수 초과 시 상단 잘라내기
            if logs:
                line_count = int(self.log_textbox.index('end-1c').split('.')[0])
                if line_count > self.MAX_LOG_LINES:
                    trim_to = line_count - self.MAX_LOG_LINES
                    self.log_textbox.delete('1.0', f'{trim_to + 1}.0')

            # 통계 인라인 업데이트
            if self.monitor:
                s = self.monitor.stats['success']
                f = self.monitor.stats['failed']
                self.stats_label.configure(text=f"✅ {s}  ❌ {f}")

            # 다음 업데이트 예약 (ID 저장)
            self._log_timer_id = self.after(100, self.update_logs)
        except Exception:
            # 앱 종료 중 위젯 파괴 등으로 발생하는 예외 무시
            pass
    
    def toggle_monitoring(self, icon=None, item=None):
        """모니터링 시작/중지"""
        if self.monitor and self.monitor.is_monitoring:
            # 중지
            self.monitor.stop_monitoring()
            self.status_label.configure(text="● 중지됨", text_color="gray")
            self.toggle_button.configure(text="시작")
            self.stats_label.configure(text="")
            self.title("파일 모니터링 및 자동 처리")
            self.add_log("모니터링이 중지되었습니다.", "info")
        else:
            # 시작
            folder_path = self.config_manager.get("monitor_folder", "")
            if not folder_path or not os.path.exists(folder_path):
                self.add_log("모니터링 폴더를 먼저 설정해주세요.", "error")
                self.open_settings()
                return

            if not self.monitor:
                self.monitor = FileMonitor(self.config_manager, self.add_log)

            if self.monitor.start_monitoring(folder_path):
                self.status_label.configure(text="● 모니터링 중", text_color="green")
                self.toggle_button.configure(text="중지")
                self.folder_label.configure(text=f"폴더: {folder_path}")
                self.title("파일 모니터링 - 모니터링 중")
                self.add_log(f"모니터링 시작: {folder_path}", "success")
            else:
                self.add_log("모니터링 시작에 실패했습니다.", "error")
    
    def process_existing_files_once(self):
        """기존 파일들을 1회 처리 (모니터링 없이)"""
        folder_path = self.config_manager.get("monitor_folder", "")
        if not folder_path or not os.path.exists(folder_path):
            self.add_log("모니터링 폴더를 먼저 설정해주세요.", "error")
            self.open_settings()
            return
        
        # 모니터 인스턴스가 없으면 생성
        if not self.monitor:
            self.monitor = FileMonitor(self.config_manager, self.add_log)
        
        # 별도 스레드에서 실행 (UI 블로킹 방지)
        def run_process():
            self.monitor.process_existing_files(folder_path)
        
        threading.Thread(target=run_process, daemon=True).start()
        self.add_log(f"기존 파일 처리 시작: {folder_path}", "info")
    
    def process_pdf_conversion_once(self):
        """모니터링 폴더의 모든 HWP/HWPX 파일을 PDF로 변환 (수동 실행)"""
        folder_path = self.config_manager.get("monitor_folder", "")
        if not folder_path or not os.path.exists(folder_path):
            self.add_log("모니터링 폴더를 먼저 설정해주세요.", "error")
            self.open_settings()
            return
        
        # 모니터 인스턴스가 없으면 생성
        if not self.monitor:
            self.monitor = FileMonitor(self.config_manager, self.add_log)
        
        # 별도 스레드에서 실행 (UI 블로킹 방지)
        def run_pdf_conversion():
            try:
                extensions = self.config_manager.get("extensions", [])
                extensions_lower = [ext.lower() for ext in extensions]
                
                # HWP/HWPX 파일 찾기
                hwp_files = []
                for filename in os.listdir(folder_path):
                    filepath = os.path.join(folder_path, filename)
                    if not os.path.isfile(filepath):
                        continue
                    
                    ext = os.path.splitext(filename)[1].lower()
                    if ext in ['.hwp', '.hwpx'] and ext in extensions_lower:
                        hwp_files.append(filepath)
                
                if not hwp_files:
                    self.add_log("PDF 변환할 HWP/HWPX 파일이 없습니다.", "info")
                    return
                
                self.add_log(f"PDF 변환 시작: {len(hwp_files)}개 파일", "info")
                
                # PDF 출력 폴더 설정 확인
                pdf_output_folder = self.config_manager.get("pdf_output_folder", "")
                if pdf_output_folder and os.path.exists(pdf_output_folder):
                    output_dir = pdf_output_folder
                else:
                    output_dir = None  # 원본 파일과 같은 폴더에 저장
                
                # 각 파일을 PDF 변환 큐에 추가
                for filepath in hwp_files:
                    filename = os.path.basename(filepath)
                    self.monitor.pdf_queue.add_task(filepath, output_dir, filename)
                
                self.add_log(f"PDF 변환 작업 {len(hwp_files)}개가 큐에 추가되었습니다.", "success")
                
            except Exception as e:
                self.add_log(f"PDF 변환 오류: {str(e)}", "error")
        
        threading.Thread(target=run_pdf_conversion, daemon=True).start()
    
    def open_settings(self, icon=None, item=None):
        """설정 창 열기"""
        settings_window = SettingsWindow(self, self.config_manager)
        settings_window.grab_set()
        self.wait_window(settings_window)
        
        # 설정 변경 후 UI 업데이트
        folder_path = self.config_manager.get("monitor_folder", "")
        if folder_path:
            self.folder_label.configure(text=f"폴더: {folder_path}")
        
        # 모니터링 중이면 재시작
        if self.monitor and self.monitor.is_monitoring:
            self.toggle_monitoring()
            self.after(500, self.toggle_monitoring)
    
    def on_closing(self):
        """창 닫기 이벤트"""
        if self.monitor:
            self.monitor.stop_monitoring()
        
        # 창 크기 저장
        geometry = self.geometry()
        self.config_manager.set("window_geometry", geometry)
        
        # 창 숨기기 (트레이에만 표시)
        self.withdraw()
    
    def quit_app(self, icon=None, item=None):
        """애플리케이션 종료"""
        # 로그 업데이트 타이머 취소
        if self._log_timer_id:
            self.after_cancel(self._log_timer_id)
            self._log_timer_id = None

        if self.monitor:
            self.monitor.stop_monitoring()

        if self.tray_icon:
            self.tray_icon.stop()

        self.quit()
        self.destroy()


class SettingsWindow(ctk.CTkToplevel):
    """설정 창"""
    
    def __init__(self, parent, config_manager: ConfigManager):
        super().__init__(parent)

        self.config_manager = config_manager
        self.title("설정")
        self.geometry("580x680")
        self.transient(parent)

        # ── 탭 뷰 ──
        tabview = ctk.CTkTabview(self)
        tabview.pack(fill="both", expand=True, padx=20, pady=(20, 0))

        tab_basic = tabview.add("기본")
        tab_convert = tabview.add("변환")
        tab_other = tabview.add("기타")

        # ══════════════════════════════
        # 탭 1: 기본
        # ══════════════════════════════
        basic_scroll = ctk.CTkScrollableFrame(tab_basic)
        basic_scroll.pack(fill="both", expand=True)

        # 모니터링 폴더
        ctk.CTkLabel(
            basic_scroll,
            text="모니터링 폴더",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        folder_input_frame = ctk.CTkFrame(basic_scroll)
        folder_input_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.folder_entry = ctk.CTkEntry(folder_input_frame)
        self.folder_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.folder_entry.insert(0, config_manager.get("monitor_folder", ""))

        ctk.CTkButton(
            folder_input_frame, text="찾기", command=self.browse_folder, width=80
        ).pack(side="right")

        # 처리할 확장자 (그룹별)
        ctk.CTkLabel(
            basic_scroll,
            text="처리할 확장자",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.extension_vars = {}
        current_extensions = config_manager.get("extensions", [])

        ext_groups = [
            ("── 한글 ──", [".hwp", ".hwpx"]),
            ("── MS Office ──", [".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx"]),
            ("── 기타 ──", [".pdf"]),
        ]
        for group_label, exts in ext_groups:
            ctk.CTkLabel(
                basic_scroll,
                text=group_label,
                font=ctk.CTkFont(size=11),
                text_color="gray"
            ).pack(anchor="w", padx=15, pady=(8, 2))
            for ext in exts:
                var = ctk.BooleanVar(value=ext in current_extensions)
                self.extension_vars[ext] = var
                ctk.CTkCheckBox(basic_scroll, text=ext, variable=var).pack(
                    anchor="w", padx=25, pady=2
                )

        # ══════════════════════════════
        # 탭 2: 변환
        # ══════════════════════════════
        convert_scroll = ctk.CTkScrollableFrame(tab_convert)
        convert_scroll.pack(fill="both", expand=True)

        # 자동 PDF 변환 스위치
        ctk.CTkLabel(
            convert_scroll,
            text="자동 PDF 변환",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.auto_convert_var = ctk.BooleanVar(
            value=config_manager.get("auto_convert_pdf", True)
        )
        ctk.CTkSwitch(
            convert_scroll,
            text="모니터링 중 자동으로 PDF 변환",
            variable=self.auto_convert_var,
            onvalue=True,
            offvalue=False
        ).pack(anchor="w", padx=20, pady=(0, 10))

        # PDF 출력 폴더
        ctk.CTkLabel(
            convert_scroll,
            text="PDF 출력 폴더",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        ctk.CTkLabel(
            convert_scroll,
            text="비워두면 원본 파일과 같은 폴더에 저장",
            font=ctk.CTkFont(size=11),
            text_color="gray"
        ).pack(anchor="w", padx=15, pady=(0, 5))

        pdf_out_input = ctk.CTkFrame(convert_scroll)
        pdf_out_input.pack(fill="x", padx=10, pady=(0, 10))

        self.pdf_output_entry = ctk.CTkEntry(pdf_out_input)
        self.pdf_output_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        pdf_output_folder = config_manager.get("pdf_output_folder", "")
        if pdf_output_folder:
            self.pdf_output_entry.insert(0, pdf_output_folder)

        ctk.CTkButton(
            pdf_out_input, text="찾기", command=self.browse_pdf_output_folder, width=80
        ).pack(side="right")

        # Hancom PDF 프린터 이름
        ctk.CTkLabel(
            convert_scroll,
            text="Hancom PDF 프린터 이름",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.printer_entry = ctk.CTkEntry(convert_scroll)
        self.printer_entry.pack(fill="x", padx=10, pady=(0, 10))
        self.printer_entry.insert(0, config_manager.get("hancom_pdf_printer", "Hancom PDF"))

        # HWPX 변환기 경로
        ctk.CTkLabel(
            convert_scroll,
            text="HWPX 변환기 경로",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        hwpx_input_frame = ctk.CTkFrame(convert_scroll)
        hwpx_input_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.hwpx_entry = ctk.CTkEntry(hwpx_input_frame)
        self.hwpx_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.hwpx_entry.insert(0, config_manager.get("hwpx_converter_path", ""))

        ctk.CTkButton(
            hwpx_input_frame, text="찾기", command=self.browse_hwpx_converter, width=80
        ).pack(side="right")

        # ══════════════════════════════
        # 탭 3: 기타
        # ══════════════════════════════
        other_scroll = ctk.CTkScrollableFrame(tab_other)
        other_scroll.pack(fill="both", expand=True)

        # 로그 파일 저장 스위치
        ctk.CTkLabel(
            other_scroll,
            text="로그 파일",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.save_logs_var = ctk.BooleanVar(value=config_manager.get("save_logs", False))
        ctk.CTkSwitch(
            other_scroll,
            text="로그 파일 저장",
            variable=self.save_logs_var,
            onvalue=True,
            offvalue=False,
            command=self._toggle_log_path_state
        ).pack(anchor="w", padx=20, pady=(0, 5))

        log_path_frame = ctk.CTkFrame(other_scroll)
        log_path_frame.pack(fill="x", padx=10, pady=(0, 10))

        ctk.CTkLabel(log_path_frame, text="로그 파일 경로:").pack(
            side="left", padx=(0, 10)
        )
        self.log_path_entry = ctk.CTkEntry(log_path_frame)
        self.log_path_entry.pack(side="left", fill="x", expand=True)
        self.log_path_entry.insert(0, config_manager.get("log_file_path", "monitor_log.txt"))

        # 테마
        ctk.CTkLabel(
            other_scroll,
            text="테마",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.theme_var = ctk.StringVar(value=config_manager.get("theme", "dark"))
        ctk.CTkRadioButton(
            other_scroll, text="다크", variable=self.theme_var, value="dark"
        ).pack(anchor="w", padx=20, pady=2)
        ctk.CTkRadioButton(
            other_scroll, text="라이트", variable=self.theme_var, value="light"
        ).pack(anchor="w", padx=20, pady=2)

        # 디버그 모드
        ctk.CTkLabel(
            other_scroll,
            text="개발자",
            font=ctk.CTkFont(size=14, weight="bold")
        ).pack(anchor="w", padx=10, pady=(10, 5))

        self.debug_mode_var = ctk.BooleanVar(
            value=config_manager.get("debug_mode", False)
        )
        ctk.CTkCheckBox(
            other_scroll, text="디버그 모드", variable=self.debug_mode_var
        ).pack(anchor="w", padx=20, pady=(0, 10))

        # 초기 로그 경로 활성화 상태 설정
        self._toggle_log_path_state()

        # ── 저장/취소 버튼 (탭 밖 하단 고정) ──
        button_frame = ctk.CTkFrame(self)
        button_frame.pack(fill="x", padx=20, pady=(5, 20))

        ctk.CTkButton(
            button_frame, text="저장", command=self.save_settings, width=100
        ).pack(side="right", padx=10)

        ctk.CTkButton(
            button_frame, text="취소", command=self.destroy, width=100
        ).pack(side="right")
    
    def _toggle_log_path_state(self):
        """로그 저장 스위치 상태에 따라 경로 입력 활성/비활성"""
        state = "normal" if self.save_logs_var.get() else "disabled"
        self.log_path_entry.configure(state=state)

    def browse_folder(self):
        """폴더 선택 다이얼로그"""
        folder = filedialog.askdirectory(title="모니터링 폴더 선택")
        if folder:
            self.folder_entry.delete(0, "end")
            self.folder_entry.insert(0, folder)
    
    def browse_pdf_output_folder(self):
        """PDF 출력 폴더 선택 다이얼로그"""
        folder = filedialog.askdirectory(title="PDF 출력 폴더 선택")
        if folder:
            self.pdf_output_entry.delete(0, "end")
            self.pdf_output_entry.insert(0, folder)

    def browse_hwpx_converter(self):
        """HWPX 변환기 실행 파일 선택 다이얼로그"""
        filepath = filedialog.askopenfilename(
            title="HWPX 변환기 선택",
            filetypes=[("실행 파일", "*.exe"), ("모든 파일", "*.*")]
        )
        if filepath:
            self.hwpx_entry.delete(0, "end")
            self.hwpx_entry.insert(0, filepath)

    def save_settings(self):
        """설정 저장"""
        # 폴더 경로 유효성 검사
        folder_path = self.folder_entry.get().strip()
        if folder_path and not os.path.exists(folder_path):
            messagebox.showerror("오류", "폴더를 찾을 수 없습니다.")
            return

        # PDF 출력 폴더 유효성 검사
        pdf_output_folder = self.pdf_output_entry.get().strip()
        if pdf_output_folder and not os.path.exists(pdf_output_folder):
            messagebox.showerror("오류", "PDF 출력 폴더를 찾을 수 없습니다.")
            return

        # 모든 설정을 한 번에 저장 (파일 I/O 1회)
        theme = self.theme_var.get()
        self.config_manager.batch_update({
            "monitor_folder": folder_path,
            "extensions": [ext for ext, var in self.extension_vars.items() if var.get()],
            "pdf_output_folder": pdf_output_folder,
            "hancom_pdf_printer": self.printer_entry.get().strip() or "Hancom PDF",
            "hwpx_converter_path": self.hwpx_entry.get().strip(),
            "save_logs": self.save_logs_var.get(),
            "log_file_path": self.log_path_entry.get().strip(),
            "theme": theme,
            "auto_convert_pdf": self.auto_convert_var.get(),
            "debug_mode": self.debug_mode_var.get(),
        })

        ctk.set_appearance_mode(theme)
        messagebox.showinfo("저장 완료", "설정이 저장되었습니다.")
        self.destroy()


def main():
    """메인 함수"""
    app = MonitorApp()
    app.protocol("WM_DELETE_WINDOW", app.on_closing)
    app.mainloop()


if __name__ == "__main__":
    main()

