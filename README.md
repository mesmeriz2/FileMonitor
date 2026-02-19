# FileMonitor

Windows용 **파일 모니터링 및 한글(HWP/HWPX) 자동화 도구**입니다. 지정 폴더를 감시하면서 새로 추가된 파일에 **날짜 헤더(YYMMDD)** 를 붙이고, HWP 파일을 HWPX로 변환하며, **HWP/HWPX를 PDF로 변환**할 수 있습니다.

---

## 주요 기능

- **폴더 모니터링**  
  지정한 폴더에 새 파일이 추가되면 자동으로 처리합니다.
- **날짜 헤더 추가**  
  파일명 앞에 `YYMMDD` 형식의 날짜 접두사를 붙입니다. 이미 다른 형식(YYYYMMDD, YYYY.MM.DD 등)으로 된 날짜가 있으면 6자리 형식으로 통일합니다.
- **HWP → HWPX 변환**  
  모니터링 중인 HWP 파일을 HWPX로 변환할 수 있습니다. (한컴 HwpxConverter 사용)
- **HWP/HWPX → PDF 변환**  
  한컴오피스(pyhwpx)를 이용해 HWP/HWPX 파일을 PDF로 변환합니다.  
  - 드래그 앤 드롭 또는 **파일 선택** 버튼으로 변환할 파일을 지정  
  - **1회 실행**으로 모니터링 폴더 내 모든 한글 파일 일괄 PDF 변환  
  - PDF 출력 폴더를 별도 지정 가능 (비워두면 원본과 같은 폴더에 저장)
- **시스템 트레이**  
  창을 최소화하면 트레이 아이콘으로 줄어들어 백그라운드에서 실행할 수 있습니다.
- **다크/라이트 테마**  
  설정에서 테마를 선택할 수 있습니다.

---

## 요구 사항

- **Windows** (한컴오피스 COM 연동 사용)
- **Python 3.10+** (권장: 3.12 이하 — tkinterdnd2 호환성)
- **한컴오피스** 설치 (PDF 변환용, Hancom PDF 프린터 포함)
- **HwpxConverter** (HWP→HWPX 변환 시):  
  기본 경로 `C:\Program Files (x86)\Hnc\HwpxConverter\HwpxConverter.exe`  
  다른 경로는 설정에서 지정 가능

---

## 설치

```bash
git clone https://github.com/mesmeriz2/FileMonitor.git
cd FileMonitor
pip install -r requirements.txt
```

### 의존성 요약

| 패키지 | 용도 |
|--------|------|
| customtkinter | GUI |
| tkinterdnd2 | 드래그 앤 드롭 (선택) |
| watchdog | 폴더 모니터링 |
| pystray | 시스템 트레이 |
| Pillow | 트레이 아이콘 등 이미지 |
| pyhwpx | 한글 COM 연동, PDF 변환 |
| pywin32 | Windows COM |

---

## 사용 방법

1. **실행**  
   `python file_monitor.py`
2. **설정**  
   ⚙️ **설정** 버튼에서 다음을 지정합니다.  
   - **모니터링 폴더**: 감시할 폴더 경로  
   - **처리할 확장자**: `.hwp`, `.hwpx`, `.doc`, `.docx`, `.xls`, `.xlsx`, `.ppt`, `.pptx`, `.pdf` 중 선택  
   - **HWPX 변환기 경로**: HwpxConverter 실행 파일 경로  
   - **PDF 출력 폴더**: 비워두면 원본과 같은 폴더에 PDF 저장  
   - **테마**: 다크 / 라이트  
   - **로그 저장**: 로그 파일 저장 여부 및 경로
3. **모니터링 시작**  
   **시작** 버튼을 누르면 해당 폴더 감시가 시작됩니다.  
   - 새 파일이 들어오면 날짜 헤더 추가 및(해당 시) HWP→HWPX 변환이 자동 수행됩니다.
4. **PDF 변환**  
   - **PDF 변환** 버튼 → 모니터링 폴더 내 HWP/HWPX를 PDF로 변환  
   - **파일 선택** 버튼 또는 드롭으로 특정 파일만 선택해 PDF 변환  
   - 변환 작업은 큐로 순차 처리됩니다.
5. **1회 실행**  
   **1회 실행** 버튼으로 현재 폴더에 이미 있는 파일만 한 번만 처리(날짜 헤더 + HWPX 변환)할 수 있습니다.

설정은 `config.json`에 저장됩니다. (저장소에는 포함되지 않음)

---

## 설정 (config.json)

| 키 | 설명 |
|----|------|
| `monitor_folder` | 모니터링할 폴더 경로 |
| `extensions` | 처리할 확장자 목록 |
| `pdf_output_folder` | PDF 저장 폴더 (빈 문자열이면 원본과 동일 폴더) |
| `hwpx_converter_path` | HwpxConverter 실행 파일 경로 |
| `save_logs` | 로그 파일 저장 여부 |
| `log_file_path` | 로그 파일 경로 |
| `theme` | `"dark"` / `"light"` |
| `window_geometry` | 창 크기 등 |

---

## exe 빌드 (PyInstaller)

Windows에서 실행 파일로 패키징할 때:

```powershell
.\build_exe.ps1
```

또는:

```powershell
python -m PyInstaller --noconfirm file_monitor.spec
```

빌드 결과는 `dist\FileMonitor.exe`에 생성됩니다.

---

## HWP PDF 변환 시 보안 경고 없애기

한컴 오토메이션으로 파일을 열 때 나오는 보안 경고를 없애려면:

1. 한컴 개발자 사이트에서 **오토메이션용 보안 모듈** DLL(`SecurityModule.dll` 등)을 받아 로컬에 저장합니다.
2. 레지스트리에 등록  
   - 경로: `HKEY_CURRENT_USER\Software\HNC\HwpAutomation\Modules`  
   - 이름: `SecurityModule`  
   - 값: DLL 전체 경로 (따옴표 없이)
3. (선택) `HKEY_CURRENT_USER\Software\HNC\HwpAutomation`에 `ExtOpenWarning` = `0` 설정  
4. 코드에서 한글 인스턴스 생성 직후, `Open()` 호출 전에  
   `hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")`  
   를 호출하도록 수정하면 됩니다. (자세한 절차는 한컴 오토메이션 문서 참고)

---

## 라이선스

이 프로젝트의 라이선스는 저장소에 별도 명시가 없는 경우 기본 규칙을 따릅니다.
