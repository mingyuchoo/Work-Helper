# 업무 도우미 (데스크톱 앱)

세 가지 반복 업무(**엑셀 취합 · 주소록 기반 메일 수신자 생성 · 폴더 생성**)를
화면에서 클릭으로 실행하는 데스크톱 앱입니다.

동작 방식: **pywebview** 창 안에 화면(HTML)을 띄우고, 버튼을 누르면
파이썬 백엔드(`backend/*.py`)가 내 PC에서 실제 작업을 합니다.
모든 처리는 로컬에서만 이루어지며, 파일이 외부로 나가지 않습니다.

## 폴더 구조
```
(프로젝트 폴더)/
├─ app.py                  ← 진입점 (pywebview 창 + JS↔파이썬 연결)
├─ backend/
│  ├─ mail_recipients.py   ← 주소록 → 수신/참조 이메일
│  ├─ xls_merge.py         ← 엑셀 시트 취합
│  └─ folders.py           ← 폴더 생성
├─ web/
│  ├─ index.html           ← 화면 + 동작 로직
│  └─ fonts/               ← (선택) 폰트
├─ 빌드.bat                 ← 수정 후 exe 다시 만들기
└─ dist/WorkHelper/         ← 배포용 실행 폴더
```

## 실행 방법

### 1) 개발/테스트 (파이썬으로 바로 실행)
```
python app.py
```
> 처음 실행 시 `pip install pywebview openpyxl` 필요.
> Windows 10/11 에는 WebView2 런타임이 기본 설치돼 있어 별도 설치가 거의 필요 없습니다.

### 2) 배포용 빌드 만들기
`빌드.bat` 더블클릭 → `dist\WorkHelper\` 폴더 생성 (onedir 방식).
```
dist\WorkHelper\
├─ WorkHelper.exe                 ← 실행 파일 (이걸 더블클릭)
├─ MicrosoftEdgeWebview2Setup.exe ← WebView2 설치파일 (동봉)
└─ _internal\                     ← 필수 구성요소 (같이 있어야 실행됨)
```
- **폴더 전체**를 zip 으로 묶어 배포하세요. (예: `dist\WorkHelper_배포.zip`)
- 받는 사람은 **압축을 푼 뒤** 폴더 안의 `WorkHelper.exe` 를 실행합니다.
- ⚠️ exe 만 따로 빼내면 실행되지 않습니다(`_internal` 폴더가 옆에 있어야 함).

> **왜 단일 exe(onefile)가 아닌가?** onefile 은 매 실행마다 파일을 임시폴더로 압축 해제하고
> 페이지를 불러오는 데 오래 걸려, 그동안 창이 멈춰 "응답없음" 으로 보였습니다.
> onedir(폴더) 방식은 압축 해제가 없어 **몇 초 만에** 바로 뜹니다.

## ⚠️ WebView2 런타임
화면을 그릴 때 **Microsoft Edge WebView2 런타임**을 사용합니다.
- Windows 11 / 최신 Windows 10: 대부분 기본 설치되어 그대로 실행됩니다.
- 없는 PC: 앱이 시작할 때 감지해 **"지금 설치할까요?"** 안내를 띄우고,
  [예] 를 누르면 동봉된 `MicrosoftEdgeWebview2Setup.exe` 를 실행합니다. 설치 후 다시 실행하면 정상 동작합니다.

## 기능 사용법
- **주소록**: 주소록 엑셀 선택 → 대상 팀 입력 → 수신/참조 역할 선택 → 실행.
  - 대상 팀 입력: `전체`(모든 팀) · `사업부명/담당명`(그 조직의 모든 팀) · `팀명`(해당 팀). 쉼표로 혼합 가능.
- **엑셀 취합**: 폴더 선택 → 접두어·시트명·헤더 행번호 입력 → 실행 → `[취합]…xlsx` 생성.
- **폴더 생성**: 엑셀 선택 → 시트·상위/하위 컬럼·번호방식 선택(미리보기) → 생성 위치 선택 → 실행.

## 유지보수 메모
- 화면(디자인/문구): `web/index.html`. 스타일 상수는 파일 상단 `C`, `INPUT_BASE` 등.
- 처리 로직: `backend/` 3개 파일. `input()`/파일선택창/`print` 만 제거하고 함수화한 것이라 로직은 동일.
- 파이썬↔화면 연결: JS에서 `window.pywebview.api.<메서드>()` → `app.py` 의 `Api` 클래스 메서드.
  각 메서드는 `{"ok": true, ...}` 또는 `{"ok": false, "error": "..."}` 형태의 dict 반환.
- 진단 로그: 실행 시마다 `%TEMP%\WorkHelper.log`. (개인정보 보호를 위해 API 인자 '값'은 기록하지 않고 개수만 남김)
  개발자도구가 필요하면 `set APP_DEBUG=1` 후 실행 → 창에서 우클릭→검사.
