@echo off
chcp 65001 >nul
REM ── 업무 도우미: 코드 수정 후 exe 다시 만들기 ──
REM 이 파일이 있는 폴더에서 실행하세요. (더블클릭 또는 터미널)

cd /d "%~dp0"

echo [1/3] 필요 패키지 확인/설치...
python -m pip install --quiet pywebview openpyxl pyinstaller

echo [2/3] exe 빌드 시작... (몇 분 걸립니다)
REM --onedir: 파일을 폴더에 그대로 두어 실행 속도가 빠르고 안정적입니다.
REM (--onefile 은 매 실행마다 임시폴더로 압축해제 → 로딩이 느려 '응답없음'처럼 보일 수 있어 권장하지 않음)
python -m PyInstaller --onedir --windowed --name "WorkHelper" ^
  --add-data "web;web" ^
  --collect-all webview ^
  --collect-submodules openpyxl ^
  --noconfirm ^
  app.py

echo [3/3] WebView2 설치파일을 배포 폴더에 복사...
if exist "dist\MicrosoftEdgeWebview2Setup.exe" copy /Y "dist\MicrosoftEdgeWebview2Setup.exe" "dist\WorkHelper\" >nul

echo.
echo ============================================================
echo  완료!  배포 폴더: dist\WorkHelper\
echo   - WorkHelper.exe  (실행 파일)
echo   - MicrosoftEdgeWebview2Setup.exe (WebView2 설치파일)
echo   - _internal\  (필수 구성요소 · 함께 있어야 실행됨)
echo  이 폴더 전체를 zip 으로 묶어 배포하세요.
echo ============================================================
echo  (개발 중 오류를 콘솔에서 보려면: set APP_DEBUG=1 ^&^& python app.py)
echo.
pause
