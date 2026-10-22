@echo off
chcp 65001 >nul
REM 매일 09:00 에 update_btc.py 를 실행하는 작업 스케줄러 등록 (한 번만 실행)
cd /d "%~dp0"

where pythonw >nul 2>nul
if errorlevel 1 (
  echo [오류] Python 이 설치되어 있지 않습니다. https://www.python.org 에서 설치하세요 ^(Add to PATH 체크^).
  pause & exit /b 1
)
for /f "delims=" %%i in ('where pythonw') do if not defined PYW set "PYW=%%i"

echo openpyxl 설치 중...
python -m pip install --quiet --upgrade openpyxl

echo 작업 스케줄러 등록 중...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$a = New-ScheduledTaskAction -Execute '%PYW%' -Argument '\"%~dp0update_btc.py\"' -WorkingDirectory '%~dp0';" ^
  "$t = New-ScheduledTaskTrigger -Daily -At 09:00;" ^
  "$s = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;" ^
  "Register-ScheduledTask -TaskName 'BTC_PowerLaw_Update' -Action $a -Trigger $t -Settings $s -Force | Out-Null"
if errorlevel 1 ( echo [오류] 등록 실패 & pause & exit /b 1 )

echo 지금 한 번 실행해서 빠진 기간을 채웁니다...
python "%~dp0update_btc.py"
type "%~dp0update_btc.log"
echo.
echo 완료: 매일 오전 9시에 자동 업데이트됩니다. (작업 이름: BTC_PowerLaw_Update)
pause
