@echo off
chcp 65001 >nul
cd /d "%~dp0"
python "%~dp0kakao_notify.py" setup
