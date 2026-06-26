@echo off
chcp 65001 >nul 2>&1
title 🍅 番茄钟
cls

echo.
echo.
echo           ████████████████████████████████
echo           ██                          ██
echo           ██    🍅  番  茄  钟  🍅     ██
echo           ██     Pomodoro Timer       ██
echo           ██                          ██
echo           ████████████████████████████████
echo.
echo           专注 25 分钟 · 休息 5 分钟
echo.
echo           ⏳ 正在启动中...
echo.
echo.
start "" "%~dp0pomodoro-app\dist\win-unpacked\番茄钟.exe"
exit
