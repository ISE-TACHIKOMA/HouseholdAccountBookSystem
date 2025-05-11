setlocal enabledelayedexpansion
cd %~dp0
if exist クレカ支払い割合.exe del /f クレカ支払い割合.exe


pyinstaller -wF main.py --onefile --noconsole --name クレカ支払い割合 

move dist\クレカ支払い割合.exe クレカ支払い割合.exe

rd /s /q dist
rd /s /q build
del /s *.spec

paused