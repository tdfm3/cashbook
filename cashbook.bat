@rem --------------------------------------------------------------------------------
@rem cashbookの起動バッチ
@rem --------------------------------------------------------------------------------

@echo off
setlocal EnableDelayedExpansion
cd /d %~dp0

rem 自動権限昇格 (管理者モードが必要なら)
rem openfiles > nul
rem if errorlevel 1 PowerShell.exe -Command Start-Process \"%~f0\" -Verb runas & exit

rem 対象指定
set "targetNm=04_現金出納帳.xlsx"
set "workNm=04_現金出納帳_WORK.xlsx"
set "sheetNm=R06"
set "startRow=5"
set "startEnd=160"


rem Pythonスクリプト名
set "scriptNm="
set "scriptNm=%~n0.py" & rem batとpyが同名ならコレ

rem copy /Y !targetNm! !workNm!

rem Pythonスクリプトの実行
rem python "!scriptNm!"
rem python "!scriptNm!" !workNm! !sheetNm! !startRow! !startEnd!
python "!scriptNm!" !targetNm! !sheetNm! !startRow! !startEnd!

rem del !workNm!

echo errorlevel=!errorlevel!

pause:
