@rem --------------------------------------------------------------------------------
@rem cashbook�̋N���o�b�`
@rem --------------------------------------------------------------------------------

@echo off
setlocal EnableDelayedExpansion
cd /d %~dp0

rem �����������i (�Ǘ��҃��[�h���K�v�Ȃ�)
rem openfiles > nul
rem if errorlevel 1 PowerShell.exe -Command Start-Process \"%~f0\" -Verb runas & exit

rem �Ώێw��
set "targetNm=04_�����o�[��.xlsx"
set "workNm=04_�����o�[��_WORK.xlsx"
set "sheetNm=R06"
set "startRow=5"
set "startEnd=160"


rem Python�X�N���v�g��
set "scriptNm="
set "scriptNm=%~n0.py" & rem bat��py�������Ȃ�R��

rem copy /Y !targetNm! !workNm!

rem Python�X�N���v�g�̎��s
rem python "!scriptNm!"
rem python "!scriptNm!" !workNm! !sheetNm! !startRow! !startEnd!
python "!scriptNm!" !targetNm! !sheetNm! !startRow! !startEnd!

rem del !workNm!

echo errorlevel=!errorlevel!

pause:
