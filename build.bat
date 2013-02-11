@echo off
set PYINSTALLER=c:\Users\panurg\Downloads\pyinstaller-2.0\pyinstaller.py
set NAME=wiz
set UPXDIR=c:\Users\panurg\Downloads\upx308w\
set FLAGS=-F -n %NAME% --upx-dir=%UPXDIR%
set FILES=wiz.py
set SPEC=wiz.spec

%PYINSTALLER% %FLAGS% %FILES%
rem %PYINSTALLER% %SPEC%

