@echo off
chcp 65001 >nul
cd /d "%~dp0"
title Word 파일을 한글로 변환

if "%~1"=="" goto :usage

call :findpy
if not defined PYEXE goto :nopython

echo.
"%PYEXE%" -m docx2hwpx.cli %*
echo.
echo 변환이 끝났습니다. 한글(.hwpx) 파일은 원본과 같은 폴더에 있습니다.
echo.
pause
exit /b 0

:usage
echo.
echo  Word 파일(.docx)이나 폴더를 이 파일 아이콘 위로
echo  끌어다 놓으면 한글 파일(.hwpx)로 바꿔 드립니다.
echo.
echo  여러 개를 한꺼번에 끌어다 놓아도 됩니다.
echo.
pause
exit /b 0

:findpy
set PYEXE=
for %%P in (python.exe) do if not defined PYEXE if exist "%%~$PATH:P" set PYEXE=%%~$PATH:P
if not defined PYEXE if exist "%USERPROFILE%\anaconda3\python.exe" set PYEXE=%USERPROFILE%\anaconda3\python.exe
if not defined PYEXE if exist "%USERPROFILE%\miniconda3\python.exe" set PYEXE=%USERPROFILE%\miniconda3\python.exe
if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe
if not defined PYEXE if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set PYEXE=%LOCALAPPDATA%\Programs\Python\Python311\python.exe
exit /b 0

:nopython
echo.
echo 파이썬을 찾지 못했습니다. 먼저 "1_처음에_한번_설치.bat" 을 실행해 주세요.
echo.
pause
exit /b 1
