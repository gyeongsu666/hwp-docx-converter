@echo off
chcp 65001 >nul
cd /d "%~dp0"
title hwp2docx 설치

call :findpy
if not defined PYEXE goto :nopython

echo 파이썬: %PYEXE%
echo.
echo 필요한 패키지를 설치합니다...
echo.
"%PYEXE%" -m pip install --upgrade olefile python-docx
if errorlevel 1 goto :pipfail

echo.
echo ============================================
echo  설치가 끝났습니다.
echo  이제 한글 파일이나 폴더를
echo  "2_한글파일을_여기에_끌어다놓기.bat" 위로
echo  끌어다 놓으면 Word 파일이 만들어집니다.
echo ============================================
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
echo 파이썬을 찾지 못했습니다.
echo 시작 메뉴에서 "Anaconda Prompt"를 열고, 아래 한 줄을 붙여 넣어 주세요.
echo.
echo    pip install olefile python-docx
echo.
pause
exit /b 1

:pipfail
echo.
echo 패키지 설치에 실패했습니다. 인터넷 연결을 확인하거나,
echo Anaconda Prompt 에서 직접 실행해 보세요:  pip install olefile python-docx
echo.
pause
exit /b 1
