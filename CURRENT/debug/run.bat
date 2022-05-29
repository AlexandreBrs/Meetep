@ECHO off
COLOR 1A
TITLE Easy Meeting Planner
SET APPLI_DIR=%~dp0
REM SET PYTHON_EXE=C:\Appl\Python\Anaconda2_py27\python.exe
SET PYTHON_EXE=%APPLI_DIR%\Anaconda2_py27\python.exe
%PYTHON_EXE% %APPLI_DIR%\Meetep.py
PAUSE
EXIT