@echo off
REM Language: bat
REM Path: excel-to.pdf.cmd
set DIR=%~dp0
python "%DIR%\main.py" %*
