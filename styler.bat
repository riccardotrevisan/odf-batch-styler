@echo off
rem Drop .odt files onto this file, or run: styler.bat [--dry-run] file.odt ...
wsl.exe -e bash /home/trevi/odf-batch-styler/styler-wsl.sh %*
if not defined STYLER_NO_PAUSE pause
