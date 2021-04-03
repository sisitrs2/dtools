@echo off
python "C:\Program Files (x86)\bin\s.py" %1 > buff.txt
set /p SPATH=<buff.txt
DEL buff.txt
cd %SPATH%
