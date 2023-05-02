@echo off
sqlldr 'KPMG/Asd$1234@HISTDB' control='Control.txt' log='Results.log' direct='true'
pause


