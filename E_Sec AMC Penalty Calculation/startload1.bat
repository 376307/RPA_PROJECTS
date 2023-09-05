@echo off
sqlldr 'KPMG/Asd$1234@HISTDB' control='Control1.txt' log='Results1.log' direct='true'
pause
