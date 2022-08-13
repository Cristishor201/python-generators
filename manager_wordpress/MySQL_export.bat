@echo off
call settings.bat

:preparation
:: 25.12.2022
For /f "tokens=1-3 delims=/. " %%a in ('date /t') do (set mydate=%%b.%%a.%%c)
:: 20-01
For /f "tokens=1-2 delims=/:" %%a in ('time /t') do (set mytime=%%a-%%b)

:open_MAMP
%mamp%/MAMP.exe

:export
%mamp_mysql_bin%/mysqldump.exe -u%name% -p%password% %database% > "%project_root%\%database%.sql"