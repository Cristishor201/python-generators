@echo off
call settings.bat

:open_MAMP
%mamp%/MAMP.exe

:create_table
%mamp_mysql_bin%/mysql.exe -u%name% -p%password% -e "CREATE DATABASE IF NOT EXISTS %database%2"

:import
%mamp_mysql_bin%/mysql.exe -u%name% -p%password% %database% < "%project_root%/%database%.sql"