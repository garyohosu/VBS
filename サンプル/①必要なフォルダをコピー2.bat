rem 必要なフォルダのみコピー 2016/05/06 hantani

set kisyu=RS-387-9001\RS-387-9001

set stage=102
call :copyData

set stage=201
call :copyData

set stage=202
call :copyData

set stage=301
call :copyData

set stage=302
call :copyData

set stage=311
call :copyData

set stage=312
call :copyData

set stage=401
call :copyData

set stage=SER
call :copyData

echo 終了
pause
exit


:copyData
xcopy %kisyu%\%stage% data\%stage%\
exit /b

