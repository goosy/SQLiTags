set "TS=%date:~0,4%%date:~5,2%%date:~8,2%%time:~0,2%%time:~3,2%%time:~6,2%"
set "TS=%TS: =0%"
C:\WINDOWS\SysWOW64\CScript.exe sqlitags.wsh.js restart /date:%TS%