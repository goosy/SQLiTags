dim cmd, fore
cmd = WScript.Arguments(0)
if cmd <> "stop" and cmd <> "restart" and cmd <> "list" then
  cmd = "start"
end if
fore = 1
if WScript.Arguments.count > 1 then
  if WScript.Arguments(1) = "/b" then
    fore = 0
  end if
end if

dim objWMIService, colItems, objItem, prev_procs, exists_prev_proc
Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
Set colItems = objWMIService.ExecQuery( _
  "SELECT * FROM Win32_Process WHERE Name = 'cscript.exe' OR Name = 'wscript.exe'", _
  , _
  48 _
)
set prev_procs = CreateObject("System.Collections.ArrayList")
exists_prev_proc = false
For Each objItem in colItems 
  if InStr(objItem.CommandLine, "sqlitags") > 0 then
    exists_prev_proc = true
    if cmd = "restart" or cmd = "stop" then 
      WScript.CreateObject("WScript.Shell").Run "taskkill /F /PID " & objItem.ProcessId, 0, true
    else
      prev_procs.add objItem.ProcessId
    end if
  end if
Next

if cmd = "restart" or (cmd = "start" and not exists_prev_proc) then
  WScript.CreateObject("WScript.Shell").Run "C:\\WINDOWS\\SysWOW64\\CScript.exe sqlitags.wsh.js", fore
end if
if cmd = "list" then
  WScript.Echo "已有" & prev_procs.count & "个进程: " & join(prev_procs.ToArray(), ", ")
end if
