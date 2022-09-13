Option Explicit

const DBF = "D:/config/wincc.db"
const CONF = "D:/config/config.ini"
Dim conn_str : conn_str = Replace("Driver={SQLite3 ODBC Driver};Database=@FSPEC@;StepAPI=;Timeout=", "@FSPEC@", DBF )
Dim conn : Set conn = CreateObject( "ADODB.Connection" )
Dim timestamp, timezone

' get timezone on startup
Sub getTimeZone()
  Dim objWMIService, colItems, objItem
  Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
  Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem",,48)
  For Each objItem in colItems
    TimeZone = objItem.CurrentTimeZone * 60
  Next
End Sub
getTimeZone

' set timestamp
Sub setDT(dt_str)
  ' Year(dt), Month(dt), Day(dt), Hour(dt), Minute(dt)
  Dim dt
  If LCase(dt_str) = "now" Then
    dt = Now
  Else
    dt = CDate(dt_str)
  End If
  timestamp = DateDiff( _
    "s", _
    "1970/01/01 00:00:00", _
    FormatDateTime(dt, 2) & " " & FormatDateTime(dt, 4) & ":0" _
  ) - timezone
End Sub

Dim tags_1N, tags_10N, tags_30N, tags_1H, tags_12H, tags_1D, tags_1M
Set tags_1N = CreateObject("System.Collections.ArrayList")
Set tags_10N = CreateObject("System.Collections.ArrayList")
Set tags_30N = CreateObject("System.Collections.ArrayList")
Set tags_1H = CreateObject("System.Collections.ArrayList")
Set tags_12H = CreateObject("System.Collections.ArrayList")
Set tags_1D = CreateObject("System.Collections.ArrayList")
Set tags_1M = CreateObject("System.Collections.ArrayList")

' parse line includes '=' in ini file
Function parse_line(lineText)
  Dim tag, pair, tagdesc
  pair = Split(lineText, "=", 2)
  If ubound(pair) = 1 Then
    tagdesc = Split(Trim(pair(1)), ",", 3)
    If ubound(tagdesc) = 2 Then
      Set tag = CreateObject("Scripting.Dictionary")
      tag.Add "name", Trim(pair(0))
      tag.Add "tagname", tagdesc(0)
      tag.Add "valid", tagdesc(1)
      select case tagdesc(2)
        case "1minute"
          tags_1N.add tag
        case "10minute"
          tags_10N.add tag
        case "30minute"
          tags_30N.add tag
        case "12hours"
          tags_12H.add tag
        case "1day"
          tags_1D.add tag
        case "1month"
          tags_1M.add tag
        case else
          tags_1H.add tag
      end select
    End If
  End If
End Function

' initialize from configuration
Sub init
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  Dim sSQL : sSQL = "CREATE TABLE tags (time TIMESTAMP NOT NULL, name TEXT NOT NULL, value REAL, PRIMARY KEY (time, name))"

  If Not fso.FileExists(DBF) Then 'if possiable create file
    fso.CreateTextFile DBF, True
    On Error Resume Next
    conn.Open conn_str
    If Err <> 0 Then
      HMIRuntime.Trace "can't open DB: " & Err.Description
    Else
      HMIRuntime.Trace "connected to " & DBF
      conn.Execute sSQL
      If Err = 0 Then
        HMIRuntime.Trace "table tags created"
      End If
    End If
    conn.Close
    On Error GoTo 0
  End If

  If Not fso.FileExists(CONF) Then
    HMIRuntime.Trace  "configure file " & CONF & " not exists!" & vbCrLf
    Return
  End If
  Const ForReading = 1
  Dim INIFile : Set INIFile = fso.OpenTextFile(CONF, ForReading)

  Do While Not INIFile.AtEndOfStream
    Dim section, comment, strLine
    strLine = Trim(INIFile.Readline)
    If strLine <> "" Then
      If Left(strLine, 1) = ";" OR Left(strLine, 1) = "#" Then
        comment = Trim(Mid(strLine, 2, Len(strLine) - 1))
      ElseIf Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
        section = Trim(Mid(strLine, 2, Len(strLine) - 2))
      ElseIf section = "tags" Then
        parse_line(strLine)
      End If
    End If
  Loop

  INIFile.Close
  Set INIFile = Nothing
End Sub

' archive tags to sqlite
Sub saveTags(tags)
  Dim tag, valid, tagvalue, sSQL
  On Error Resume Next
  conn.Open conn_str
  If Err <> 0 Then
    HMIRuntime.Trace "can't open DB: " & Err.Description
  Else
    For Each tag In tags
      tagvalue = HMIRuntime.Tags(tag("tagname")).Read
      If HMIRuntime.Tags(tag("valid")).Read = 1 Then
        sSQL = "INSERT OR IGNORE INTO tags VALUES (" & timestamp & ", '" & tag("name") & "', " & tagvalue & ");"
        conn.Execute sSQL
        If Err <> 0 Then
          HMIRuntime.Trace "error on execute: " & sSQL & vbCrLf
        Else
          HMIRuntime.Trace "save: " & timestamp & ", '" & tag("name") & "', " & tagvalue & vbCrLf
        End If
      End If
    Next
  End If
  conn.Close
  On Error GoTo 0
End Sub

' get historical data
Function getHisTag(tagname, datastr)
  Dim timestamp : timestamp = DateDiff("s", "1970/01/01 00:00:00", datastr) - timezone
  Dim sSQL : sSQL = "SELECT value FROM tags WHERE " & _
    "name = '" &  tagname & _
    "' AND time = '" & timestamp & "';"
  On Error Resume Next
  conn.Open conn_str
  If Err <> 0 Then
    HMIRuntime.Trace "can't open DB: " & Err.Description & vbCrLf
  Else
    Dim rs : Set rs = conn.Execute( sSQL )
    If Err <> 0 Then
      HMIRuntime.Trace "error on execute: " & sSQL & vbCrLf
    Else
      If NOT rs.EOF Then
        Dim f : For Each f In rs.Fields
          getHisTag = f.value
        Next
      Else
        getHisTag = -100000.0
      End If
      HMIRuntime.Trace "read: " & timestamp & " '" & tagname & "' " & getHisTag & vbCrLf
    End If
  End If
  conn.Close
  On Error GoTo 0
End Function

' archive function for WinCC actions
Sub save1N()
  setDT("now")
  saveTags(tags_1N)
End Sub
Sub save10N()
  setDT("now")
  saveTags(tags_10N)
End Sub
Sub save30N()
  setDT("now")
  saveTags(tags_30N)
End Sub
Sub save1H()
  setDT("now")
  saveTags(tags_1H)
End Sub
Sub save12H()
  setDT("now")
  saveTags(tags_12H)
End Sub
Sub save1D()
  setDT("now")
  saveTags(tags_1D)
End Sub
Sub save1M()
  setDT("now")
  saveTags(tags_1M)
End Sub

' init

' ======================模拟WinCC环境
Class Runtime
  public pTags
  Private Sub Class_Initialize
    Set pTags = CreateObject("Scripting.Dictionary")
  End Sub
  Function Tags(str)
    If pTags.Exists(str) Then
      set Tags = pTags(str)
    Else
      Dim HMITag
      set HMITag = new HMITags
      HMITag.rawStr = str
      pTags.Add str, HMITag
      Set Tags = HMITag
    End If
  End Function
  Sub Trace (msg)
    WScript.Echo msg
  End Sub
End Class

Class HMITags
    Public Value
    Public rawStr
    Private Sub Class_Initialize
        rawStr = ""
    End Sub
    Public Function Read()
        Read = Value
    End Function
    Public Function Write(v)
        Value = v
    End Function
End Class

Dim HMIRuntime : set HMIRuntime = new Runtime

' 预置几个WinCC变量
Dim kv
For Each kv in Array( _
  Array("template_folder", Left(wscript.scriptfullname,InStrRev(wscript.scriptfullname,"\")) & "template\"), _
  Array("GR_S7/AIT0201.WIO", 0.3), _
  Array("GR_S7/AIT0201.work_F", 1), _
  Array("GR_S7/Flow33.work_F", 1), _
  Array("GR_S7/Flow34.work_F", 1), _
  Array("GR_S7/Flow33.mass", 19588.7), _
  Array("GR_S7/Flow33.density", 98.3), _
  Array("GR_S7/Flow33.temperature", 38.2), _
  Array("GR_S7/Flow33.volume", 26555.2), _
  Array("GR_S7/Flow33.volume_flow_rate", 21555.2), _
  Array("GR_S7/Flow33.mass_flow_rate", 38.2), _
  Array("GR_S7/Flow34.mass", 69553.2), _
  Array("GR_S7/Flow34.density", 98.2), _
  Array("GR_S7/Flow34.temperature", 38.2), _
  Array("GR_S7/Flow34.volume", 35668.2), _
  Array("GR_S7/Flow34.volume_flow_rate", 33668.2), _
  Array("GR_S7/Flow34.mass_flow_rate", 38.2) _
)
  Dim HMITag
  set HMITag = new HMITags
  HMITag.Value = kv(1)
  HMIRuntime.pTags.Add kv(0), HMITag
Next

init
save1N()
save10N()
save30N()
save1H()
save12H()
save1D()
save1M()
getHisTag "MF33_M", "2022-9-10 23:45"