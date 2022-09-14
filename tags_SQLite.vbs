Option Explicit
Const DBF = "D:/config/wincc.db"
Const CONF = "D:/config/config.ini"

Dim conn_str : conn_str = Replace("Driver={SQLite3 ODBC Driver};Database=@FSPEC@;StepAPI=;Timeout=", "@FSPEC@", DBF )
Dim conn : Set conn = CreateObject( "ADODB.Connection" )
Dim timestamp, timezone, Y, M, D, H, N, W

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
  Y = Year(dt)
  M = Month(dt)
  D = Day(dt)
  H = Hour(dt)
  N = Minute(dt)
  W = Weekday(dt, vbMonday)
  timestamp = DateDiff( _
    "s", _
    "1970-01-01 00:00:00", _
    Y & "-" & M & "-" & D & " " & H & ":" & N & ":0" _
  ) - timezone
End Sub

Dim tags_1N, tags_10N, tags_30N, tags_1H, tags_2HO, tags_2HE, tags_12H, tags_1D, tags_1M, WinCCTags
Set tags_1N = CreateObject("System.Collections.ArrayList")
Set tags_10N = CreateObject("System.Collections.ArrayList")
Set tags_30N = CreateObject("System.Collections.ArrayList")
Set tags_1H = CreateObject("System.Collections.ArrayList")
Set tags_2HO = CreateObject("System.Collections.ArrayList")
Set tags_2HE = CreateObject("System.Collections.ArrayList")
Set tags_12H = CreateObject("System.Collections.ArrayList")
Set tags_1D = CreateObject("System.Collections.ArrayList")
Set tags_1M = CreateObject("System.Collections.ArrayList")
Set WinCCTags = CreateObject("Scripting.Dictionary")

' parse line includes '=' in ini file
Function parse_line(lineText)
  Dim item, pair, tagdesc, tagname, tag
  pair = Split(lineText, "=", 2)
  If ubound(pair) = 1 Then
    tagdesc = Split(Trim(pair(1)), ",", 3)
    If ubound(tagdesc) = 2 Then
      Set item = CreateObject("Scripting.Dictionary")
      item.Add "name", Trim(pair(0))
      item.Add "tagname", tagdesc(0)
      item.Add "valid", tagdesc(1)
      select case tagdesc(2)
        case "1minute"
          tags_1N.add item
        case "10minute"
          tags_10N.add item
        case "30minute"
          tags_30N.add item
        case "2hoursO"
          tags_2HO.add item
        case "2hoursE"
          tags_2HE.add item
        case "12hours"
          tags_12H.add item
        case "1day"
          tags_1D.add item
        case "1month"
          tags_1M.add item
        case else
          tags_1H.add item
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
  Dim stm : Set stm = CreateObject("Adodb.Stream")
  stm.Type = 2 ' adTypeText
  stm.mode = 3 ' adModeRead
  stm.charset = "utf-8"
  stm.Open
  stm.loadfromfile CONF 
  Do Until stm.EOS
    Dim section, comment, strLine
    strLine = Trim(stm.ReadText(-2))
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
  stm.Close
  Set stm = Nothing
End Sub

Sub saveTag(name, value, valid)
  Dim tagvalue, sSQL
  If valid.Read = 1 Then
    sSQL = "INSERT OR IGNORE INTO tags VALUES (" & timestamp & ", '" & name & "', " & value & ");"
    On Error Resume Next
    conn.Execute sSQL
    If Err <> 0 Then
      HMIRuntime.Trace "error on execute: " & sSQL & vbCrLf
    Else
      HMIRuntime.Trace "save: " & timestamp & ", '" & name & "', " & value & vbCrLf
    End If
    On Error GoTo 0
  End If
End Sub

' archive tags to sqlite
Sub saveTags(tags)
  Dim tag, tagname, tagvalue, valid, item
  On Error Resume Next
  conn.Open conn_str
  If Err <> 0 Then
    HMIRuntime.Trace "can't open DB: " & Err.Description
    conn.Close
    Return
  End If
  On Error GoTo 0

  For Each item In tags
    tagname = item("tagname")
    valid = item("valid")
    If WinCCTags.Exists(tagname) AND WinCCTags.Exists(valid) Then
      tagvalue = WinCCTags(item("tagname")).Read
      saveTag item("name"), tagvalue, WinCCTags(valid)
    Else
      Dim oTag : Set oTag = HMIRuntime.Tags(tagname)
      Dim oTagV : Set oTagV = HMIRuntime.Tags(valid)
      tagvalue = oTag.Read
      oTagV.Read ' test tag.QualityCode
      If 28 = oTag.QualityCode OR 28 = oTagV.QualityCode Then
        tags.Remove item ' remove config item for no such tag
      Else
        WinCCTags.Add tagname, oTag
        If NOT WinCCTags.Exists(valid) Then
          WinCCTags.Add valid, oTagV
        End If
        saveTag item("name"), tagvalue, oTagV
      End If
    End If
  Next
  conn.Close
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
    conn.Close
    getHisTag = -100000.0
    Return
  End If
  On Error GoTo 0

  On Error Resume Next
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
  conn.Close
  On Error GoTo 0
End Function

' archive function for WinCC actions
Sub save1N()
  setDT Now
  saveTags(tags_1N)
End Sub
Sub save10N()
  setDT Now
  saveTags(tags_10N)
  If N = 30 OR N = 0 Then
    saveTags(tags_30N)
  End If
End Sub
Sub save1H()
  setDT Now
  saveTags(tags_1H)
  If H MOD 2 = 0 Then
    saveTags(tags_2HE)
  Else
    saveTags(tags_2HO)
  End If
End Sub
Sub save12H()
  setDT Now
  saveTags(tags_12H)
  If H = 0 Then
    saveTags(tags_1D)
  End If
End Sub
Sub save1M()
  setDT Now
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
        QualityCode = 80
    End Sub
    Public Function Read()
        Read = Value
    End Function
    Public Function Write(v)
        Value = v
    End Function
    Public QualityCode
End Class

Dim HMIRuntime : set HMIRuntime = new Runtime

' 预置几个WinCC变量
Dim kv
For Each kv in Array( _
  Array("template_folder", Left(wscript.scriptfullname,InStrRev(wscript.scriptfullname,"\")) & "template\"), _
  Array("GR_S7/AIT0201.WIO", 0.3), _
  Array("GR_S7/AIT0201.work_F", 1), _
  Array("GR_S7/Flow33.work_F", 1), _
  Array("GR_S7/Flow33.work_F1", 1), _
  Array("GR_S7/Flow34.work_F", 1), _
  Array("GR_S7/Flow34.work_F1", 1), _
  Array("GR_S7/Flow33.mass", 19588.7), _
  Array("GR_S7/Flow33.density", 98.3), _
  Array("GR_S7/Flow33.temperature", 38.2), _
  Array("GR_S7/Flow33.volume", 26555.2), _
  Array("GR_S7/Flow33.volume_flow_rate", 21555.2), _
  Array("GR_S7/Flow33.mass_flow_rate", 38.2), _
  Array("GR_S7/Flow33.oil_mass", 33553.2), _
  Array("GR_S7/Flow33.water_mass", 22553.2), _
  Array("GR_S7/Flow34.mass", 69553.2), _
  Array("GR_S7/Flow34.density", 98.2), _
  Array("GR_S7/Flow34.temperature", 38.2), _
  Array("GR_S7/Flow34.volume", 35668.2), _
  Array("GR_S7/Flow34.volume_flow_rate", 33668.2), _
  Array("GR_S7/Flow34.mass_flow_rate", 38.2), _
  Array("GR_S7/Flow34.oil_mass", 33553.2), _
  Array("GR_S7/Flow34.water_mass", 22553.2) _
)
  Dim HMITag
  set HMITag = new HMITags
  HMITag.Value = kv(1)
  HMIRuntime.pTags.Add kv(0), HMITag
Next

init
save1N()
save10N()
save1H()
save12H()
save1M()
getHisTag "MF33_M", "2022-9-15 4:15"