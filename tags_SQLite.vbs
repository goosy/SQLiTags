Option Explicit

const DBF = "D:/config/test.db"
const CNF = "D:/config/config.ini"
Dim conn_str : conn_str = Replace("Driver={SQLite3 ODBC Driver};Database=@FSPEC@;StepAPI=;Timeout=", "@FSPEC@", DBF )
Dim conn : Set conn = CreateObject( "ADODB.Connection" )

Dim timestamp
' 设置时间
Sub setDT(dt_str)
  ' Year(dt), Month(dt), Day(dt), Hour(dt), Minute(dt)
  Dim dt
  If LCase(dt_str) = "now" Then
    dt = Now ' 设置当前时间
  Else
    dt = CDate(dt_str)
  End If
  timestamp = DateDiff( _
    "s", _
    "1970/01/01 00:00:00", _
    FormatDateTime(dt, 2) & " " & FormatDateTime(dt, 4) & ":0" _
  )
End Sub

Dim tags_1I, tags_10I, tags_30I, tags_1H, tags_12H, tags_1D, tags_1M
Set tags_1I = CreateObject("System.Collections.ArrayList")
Set tags_10I = CreateObject("System.Collections.ArrayList")
Set tags_30I = CreateObject("System.Collections.ArrayList")
Set tags_1H = CreateObject("System.Collections.ArrayList")
Set tags_12H = CreateObject("System.Collections.ArrayList")
Set tags_1D = CreateObject("System.Collections.ArrayList")
Set tags_1M = CreateObject("System.Collections.ArrayList")

' 解析ini文件有=号的行
Function parseLine(lineText)
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
          tags_1I.add tag
        case "10minute"
          tags_10I.add tag
        case "30minute"
          tags_30I.add tag
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

' 配置初始化
Sub init
  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

  If Not fso.FileExists(DBF) Then 'if possiable create file
    fso.CreateTextFile DBF, True
    conn.Open conn_str
    HmiRuntime.Trace "connected to " & DBF
    Dim sSQL : sSQL = "CREATE TABLE tags (time TIMESTAMP NOT NULL, name TEXT NOT NULL, value REAL, PRIMARY KEY (time, name))"
    conn.Execute sSQL
    HmiRuntime.Trace "table tags created"
    conn.Close
  End If

  If Not fso.FileExists(CNF) Then
    Msgbox "configure file " & CNF & " not exists!"
    Return
  End If
  Const ForReading = 1
  Dim INIFile : Set INIFile = fso.OpenTextFile(CNF, ForReading)

  Do While Not INIFile.AtEndOfStream
    Dim section, comment, strLine
    strLine = Trim(INIFile.Readline)
    If strLine <> "" Then
      If Left(strLine, 1) = ";" OR Left(strLine, 1) = "#" Then
        comment = Trim(Mid(strLine, 2, Len(strLine) - 1))
      ElseIf Left(strLine, 1) = "[" And Right(strLine, 1) = "]" Then
        section = Trim(Mid(strLine, 2, Len(strLine) - 2))
      ElseIf section = "tags" Then
        parseLine(strLine)
      End If
    End If
  Loop

  INIFile.Close
  Set INIFile = Nothing
End Sub

' 变量归档
Sub saveTags(tags)
  dim tag, valid, tagvalue, SQL
  conn.Open conn_str
  Dim sSQL : sSQL = "CREATE TABLE tags (time TIMESTAMP NOT NULL, name TEXT NOT NULL, value REAL, PRIMARY KEY (time, name))"
  For Each tag In tags
    tagvalue = HMIRuntime.Tags(tag("tagname")).Value
    If HMIRuntime.Tags(tag("valid")).Value = false Then
      tagvalue = "NULL"
    End If
    sSQL = "INSERT OR IGNORE INTO tags VALUES (" & timestamp & ", '" & tag("name") & "', " & tagvalue & ");"
    HmiRuntime.Trace sSQL
    conn.Execute sSQL
  Next
  conn.Close
End Sub

' TODO 获得历史数据
Function getTags(tagname, datastr)
  Dim sSQL : sSQL = "SELECT value FROM tags WHERE " & _
    "name = '" &  tagname & _
    "' AND time = '" & DateDiff("s", "1970/01/01 00:00:00", datastr) & "';"
  conn.Open conn_str
  Dim rs : Set rs = conn.Execute( sSQL )
  If NOT rs.EOF Then
    Dim f : For Each f In rs.Fields
      getTags = f.value
    Next
  Else
    getTags = -100000.0
  End If
  conn.Close
End Function

' 以下由 WinCC 动作调用
Sub save1I()
  saveTags(tags_1I)
End Sub
Sub save10I()
  saveTags(tags_10I)
End Sub
Sub save30I()
  saveTags(tags_30I)
End Sub
Sub save1H()
  saveTags(tags_1H)
End Sub
Sub save12H()
  saveTags(tags_12H)
End Sub
Sub save1D()
  saveTags(tags_1D)
End Sub
Sub save1M()
  saveTags(tags_1M)
End Sub

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
        If rawStr <> "" Then Value = rawStr&":"&Y&"-"&M&"-"&D&" "&H
    End Function
End Class

Dim HMIRuntime : set HMIRuntime = new Runtime

' 预置几个WinCC变量
Dim kv
For Each kv in Array( _
  Array("template_folder", Left(wscript.scriptfullname,InStrRev(wscript.scriptfullname,"\")) & "template\"), _
  Array("GR_S7/AIT0201.WIO", 0.3), _
  Array("GR_S7/AIT0201.work_F", true), _
  Array("GR_S7/Flow33.work_F", true), _
  Array("GR_S7/Flow34.work_F", false), _
  Array("GR_S7/Flow33.mass", 19588.7), _
  Array("GR_S7/Flow33.density", 98.3), _
  Array("GR_S7/Flow33.temperature", 38.2), _
  Array("GR_S7/Flow33.volume", 26555.2), _
  Array("GR_S7/Flow33.mass_flow_rate", 38.2), _
  Array("GR_S7/Flow34.mass", 69553.2), _
  Array("GR_S7/Flow34.density", 98.2), _
  Array("GR_S7/Flow34.temperature", 38.2), _
  Array("GR_S7/Flow34.volume", 35668.2), _
  Array("GR_S7/Flow34.mass_flow_rate", 38.2) _
)
  Dim HMITag
  set HMITag = new HMITags
  HMITag.Value = kv(1)
  HMIRuntime.pTags.Add kv(0), HMITag
Next

init

setDT("now")
save1I()
save10I()
save30I()
save1H()
save12H()
save1D()
save1M()
WScript.Echo getTags("MF33_M", "2022-9-10 15:45")