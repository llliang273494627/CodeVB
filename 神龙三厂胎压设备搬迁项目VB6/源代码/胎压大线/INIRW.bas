Attribute VB_Name = "winini"
'////////////////////windows ini 标准配置文件动态链接文件处理库  ///////////////////

Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "KERNEL32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileInt Lib "KERNEL32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long



Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long

Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

Const WH_KEYBOARD = 2

Dim hHook As Long



Public Function GetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal DefString As String, ByVal Path As String) As String
    Dim ResultString As String * 144, Temp As Integer
    Dim s As String, i As Integer
    
    Temp% = GetPrivateProfileString(SectionName, KeyWord, "", ResultString, 144, Path)
    '检索关键词的值
    If Temp% > 0 Then   '关键词的值不为空
        s = ""
        For i = 1 To 144
            If Asc(Mid$(ResultString, i, 1)) = 0 Then
                Exit For
            Else
                s = s & Mid$(ResultString, i, 1)
            End If
        Next
    Else
        Temp% = WritePrivateProfileString(SectionName, KeyWord, DefString, Path)
        '将缺省值写入INI文件
        s = DefString
    End If
    GetIniS = s
End Function

Public Function GetIniN(ByVal SectionName As String, ByVal KeyWord As String, _
    ByVal DefValue As Long, ByVal Path As String) As Long
    Dim d As Long, s As String
    
    d = DefValue
    GetIniN = GetPrivateProfileInt(SectionName, KeyWord, DefValue, Path)
    If d <> DefValue Then
        s = "" & d
        d = WritePrivateProfileString(SectionName, KeyWord, s, Path)
    End If
End Function

Public Sub SetIniS(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValStr As String, ByVal Path As String)
    Dim res%
    res% = WritePrivateProfileString(SectionName, KeyWord, ValStr, Path)
End Sub

Public Sub SetIniN(ByVal SectionName As String, ByVal KeyWord As String, ByVal ValInt As Integer, ByVal Path As String)
    Dim res%, s$
    s$ = str$(ValInt)
    res% = WritePrivateProfileString(SectionName, KeyWord, s$, Path)
End Sub

Public Sub DeleteAllKeyWords(ByVal SectionName As String, ByVal Path As String)
    WritePrivateProfileSection SectionName, "", Path
End Sub

Public Sub DeleteKeyWord(ByVal SectionName As String, ByVal KeyWord As String, ByVal Path As String)
    WritePrivateProfileString SectionName, KeyWord, "", Path
End Sub

Public Function GetNode(ByVal change_str As String, ByVal findstr As String, ByVal strtype As Integer) As String
'Dim nodetype() As String
'nodetype = Split(change_str, "||")
'MsgBox change_str
'MsgBox nodetype(0)

Dim Node() As String
Node = Split(change_str, vbCrLf)

Dim NodeStr() As String


For i = 0 To UBound(Node)
  If Trim(Node(i)) <> "" Then
    NodeStr = Split(Node(i), "*")
      If NodeStr(0) = findstr Then
         If Node(i) <> findstr & "*" Then
            GetNode = NodeStr(1)
            Exit Function
         Else
            GetNode = "0"
            Exit Function
         End If
      End If
  End If
Next i

GetNode = "0"
End Function

