Attribute VB_Name = "ModGlobel"
Public AppPath As String '应用程序路径
Public AppName As String '应用程序名称
Public PswMode As String '用于标识输入正确密码后弹出窗体的类型

Public WriteFlag As Integer '标识写入状态，0未写入 1正在写入 -1写入发生错误

Public LocalDBConnStr As String '本地数据库连接字符串
Public RemoteDBConnStr As String '远端数据库连接字符串

Public oIOCard As IOControl.IOCard  'IO控制对象

'***************************************************************************
' 显示弹出提示信息框
'***************************************************************************
Public Sub PopMsg(strTitle As String, strMsg As String)
    If Trim(strtile) = "" Then
        strTitle = "提示信息"
    End If
    FrmMsg.LbTitle = strTitle
    FrmMsg.LbMsg = strMsg
    FrmMsg.Show 1
End Sub


'**************************************************************************
' 检查VIN合法性
' 长度17位，（功能可扩展）
'**************************************************************************
Public Function CheckVin(ByVal strVin As String) As Boolean
On Error GoTo CheckVinErr
    Dim Result As Boolean
    
    Result = True
    
    '检查长度是否为17位，否则返回值result=false
    If Len(strVin) <> 17 Then
        Result = False
    End If
    
    CheckVin = Result
    Exit Function
CheckVinErr:
    CheckVin = False
    Exit Function
End Function

'***************************************************************************
' 检查车轮ID合法性
'***************************************************************************
Public Function CheckWheelID(ByVal strlf As String, ByVal strlb As String, _
                    ByVal strrf As String, ByVal strrb As String) As Boolean
On Error GoTo CheckWheelErr
    Dim Result As Boolean
    
    Result = True
    If strlf = "" Or strlb = "" Or strrf = "" Or strrb = "" Then
        Result = False
    End If
    
    CheckWheelID = Result
    Exit Function
CheckWheelErr:
    CheckWheelID = False
    Exit Function
End Function


Public Function GenerateKey(ByRef seed As String, seednumber As String, _
        ByRef key As String, keynumber As Integer, pin As Integer)
        
        pin = 13151720
        Dim i As Integer
        Dim j As Integer
        
        i = 0: j = 0
        
        For i = 0 To seednumber
            pin = pin ^ (LeftMove(pin, 5) + Asc(Mid(seed, i + 1, 1)) + RightMove(pin, 4))
        Next i
        
        keynumber = seednumber
        
        For j = 0 To keynumber
            key = Replace(key, Mid(key, j + 1, 1), Chr((RightMove(pin, j * 8) And &HFF)))
        Next j
End Function


Function LeftMove(ByVal iValue As Integer, iStep As Integer) As Integer
    Dim i As Integer
    For i = 1 To iStep
        iValue = iValue * 2
    Next i
    LeftMove = iValue
End Function

Function RightMove(ByVal iValue As Integer, iStep As Integer) As Integer
    Dim i As Integer
    For i = 1 To iStep
        iValue = iValue / 2
    Next i
    RightMove = iValue
End Function

Public Sub Main()
On Error GoTo Main_Err
    Set oIOCard = New IOControl.IOCard
    frmMain.Show
    Exit Sub
Main_Err:
    MsgBox "初始化参数失败，错误信息：" & Err.Description & "。请检查配置信息！"
End Sub
