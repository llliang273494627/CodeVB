Attribute VB_Name = "licence"
Option Explicit
Public Declare Function Authorize Lib "exhsys" (ByVal p1 As Long, ByVal dwIn As Long, ByVal p2 As Long, ByVal dw2 As Long) As Long
Public Declare Function GetXXXX Lib "exhsys" (ByVal buffaddr As Long, ByVal dwSize As Long) As Long

Public Function LicCheck() As Boolean

    Dim sLicKey As String, sMyKey As String, stmp As String
    Dim hdsn(0 To 13) As Byte, byt() As Byte
    Dim tm As Long: tm = Timer
    
    If Authorize(0, tm, 0, 0) <> 810 + tm Then Exit Function '竟敢换我的DLL!
    If Dir(App.Path + "\licence.dat") = "" Then
        MsgBox "没有找到授权许可文件", vbCritical + vbOKOnly, "提示"
        LicCheck = False
        Exit Function
    End If
    '1.读取硬盘的编号,运算出md5密文
    tm = GetXXXX(VarPtr(hdsn(0)), 14)
    If tm = 0 Then
'        WL "读取XXXX信息失败!"
        MsgBox "授权验证失败!"
        LicCheck = False
        Exit Function
    End If
    stmp = StrConv(hdsn, vbUnicode)
    sMyKey = Md5_String_Calc(stmp)
    '2.读取授权文件中的md5,然后比较
    Open App.Path + "\licence.dat" For Input As #71
On Error GoTo FILEERR
        sLicKey = Input(1024, #71)
On Error GoTo 0
    Close #71
    sLicKey = Mid(sLicKey, 72, 32)
    '哈哈,验证开始!
    If sMyKey <> sLicKey Then
        MsgBox "授权验证失败!", vbCritical + vbOKOnly, "提示"
        LicCheck = False
        Exit Function
    End If
'    gSysInited = False
'    frmComm.Show
'    DE.OpenConnection
'    SysReset
'    gCancel = False
    LicCheck = True
    Exit Function
FILEERR:
    MsgBox "授权文件格式错误,请与软件开发商联系", vbCritical + vbOKOnly, "提示"
    LicCheck = False
'    WL "授权文件格式错误"
    Exit Function
    
'    gGoBackHighSpeed = 1000
    
End Function

