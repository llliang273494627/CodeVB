Attribute VB_Name = "licence"
Option Explicit
Public Declare Function Authorize Lib "exhsys" (ByVal p1 As Long, ByVal dwIn As Long, ByVal p2 As Long, ByVal dw2 As Long) As Long
Public Declare Function GetXXXX Lib "exhsys" (ByVal buffaddr As Long, ByVal dwSize As Long) As Long

Public Function LicCheck() As Boolean

    Dim sLicKey As String, sMyKey As String, stmp As String
    Dim hdsn(0 To 13) As Byte, byt() As Byte
    Dim tm As Long: tm = Timer
    
    If Authorize(0, tm, 0, 0) <> 810 + tm Then Exit Function '���һ��ҵ�DLL!
    If Dir(App.Path + "\licence.dat") = "" Then
        MsgBox "û���ҵ���Ȩ����ļ�", vbCritical + vbOKOnly, "��ʾ"
        LicCheck = False
        Exit Function
    End If
    '1.��ȡӲ�̵ı��,�����md5����
    tm = GetXXXX(VarPtr(hdsn(0)), 14)
    If tm = 0 Then
'        WL "��ȡXXXX��Ϣʧ��!"
        MsgBox "��Ȩ��֤ʧ��!"
        LicCheck = False
        Exit Function
    End If
    stmp = StrConv(hdsn, vbUnicode)
    sMyKey = Md5_String_Calc(stmp)
    '2.��ȡ��Ȩ�ļ��е�md5,Ȼ��Ƚ�
    Open App.Path + "\licence.dat" For Input As #71
On Error GoTo FILEERR
        sLicKey = Input(1024, #71)
On Error GoTo 0
    Close #71
    sLicKey = Mid(sLicKey, 72, 32)
    '����,��֤��ʼ!
    If sMyKey <> sLicKey Then
        MsgBox "��Ȩ��֤ʧ��!", vbCritical + vbOKOnly, "��ʾ"
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
    MsgBox "��Ȩ�ļ���ʽ����,���������������ϵ", vbCritical + vbOKOnly, "��ʾ"
    LicCheck = False
'    WL "��Ȩ�ļ���ʽ����"
    Exit Function
    
'    gGoBackHighSpeed = 1000
    
End Function

