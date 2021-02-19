Attribute VB_Name = "modStart"
'*********************************************************************************
'** 文件名：modStart.bas
'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
'** 创建人：金正波
'** 日  期：2008-11-20
'** 修改人：
'** 日  期：
'** 描  述：闪屏控制模块
'**
'** 版  本：1.0
'*********************************************************************************
'-----------------------------------本模块控制程序启动闪屏---------------------------------------
Option Explicit
Public ErrLog As New Clog
'Public Sub Main()
''    frmSplash.Show
''    DoEvents
'
''    MDIfrmMain.Show
''    frmVIN.Show
''    frmDevStatus.Show
'    Call MSCommVINInit(frmVIN)
'    Call Initstallcom(frmDev1)
''    Call CheckDev
''    Unload frmSplash
'    MDIfrmMain.Show
'    frmVIN.Show
'    frmDevStatus.Show
''    MDIfrmMain.Show
''    frmVIN.Show
''    frmDevStatus.Show
'    frmVIN.SetFocus


'    If LicCheck = False Then Exit Sub
'    Load MDIfrmMain
'End Sub

'-------------------------------显示密码输入框----------------------------------
Public Sub ShowForm(frmAny As Form, frmEnatic As Form)
    frmAny.Left = frmEnatic.Left + (frmEnatic.Width - frmAny.Width) / 2
    frmAny.Top = frmEnatic.Top + (frmEnatic.Height - frmAny.Height) / 2
    frmAny.Show
End Sub
'-------------------------------显示密码输入框----------------------------------
Public Sub MSCommVINInit(FrmWidows As Form)
    On Error GoTo Err
        FrmWidows.MSComVIN.CommPort = GetIniS("Client", "comVIN", "", GetProjectPath() & "Setting.ini")
        FrmWidows.MSComVIN.InBufferSize = 1024
        FrmWidows.MSComVIN.OutBufferSize = 512
        FrmWidows.MSComVIN.InBufferCount = 0
        FrmWidows.MSComVIN.Settings = "9600,n,8,1"
        FrmWidows.MSComVIN.InputMode = comInputModeText
        FrmWidows.MSComVIN.RTSEnable = True
        FrmWidows.MSComVIN.RThreshold = 1
        FrmWidows.MSComVIN.PortOpen = True
        Exit Sub
Err:
    If GetIniS("Client", "BarCodeScanner", "", GetProjectPath() & "Setting.ini") = 1 Then
        MsgBox "VIN串口扫描枪出错！"
    End If
End Sub

Public Sub SetAdmk(ADobjeck As Object, status As Integer)
    ADobjeck.CompanyName = "武汉华信数据系统有限公司"
    ADobjeck.CompanyCode = "M208290000"
    ADobjeck.DeviceName = "返修区胎压检测设备"
    ADobjeck.DeviceCode = "DSG201"
    ADobjeck.CircleTime = GetIniS("Client", "AdmkScanTime", "", GetProjectPath() & "Setting.ini")
    ADobjeck.status = status
    ADobjeck.RemoteIP = GetIniS("Client", "AdmkRemoteIP", "", GetProjectPath() & "Setting.ini")
    ADobjeck.RemotePort = GetIniS("Client", "AdmkRemotePort", "", GetProjectPath() & "Setting.ini")
    ADobjeck.RunSwitch = True
End Sub
