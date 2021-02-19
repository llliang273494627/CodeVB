VERSION 5.00
Begin VB.Form FrmOption 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   Picture         =   "FrmOption.frx":0000
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox TxtScanComPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   26
      Top             =   1470
      Width           =   615
   End
   Begin VB.TextBox TxtPwd2 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   19
      Top             =   4500
      Width           =   1815
   End
   Begin VB.TextBox TxtPwd1 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   18
      Top             =   4080
      Width           =   1815
   End
   Begin VB.TextBox TxtPwd0 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   17
      Top             =   3660
      Width           =   1815
   End
   Begin VB.TextBox TxtLocalDBConnStr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   2730
      Width           =   5415
   End
   Begin VB.TextBox TxtRemoteDBConnStr 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   10
      Top             =   2310
      Width           =   5415
   End
   Begin VB.TextBox TxtLocalDBDrive 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5520
      TabIndex        =   9
      Top             =   1890
      Width           =   975
   End
   Begin VB.TextBox TxtRemoteServerIP 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   8
      Top             =   1890
      Width           =   1935
   End
   Begin VB.TextBox TxtBlueComPort 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   7
      Top             =   1470
      Width           =   615
   End
   Begin VB.TextBox TxtPopMsgStayTime 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4320
      TabIndex        =   6
      Top             =   1050
      Width           =   615
   End
   Begin VB.TextBox TxtCheckinterval 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1680
      TabIndex        =   5
      Top             =   1050
      Width           =   615
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "扫描枪串口号："
      Height          =   255
      Left            =   3060
      TabIndex        =   27
      Top             =   1500
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "状态检查间隔："
      Height          =   255
      Left            =   390
      TabIndex        =   25
      Top             =   1080
      Width           =   1515
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "VT60蓝牙串口号："
      Height          =   255
      Left            =   180
      TabIndex        =   24
      Top             =   1500
      Width           =   1515
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "远端服务器IP："
      Height          =   255
      Left            =   420
      TabIndex        =   23
      Top             =   1920
      Width           =   1515
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "远端数据库连接："
      Height          =   255
      Left            =   210
      TabIndex        =   22
      Top             =   2340
      Width           =   1515
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "本地数据库连接："
      Height          =   255
      Left            =   210
      TabIndex        =   21
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Image ImgSaveBaseinfo 
      Height          =   465
      Left            =   7320
      Picture         =   "FrmOption.frx":127D5
      Top             =   2580
      Width           =   1515
   End
   Begin VB.Image ImgSavePwd 
      Height          =   465
      Left            =   7320
      Picture         =   "FrmOption.frx":193FE
      Top             =   4440
      Width           =   1515
   End
   Begin VB.Line Line2 
      X1              =   1560
      X2              =   8880
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   8880
      Y1              =   3270
      Y2              =   3270
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "0表示手动关闭"
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "新密码确认："
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   4500
      Width           =   1095
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "新密码："
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "原密码："
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "系统参数设置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   13
      Top             =   25
      Width           =   7575
   End
   Begin VB.Image ImgClose 
      Height          =   285
      Left            =   8520
      Picture         =   "FrmOption.frx":20027
      Top             =   120
      Width           =   285
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "密码修改"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   3150
      Width           =   1455
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "基本设置"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "本地数据库盘符："
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      Height          =   255
      Left            =   5040
      TabIndex        =   2
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "消息框显示时间："
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "秒"
      Height          =   255
      Left            =   2400
      TabIndex        =   0
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "FrmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    TxtCheckinterval.Text = GetIniS("App", "CheckStateInterval", "", AppPath & "setting.ini")
    TxtPopMsgStayTime.Text = GetIniS("App", "PopMsgStayTime", "", AppPath & "setting.ini")
    TxtScanComPort.Text = GetIniS("App", "ScanGunComPort", "", AppPath & "setting.ini")
    TxtBlueComPort.Text = GetIniS("App", "BlueToothComPort", "", AppPath & "setting.ini")
    TxtRemoteServerIP.Text = GetIniS("Net", "RemoteServerIP", "", AppPath & "setting.ini")
    TxtLocalDBDrive.Text = GetIniS("Client", "LocalDBDrive", "", AppPath & "setting.ini")
    TxtRemoteDBConnStr.Text = GetIniS("Client", "DSG101DBConnStr", "", AppPath & "setting.ini")
    TxtLocalDBConnStr.Text = GetIniS("Client", "LocalDBConnStr", "", AppPath & "setting.ini")
End Sub

Private Sub ImgClose_Click()
    Unload Me
End Sub

Private Sub ImgSaveBaseinfo_Click()
    SetIniS "App", "CheckStateInterval", Trim(TxtCheckinterval.Text), AppPath & "setting.ini"
    SetIniS "App", "PopMsgStayTime", Trim(TxtPopMsgStayTime.Text), AppPath & "setting.ini"
    SetIniS "App", "ScanGunComPort", Trim(TxtScanComPort.Text), AppPath & "setting.ini"
    SetIniS "App", "BlueToothComPort", Trim(TxtBlueComPort.Text), AppPath & "setting.ini"
    SetIniS "Net", "RemoteServerIP", Trim(TxtRemoteServerIP.Text), AppPath & "setting.ini"
    SetIniS "Client", "LocalDBDrive", Trim(TxtLocalDBDrive.Text), AppPath & "setting.ini"
    SetIniS "Client", "DSG101DBConnStr", Trim(TxtRemoteDBConnStr.Text), AppPath & "setting.ini"
    SetIniS "Client", "LocalDBConnStr", Trim(TxtLocalDBConnStr.Text), AppPath & "setting.ini"
End Sub

Private Sub ImgSavePwd_Click()
    Dim strOldPsw As String
    Dim strNewPsw1 As String
    Dim strNewPsw2 As String
    
    strOldPsw = Trim(TxtPwd0.Text)
    If strOldPsw = GetIniS("App", "Psw", "", AppPath & "setting.ini") Then
        strNewPsw1 = Trim(TxtPwd1.Text)
        strNewPsw2 = Trim(TxtPwd2.Text)
        If strNewPsw1 = strNewPsw2 Then
            SetIniS "App", "Psw", strNewPsw1, AppPath & "setting.ini"
        End If
    End If
End Sub

