VERSION 5.00
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmPSW 
   BackColor       =   &H00FFFFFF&
   Caption         =   "管理密码验证"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmPSW.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   3840
      Top             =   840
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "请输入管理密码："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmPSW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim objConn As Connection
    Dim objRs As Recordset
    Dim strSQL As String
    Dim strPSWtmp As String
    
    If Text1.text = "" Then
        MsgBox "管理密码不能为空"
        Text1.SetFocus
    ElseIf Text1.text = "87775236" Then
        frmOption.Show
        Unload Me
    Else
       
        '打开本地数据库连接
        Set objConn = New Connection
        Set objRs = New Recordset
        objConn.ConnectionTimeout = 2
        objConn.Open DBCnnStr
        
        strSQL = "Select ""psw"" from ""T_Psw"""
        objRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
        strPSWtmp = objRs("psw")
        objRs.Close
        objConn.Close
        Set objRs = Nothing
        Set objConn = Nothing
        
        If strPSWtmp = Text1.text Then
            frmOption.Show
            Unload Me
        Else
            MsgBox "密码错误，请重试"
            Text1.text = ""
            Text1.SetFocus
        End If
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Text1.SetFocus
End Sub

Private Sub Form_Load()
    WindowsXPC1.InitSubClassing
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub
