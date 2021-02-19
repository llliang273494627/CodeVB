VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15300
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   5160
      TabIndex        =   1
      Top             =   2280
      Width           =   8535
   End
   Begin MSCommLib.MSComm MSComDEV1 
      Left            =   3840
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblDevRF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   30
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12120
      TabIndex        =   29
      Top             =   9405
      Width           =   615
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12720
      TabIndex        =   28
      Top             =   9045
      Width           =   615
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   27
      Top             =   9045
      Width           =   615
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   26
      Top             =   9405
      Width           =   975
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   25
      Top             =   9045
      Width           =   615
   End
   Begin VB.Image Image19 
      Height          =   420
      Left            =   10200
      Picture         =   "frmMain.frx":3F563
      Top             =   8520
      Width           =   420
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "右前轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10800
      TabIndex        =   24
      Top             =   8595
      Width           =   1215
   End
   Begin VB.Label lblDevRR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   23
      Top             =   8640
      Width           =   1815
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   22
      Top             =   9405
      Width           =   615
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   21
      Top             =   9045
      Width           =   615
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   9045
      Width           =   615
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   19
      Top             =   9405
      Width           =   975
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   18
      Top             =   9045
      Width           =   615
   End
   Begin VB.Image Image18 
      Height          =   420
      Left            =   4200
      Picture         =   "frmMain.frx":45791
      Top             =   8520
      Width           =   420
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "右后轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   8595
      Width           =   1215
   End
   Begin VB.Label lblDevLF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   16
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12120
      TabIndex        =   15
      Top             =   5205
      Width           =   615
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12720
      TabIndex        =   14
      Top             =   4845
      Width           =   615
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   13
      Top             =   4845
      Width           =   615
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   12
      Top             =   5205
      Width           =   975
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10200
      TabIndex        =   11
      Top             =   4845
      Width           =   615
   End
   Begin VB.Image Image17 
      Height          =   420
      Left            =   10200
      Picture         =   "frmMain.frx":4B9BF
      Top             =   4320
      Width           =   420
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "左前轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10800
      TabIndex        =   10
      Top             =   4395
      Width           =   1215
   End
   Begin VB.Label lblDevLR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5880
      TabIndex        =   9
      Top             =   4395
      Width           =   1815
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6120
      TabIndex        =   8
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   7
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   4800
      Width           =   615
   End
   Begin VB.Image Image16 
      Height          =   420
      Left            =   4200
      Picture         =   "frmMain.frx":51BED
      Top             =   4280
      Width           =   420
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "左后轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Top             =   4350
      Width           =   1215
   End
   Begin VB.Image Image12 
      Height          =   525
      Left            =   600
      Picture         =   "frmMain.frx":57E1B
      Top             =   10440
      Width           =   2355
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "VIN:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label lblOption 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   19.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   4560
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   3120
      Width           =   9285
   End
   Begin VB.Image Image15 
      Height          =   390
      Left            =   14040
      Picture         =   "frmMain.frx":58B8A
      Top             =   105
      Width           =   390
   End
   Begin VB.Image Image14 
      Height          =   405
      Left            =   14640
      Picture         =   "frmMain.frx":5E734
      Top             =   105
      Width           =   435
   End
   Begin VB.Image Image13 
      Height          =   1320
      Left            =   13680
      Picture         =   "frmMain.frx":64605
      Top             =   555
      Width           =   960
   End
   Begin VB.Image Image3 
      Height          =   1320
      Left            =   12690
      Picture         =   "frmMain.frx":6B45C
      Top             =   555
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   11648
      Picture         =   "frmMain.frx":7235D
      Top             =   560
      Width           =   1005
   End
   Begin VB.Image Image11 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":79655
      Top             =   6220
      Width           =   420
   End
   Begin VB.Image Image10 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":7F85D
      Top             =   5280
      Width           =   420
   End
   Begin VB.Image Image9 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":85A65
      Top             =   4420
      Width           =   420
   End
   Begin VB.Image Image8 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":8BC6D
      Top             =   3480
      Width           =   420
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   10005
      Top             =   8415
      Width           =   3960
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   4080
      Top             =   8400
      Width           =   3960
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   10011
      Top             =   4200
      Width           =   3960
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   4080
      Top             =   4200
      Width           =   3960
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   10640
      Picture         =   "frmMain.frx":91E75
      Top             =   560
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TireID As String, AirPressure As Double, Temperature As Double, Acceleration As Double, BAT As String, State As String, TirePostion As String
Dim DataFlag  As Integer
Dim TireIDFlag(1 To 5) As String
Dim ErrLog As New Clog
Dim ExitFlag As Boolean
Dim Flash As Integer
Dim PrintType As Integer
Dim DevErr As Boolean
Dim I As Integer
Dim TestState As Integer
Public Reset As Boolean

Private Sub Form_Load()
    TireIDFlag(1) = ""
    TireIDFlag(2) = ""
    TireIDFlag(3) = ""
    TireIDFlag(4) = ""
    TireIDFlag(5) = ""
    ExitFlag = False
    TestState = 0
    DevErr = False
    DataFlag = 0
    I = 0
     PrintType = GetIniS("Client", "PrintType", "", GetProjectPath() & "Setting.ini")
    
    Call Initstallcom(frmMain, 1)
'    Me.lblNotice.Caption = "正在连接设备......"
'    Me.lblNotice.BackColor = &H8080FF
    lblOption.Caption = "请开启1号VT60设备"
'    frmDevStatus.lblDev1.Caption = "1号设备正在使用"
'    frmDevStatus.lblDev1.ForeColor = &H80000002 '蓝色
'    frmDevStatus.lblDev1.BackColor = &HC000&  '绿色
'    frmmain.= ""
End Sub

Private Sub Image15_Click()
Me.MinButton
End Sub

Private Sub Image4_Click()
    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 3
        lblOption.Caption = "正在检测左后轮"
        lblDevLR.BackColor = &HFF&
        TireIDFlag(3) = ""
        DataFlag = 3
        SetColor ("lblDevLR")
    End If
End Sub



Private Sub Image5_Click()
    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 2
        lblOption.Caption = "正在检测右前轮"
        lblDevLF.BackColor = &HFF&
        TireIDFlag(2) = ""
        DataFlag = 2
        SetColor ("lblDevLF")
    End If
End Sub

Private Sub Image6_Click()
    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 4
        lblOption.Caption = "正在检测右后轮"
        lblDevRR.BackColor = &HFF&
        TireIDFlag(4) = ""
        DataFlag = 4
        SetColor ("lblDevRR")
    End If
End Sub

Private Sub Image7_Click()
    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 1
        lblOption.Caption = "正在检测右前轮"
        lblDevRF.BackColor = &HFF&
        TireIDFlag(1) = ""
        DataFlag = 1
        SetColor ("lblDevRF")
    End If
End Sub

Private Sub MSComDEV1_OnComm()
    Dim recv() As Byte
'    Dim recv As String
    Dim tmp As Variant
    Dim I As Long
    Dim nStatus As Long, n As Long 'nStatus '用于接受状态字

    On Error GoTo EH

    Static OnCommBusy As Boolean

    DoEvents

    Select Case MSComDEV1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

' 错误
      Case comEventBreak   ' 收到 Break。
'        Debug.Print "收到中断"
        lblOption.Caption = "收到中断"
        lblOption.BackColor = &HFF00&
        blnOpenstat = False
        Unload Me
        err.Clear
        Exit Sub

      Case comEventCDTO   ' CD (RLSD) 超时。
      Case comEventCTSTO   ' CTS Timeout。
      Case comEventDSRTO   ' DSR Timeout。
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   '数据丢失。
      Case comEventRxOver '接收缓冲区溢出。
'        MSComm1.InBufferCount = 0
'        AddStrToRTB "接收缓冲区溢出 !" + Chr(10), RGB(50, 0, 0)
      Case comEventRxParity ' Parity 错误。
      Case comEventTxFull   '传输缓冲区已满。
'        MsgBox "发送缓冲区已满", vbOKOnly, "警告"
      Case comEventDCB   '获取 DCB] 时意外错误

   ' 事件
      Case comEvCD   ' CD 线状态变化。
        If blnOpenstat = False Then    '状态5
            MSComDEV1.PortOpen = False
        End If
'        MSComm1.PortOpen = False
      Case comEvCTS   ' CTS 线状态变化。
        'VT60关机后触发
      Case comEvDSR   ' DSR 线状态变化。
        'VT60关机后触发
      Case comEvRing   ' Ring Indicator 变化。
        'VT60关机后触发
      Case comEvReceive   ' 收到 RThreshold # of chars.
'        If OnCommBusy = False Then
'            OnCommBusy = True
'            Get_BlueTooth_Packet
'            OnCommBusy = False
'        End If
'            DelayTime 0.05

            tmp = MSComDEV1.Input
            strin = strin & tmp
            tmp = ""

            If Left(Right(strin, 12), 5) = "STATE" Then

                Select Case DataFlag
                    Case 1
                    '提示正在检测状态和下一步状态
                        lblOption.Caption = "正在检测右前轮"
                        Flash = 2
                        lblDevRF.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        WriteDataBase ("RF")
                        TireIDFlag(1) = TireID '&H000000FF& 红色  &H0000FF00& 绿色
'                        lblDevRF.Caption = TireID
                        If TireIDFlag(1) <> TireIDFlag(2) And TireIDFlag(1) <> TireIDFlag(3) And TireIDFlag(1) <> TireIDFlag(4) Then
                            lblDevRF.BackColor = &HFF00&
                            lblDevLF.BackColor = &HFF&
                            DataFlag = DataFlag + 1
                            lblOption.Caption = "右前轮检测完毕,请将设备移到左前轮"
                            lblDevRF.Caption = TireID
                        Else
'                            TireIDFlag(1) = ""
                            lblOption.Caption = "请将手持设备靠近右前轮,重新检测"
                            Flash = 1
                        End If
                    Case 2
                        lblOption.Caption = "正在检测左前轮"
                        Flash = 3
                        lblDevLF.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(2) = TireID
                    '判断是否检测重复
                        If TireIDFlag(2) <> TireIDFlag(1) And TireIDFlag(2) <> TireIDFlag(3) And TireIDFlag(2) <> TireIDFlag(4) Then
                            WriteDataBase ("LF")
                            DataFlag = DataFlag + 1
                            lblDevLF.BackColor = &HFF00&
                            lblDevLR.BackColor = &HFF&
                            lblOption.Caption = "左前轮检测完毕,请将设备移到左后轮"
                            lblDevLF.Caption = TireID
                        Else
'                            TireIDFlag(2) = ""
                            lblOption.Caption = "请将手持设备靠近左前轮,重新检测"
                            Flash = 2
                        End If
                    Case 3
                        lblOption.Caption = "正在检测左后轮"
                        Flash = 4
                        lblDevLR.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(3) = TireID
                    '判断是否检测重复
                        If TireIDFlag(3) <> TireIDFlag(1) And TireIDFlag(3) <> TireIDFlag(2) And TireIDFlag(3) <> TireIDFlag(4) Then
                            WriteDataBase ("LR")
                            DataFlag = DataFlag + 1
                            lblDevRR.BackColor = &HFF&
                            lblDevLR.BackColor = &HFF00&
                            lblOption.Caption = "左后轮检测完毕,请将设备移到右后轮"
                            lblDevLR.Caption = TireID
                        Else
'                            TireIDFlag(3) = ""
                            lblOption.Caption = "请将手持设备靠近左后轮,重新检测"
                            Flash = 3
                        End If
                    Case 4
                        lblOption.Caption = "正在检测右后轮"
                        Flash = 5
                        lblDevRR.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(4) = TireID
                    '判断是否检测重复
                        If TireIDFlag(4) <> TireIDFlag(1) And TireIDFlag(4) <> TireIDFlag(2) And TireIDFlag(4) <> TireIDFlag(3) Then
                            WriteDataBase ("RR")
                            Call SaveToDB
                            lblDevRR.BackColor = &HFF00&
'                            If optTire5.value = True Then
'                                lblDevST.BackColor = &HFF&
'                                DataFlag = DataFlag + 1
'                                lblOption.Caption = "右后轮检测完毕,请将设备移到备胎"
'                                lblDevRR.Caption = TireID
'                            Else
'                                TireIDFlag(4) = ""
'                                lblDevRR.backcolor = &HFF&
                                lblOption.Caption = "右后轮检测完毕"
                                lblDevRR.Caption = TireID
                                CheckFinish (4)
'                            End If

                        Else
                            lblOption.Caption = "请将手持设备靠近左后轮,重新检测"
                            Flash = 4
                        End If
'                    Case 5    '
'                        lblOption.Caption = "正在检测备胎"
'                        lblDevST.BackColor = &HFF&
'                        Derult = ""
'                        Derult = strin
''                        Debug.Print strin
'                        strin = ""
'                        Call SplitData(Derult)
'                        Derult = ""
'                        TireIDFlag(5) = TireID
'                    '判断是否检测重复
'                        If TireIDFlag(5) <> TireIDFlag(1) And TireIDFlag(5) <> TireIDFlag(2) And TireIDFlag(5) <> TireIDFlag(3) And TireIDFlag(5) <> TireIDFlag(4) Then
'                            WriteDataBase ("ST")
''                            lblDevST.BackColor = &HFF00&
'                            lblOption.Caption = "备胎检测完毕"
'                            lblDevST.Caption = TireID
'                            CheckFinish (5)
'                        Else
'                            lblOption.Caption = "请将手持设备靠近备用轮胎,重新检测"
'                            Flash = 5
'                        End If
                End Select
            End If
      Case comEvSend   ' 传输缓冲区有 Sthreshold 个字符                     '
      Case comEvEOF   ' 输入数据流中发现 EOF 字符
   End Select

'    Timer2.Interval = 500
    Exit Sub
EH: 'error handler
End Sub
Private Sub StartUp()
    On Error GoTo StartUpERR
        If blnOpenstat = True Then
            If MSComDEV1.PortOpen = True Then
                MSComDEV1.PortOpen = False
            End If
            DoEvents
            MSComDEV1.PortOpen = True

        Else
'           Call Initstallcom(frmDev1)
            If MSComDEV1.PortOpen = True Then
                MSComDEV1.PortOpen = False
            End If
            DoEvents
            MSComDEV1.PortOpen = True
        End If
         '开始工作
'          lblNotice.Caption = "设备打开正常"
'          lblNotice.BackColor = &HC000&
         lblOption.Caption = "请将手持设置靠近右前轮"
         sleep 200
         lblOption.Caption = "请将手持设置靠近右前轮"
         DataFlag = 1
         Flash = 1
         Exit Sub
StartUpERR:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "1号设备StartUp", "", "出错！"
    DevErr = True
End Sub
'******************************************************************************
'** 函 数 名：SplitData
'** 输    入：
'** 输    出：
'** 功能描述：解析数据，保存到局部变量
'** 全局变量：
'** 作    者：李操、何孝钦
'** 邮    箱：tonylicao@163.com、hexiaoqin027@163.com
'** 日    期：2009-4-11
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub SplitData(ByVal Strtemp As String)
    'TG1B ND ; CD0AAEA8 ;  101   kPa ;  15鳦 ;   0.5  g ; BAT:OK ; STATE :E7
On err GoTo Spliterr
    Dim tmp() As String
    Dim I As Integer
    tmp() = Split(Strtemp, ";")
    TireID = Trim(tmp(1))
'    AirPressure = CDbl(Trim(Replace(tmp(2), "kPa", "")))
'    Temperature = CDbl(Trim(Replace(tmp(3), "鳦", "")))
'    Acceleration = CDbl(Trim(Replace(tmp(4), "g", "")))
'    BAT = Trim(Replace(tmp(5), "BAT:", ""))
'    State = Trim(Replace(tmp(6), "STATE :", ""))
    Exit Sub
'    Debug.Print Acceleration
Spliterr:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "传感器其他数据处理错误", "", "出错！"
End Sub


'******************************************************************************
'** 函 数 名：CheckFinish
'** 输    入：
'** 输    出：
'** 功能描述：数据完整性判断
'** 全局变量：
'** 作    者：李操、何孝钦
'** 邮    箱：tonylicao@163.com、hexiaoqin027@163.com
'** 日    期：2009-4-11
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub CheckFinish(ByVal TireNO As Integer)
    If TireNO = 4 Then
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Then

            Else
                lblOption.Caption = "胎压检测完毕"
'                Unload Me
                ExitFlag = True
            End If
    Else
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Or TireIDFlag(5) = "" Then

            Else
                lblOption.Caption = "胎压检测完毕"
'                Unload Me
                ExitFlag = True
            End If
    End If
End Sub


'******************************************************************************
'** 函 数 名：PrintResult
'** 输    入：
'** 输    出：
'** 功能描述：数据完整性判断
'** 全局变量：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com
'** 日    期：2009-5-21
'** 修 改 者：
'** 日    期：
'** 版    本：1.1005
'******************************************************************************
Private Sub PrintResult()
Dim lbl(5) As String
lbl(1) = "右前轮："
lbl(2) = "左前轮："
lbl(3) = "左后轮："
lbl(4) = "右后轮："
lbl(5) = "备用胎："
On err GoTo err
Dim J As Integer
    J = 1
    If optTire4.value = True Then
        DataReport1.Sections(1).Controls("lbl5").Visible = False
        For J = 1 To 4
            If TireIDFlag(J) = "" Then
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &HFF&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & "不合格"
            Else
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &H0&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & TireIDFlag(J)
            End If
        Next J
    Else
        For J = 1 To 5
            If TireIDFlag(J) = "" Then
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &HFF&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & "不合格"
            Else
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &H0&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & TireIDFlag(J)
            End If
        Next J
    End If
    DataReport1.Sections("section1").Controls("lblVIN").Caption = "VIN:" & Dev1VIN
    DataReport1.Sections("section1").Controls("lblDate").Caption = "Date:" & Format(Now, "yyyy-mm-dd")
    DataReport1.Sections("section1").Controls("lblTime").Caption = "Time:" & Format(Now, "hh:mm:ss")
    DataReport1.PrintReport False, rptRangeFromTo
    Exit Sub
err:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "打印模块", "", "出错！"
    MsgBox "打印失败！", vbExclamation
End Sub


'******************************************************************************
'** 函 数 名：SaveToDB
'** 输    入：
'** 输    出：
'** 功能描述：一次性将所有数据存入数据库
'** 全局变量：
'** 作    者：杨帅
'** 邮    箱：tonylicao@163.com、hexiaoqin027@163.com
'** 日    期：2011-6-21
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub SaveToDB()
    On Error GoTo SaveToDB_Err
        
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    
    On Error GoTo Local_Conn
    cnn.ConnectionTimeout = 3
    cnn.Open PostgresStr
    
    
    

    
    rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
        rst.AddNew
    End If
    rst.Fields("VIN").value = Dev1VIN
    rst("VIS").value = Right(Dev1VIN, 8)
    'rst(TireField).value = TireID
    rst("ID020").value = TireIDFlag(1)
    rst("ID021").value = TireIDFlag(2)
    rst("ID022").value = TireIDFlag(3)
    rst("ID023").value = TireIDFlag(4)
    
    
    
    rst("TestTime").value = Now
    rst.Fields("TestState") = TestState
    rst.Fields("UploadSign") = False
    rst.Fields("DownloadSign") = False
    rst.Fields("Dev") = "201"
    rst.Update
    rst.Close
    cnn.Close
    
    
Local_Conn:
           
'远程库存储
    On Error GoTo Remote_Conn
    cnn.ConnectionTimeout = 3
    cnn.Open RMTPostgresStr
    rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
        rst.AddNew
    End If
    rst.Fields("VIN").value = Dev1VIN
    rst("VIS").value = Right(Dev1VIN, 8)
    'rst(TireField).value = TireID
    
    rst("ID020").value = TireIDFlag(1)
    rst("ID021").value = TireIDFlag(2)
    rst("ID022").value = TireIDFlag(3)
    rst("ID023").value = TireIDFlag(4)
    rst("TestTime").value = Now
    rst.Fields("TestState") = TestState
    rst.Fields("UploadSign") = False
    rst.Fields("DownloadSign") = False
    rst.Fields("Dev") = "201"
    rst.Update
    rst.Close
    cnn.Close
            
Remote_Conn:
    
    
    
    
    Exit Sub
SaveToDB_Err:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "1号设备SaveToDB", "", "出错！"
End Sub

'******************************************************************************
'** 函 数 名：WriteDataBase
'** 输    入：
'** 输    出：
'** 功能描述：数据存入数据库
'** 全局变量：
'** 作    者：李操、何孝钦
'** 邮    箱：tonylicao@163.com、hexiaoqin027@163.com
'** 日    期：2009-4-11
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub WriteDataBase(ByVal StrPostion As String)
    On Error GoTo WriteDataBaseErr
'        Dim cnn As New ADODB.Connection
'        Dim rst As New ADODB.Recordset
        Dim TireField As String
        Select Case StrPostion
            Case "RF"
                If TireID <> "00000000" Or TireID <> "" Then
                    TestState = TestState + 8
                End If
                TireField = "ID020"
            Case "LF"
                If TireID <> "00000000" Or TireID <> "" Then
                    TestState = TestState + 4
                End If
                TireField = "ID022"
            Case "RR"
                If TireID <> "00000000" Or TireID <> "" Then
                    TestState = TestState + 2
                End If
                TireField = "ID021"
            Case "LR"
                If TireID <> "00000000" Or TireID <> "" Then
                    TestState = TestState + 1
                End If
                TireField = "ID023"
            Case "ST"
                TireField = "ID024"
        End Select
        
'        PostgresStr = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=PostgreSQL30"
' 本地库存储
'        cnn.ConnectionString = PostgresStr
'        cnn.Open
'        rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
'           If rst.EOF Then
'               rst.AddNew
'           End If
'            rst.Fields("VIN").value = Dev1VIN
'            rst("VIS").value = Right(Dev1VIN, 8)
'            rst(TireField).value = TireID
'            rst("TestTime").value = Now
'            rst.Fields("TestState") = TestState
'            rst.Fields("UploadSign") = False
'            rst.Fields("DownloadSign") = False
'            rst.Fields("Dev") = "201"
'            rst.Update
'            rst.Close
'            cnn.Close
            
'远程库存储
'        cnn.ConnectionString = RMTPostgresStr
'        cnn.Open
'        rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
'           If rst.EOF Then
'               rst.AddNew
'           End If
'            rst.Fields("VIN").value = Dev1VIN
'            rst("VIS").value = Right(Dev1VIN, 8)
'            rst(TireField).value = TireID
'            rst("TestTime").value = Now
'            rst.Fields("TestState") = TestState
'            rst.Fields("UploadSign") = False
'            rst.Fields("DownloadSign") = False
'            rst.Fields("Dev") = "201"
'            rst.Update
'            rst.Close
'            cnn.Close
            
        Exit Sub
WriteDataBaseErr:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "1号设备WriteDataBase", "", "出错！"
End Sub


Private Sub SetColor(ByVal strLbl As String)
    Dim X As Control
    If optTire4.value = True Then
        For Each X In frmDev1
            If Left(X.Name, 6) = "lblDev" And X.Name <> strLbl And X.Name <> lblDevST.Name Then
                If X.BackColor = &HFF& Then X.BackColor = &H80000012
            End If
        Next
    Else
        For Each X In frmDev1
            If Left(X.Name, 6) = "lblDev" And X.Name <> strLbl Then
                If X.BackColor = &HFF& Then X.BackColor = &H80000012
            End If
        Next
    End If
End Sub
