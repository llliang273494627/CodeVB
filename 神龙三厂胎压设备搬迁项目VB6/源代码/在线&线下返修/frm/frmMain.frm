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
   StartUpPosition =   3  '����ȱʡ
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
         Name            =   "����"
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
      Caption         =   "ѹ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǰ�֣�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�Һ��֣�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǰ�֣�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����֣�"
      BeginProperty Font 
         Name            =   "����"
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
'    Me.lblNotice.Caption = "���������豸......"
'    Me.lblNotice.BackColor = &H8080FF
    lblOption.Caption = "�뿪��1��VT60�豸"
'    frmDevStatus.lblDev1.Caption = "1���豸����ʹ��"
'    frmDevStatus.lblDev1.ForeColor = &H80000002 '��ɫ
'    frmDevStatus.lblDev1.BackColor = &HC000&  '��ɫ
'    frmmain.= ""
End Sub

Private Sub Image15_Click()
Me.MinButton
End Sub

Private Sub Image4_Click()
    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 3
        lblOption.Caption = "���ڼ�������"
        lblDevLR.BackColor = &HFF&
        TireIDFlag(3) = ""
        DataFlag = 3
        SetColor ("lblDevLR")
    End If
End Sub



Private Sub Image5_Click()
    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 2
        lblOption.Caption = "���ڼ����ǰ��"
        lblDevLF.BackColor = &HFF&
        TireIDFlag(2) = ""
        DataFlag = 2
        SetColor ("lblDevLF")
    End If
End Sub

Private Sub Image6_Click()
    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 4
        lblOption.Caption = "���ڼ���Һ���"
        lblDevRR.BackColor = &HFF&
        TireIDFlag(4) = ""
        DataFlag = 4
        SetColor ("lblDevRR")
    End If
End Sub

Private Sub Image7_Click()
    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 1
        lblOption.Caption = "���ڼ����ǰ��"
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
    Dim nStatus As Long, n As Long 'nStatus '���ڽ���״̬��

    On Error GoTo EH

    Static OnCommBusy As Boolean

    DoEvents

    Select Case MSComDEV1.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

' ����
      Case comEventBreak   ' �յ� Break��
'        Debug.Print "�յ��ж�"
        lblOption.Caption = "�յ��ж�"
        lblOption.BackColor = &HFF00&
        blnOpenstat = False
        Unload Me
        err.Clear
        Exit Sub

      Case comEventCDTO   ' CD (RLSD) ��ʱ��
      Case comEventCTSTO   ' CTS Timeout��
      Case comEventDSRTO   ' DSR Timeout��
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   '���ݶ�ʧ��
      Case comEventRxOver '���ջ����������
'        MSComm1.InBufferCount = 0
'        AddStrToRTB "���ջ�������� !" + Chr(10), RGB(50, 0, 0)
      Case comEventRxParity ' Parity ����
      Case comEventTxFull   '���仺����������
'        MsgBox "���ͻ���������", vbOKOnly, "����"
      Case comEventDCB   '��ȡ DCB] ʱ�������

   ' �¼�
      Case comEvCD   ' CD ��״̬�仯��
        If blnOpenstat = False Then    '״̬5
            MSComDEV1.PortOpen = False
        End If
'        MSComm1.PortOpen = False
      Case comEvCTS   ' CTS ��״̬�仯��
        'VT60�ػ��󴥷�
      Case comEvDSR   ' DSR ��״̬�仯��
        'VT60�ػ��󴥷�
      Case comEvRing   ' Ring Indicator �仯��
        'VT60�ػ��󴥷�
      Case comEvReceive   ' �յ� RThreshold # of chars.
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
                    '��ʾ���ڼ��״̬����һ��״̬
                        lblOption.Caption = "���ڼ����ǰ��"
                        Flash = 2
                        lblDevRF.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        WriteDataBase ("RF")
                        TireIDFlag(1) = TireID '&H000000FF& ��ɫ  &H0000FF00& ��ɫ
'                        lblDevRF.Caption = TireID
                        If TireIDFlag(1) <> TireIDFlag(2) And TireIDFlag(1) <> TireIDFlag(3) And TireIDFlag(1) <> TireIDFlag(4) Then
                            lblDevRF.BackColor = &HFF00&
                            lblDevLF.BackColor = &HFF&
                            DataFlag = DataFlag + 1
                            lblOption.Caption = "��ǰ�ּ�����,�뽫�豸�Ƶ���ǰ��"
                            lblDevRF.Caption = TireID
                        Else
'                            TireIDFlag(1) = ""
                            lblOption.Caption = "�뽫�ֳ��豸������ǰ��,���¼��"
                            Flash = 1
                        End If
                    Case 2
                        lblOption.Caption = "���ڼ����ǰ��"
                        Flash = 3
                        lblDevLF.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(2) = TireID
                    '�ж��Ƿ����ظ�
                        If TireIDFlag(2) <> TireIDFlag(1) And TireIDFlag(2) <> TireIDFlag(3) And TireIDFlag(2) <> TireIDFlag(4) Then
                            WriteDataBase ("LF")
                            DataFlag = DataFlag + 1
                            lblDevLF.BackColor = &HFF00&
                            lblDevLR.BackColor = &HFF&
                            lblOption.Caption = "��ǰ�ּ�����,�뽫�豸�Ƶ������"
                            lblDevLF.Caption = TireID
                        Else
'                            TireIDFlag(2) = ""
                            lblOption.Caption = "�뽫�ֳ��豸������ǰ��,���¼��"
                            Flash = 2
                        End If
                    Case 3
                        lblOption.Caption = "���ڼ�������"
                        Flash = 4
                        lblDevLR.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(3) = TireID
                    '�ж��Ƿ����ظ�
                        If TireIDFlag(3) <> TireIDFlag(1) And TireIDFlag(3) <> TireIDFlag(2) And TireIDFlag(3) <> TireIDFlag(4) Then
                            WriteDataBase ("LR")
                            DataFlag = DataFlag + 1
                            lblDevRR.BackColor = &HFF&
                            lblDevLR.BackColor = &HFF00&
                            lblOption.Caption = "����ּ�����,�뽫�豸�Ƶ��Һ���"
                            lblDevLR.Caption = TireID
                        Else
'                            TireIDFlag(3) = ""
                            lblOption.Caption = "�뽫�ֳ��豸���������,���¼��"
                            Flash = 3
                        End If
                    Case 4
                        lblOption.Caption = "���ڼ���Һ���"
                        Flash = 5
                        lblDevRR.BackColor = &HFF&
                        Derult = ""
                        Derult = strin
'                        Debug.Print strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(4) = TireID
                    '�ж��Ƿ����ظ�
                        If TireIDFlag(4) <> TireIDFlag(1) And TireIDFlag(4) <> TireIDFlag(2) And TireIDFlag(4) <> TireIDFlag(3) Then
                            WriteDataBase ("RR")
                            Call SaveToDB
                            lblDevRR.BackColor = &HFF00&
'                            If optTire5.value = True Then
'                                lblDevST.BackColor = &HFF&
'                                DataFlag = DataFlag + 1
'                                lblOption.Caption = "�Һ��ּ�����,�뽫�豸�Ƶ���̥"
'                                lblDevRR.Caption = TireID
'                            Else
'                                TireIDFlag(4) = ""
'                                lblDevRR.backcolor = &HFF&
                                lblOption.Caption = "�Һ��ּ�����"
                                lblDevRR.Caption = TireID
                                CheckFinish (4)
'                            End If

                        Else
                            lblOption.Caption = "�뽫�ֳ��豸���������,���¼��"
                            Flash = 4
                        End If
'                    Case 5    '
'                        lblOption.Caption = "���ڼ�ⱸ̥"
'                        lblDevST.BackColor = &HFF&
'                        Derult = ""
'                        Derult = strin
''                        Debug.Print strin
'                        strin = ""
'                        Call SplitData(Derult)
'                        Derult = ""
'                        TireIDFlag(5) = TireID
'                    '�ж��Ƿ����ظ�
'                        If TireIDFlag(5) <> TireIDFlag(1) And TireIDFlag(5) <> TireIDFlag(2) And TireIDFlag(5) <> TireIDFlag(3) And TireIDFlag(5) <> TireIDFlag(4) Then
'                            WriteDataBase ("ST")
''                            lblDevST.BackColor = &HFF00&
'                            lblOption.Caption = "��̥������"
'                            lblDevST.Caption = TireID
'                            CheckFinish (5)
'                        Else
'                            lblOption.Caption = "�뽫�ֳ��豸����������̥,���¼��"
'                            Flash = 5
'                        End If
                End Select
            End If
      Case comEvSend   ' ���仺������ Sthreshold ���ַ�                     '
      Case comEvEOF   ' �����������з��� EOF �ַ�
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
         '��ʼ����
'          lblNotice.Caption = "�豸������"
'          lblNotice.BackColor = &HC000&
         lblOption.Caption = "�뽫�ֳ����ÿ�����ǰ��"
         sleep 200
         lblOption.Caption = "�뽫�ֳ����ÿ�����ǰ��"
         DataFlag = 1
         Flash = 1
         Exit Sub
StartUpERR:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "1���豸StartUp", "", "����"
    DevErr = True
End Sub
'******************************************************************************
'** �� �� ����SplitData
'** ��    �룺
'** ��    ����
'** �����������������ݣ����浽�ֲ�����
'** ȫ�ֱ�����
'** ��    �ߣ���١���Т��
'** ��    �䣺tonylicao@163.com��hexiaoqin027@163.com
'** ��    �ڣ�2009-4-11
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub SplitData(ByVal Strtemp As String)
    'TG1B ND ; CD0AAEA8 ;  101   kPa ;  15�C ;   0.5  g ; BAT:OK ; STATE :E7
On err GoTo Spliterr
    Dim tmp() As String
    Dim I As Integer
    tmp() = Split(Strtemp, ";")
    TireID = Trim(tmp(1))
'    AirPressure = CDbl(Trim(Replace(tmp(2), "kPa", "")))
'    Temperature = CDbl(Trim(Replace(tmp(3), "�C", "")))
'    Acceleration = CDbl(Trim(Replace(tmp(4), "g", "")))
'    BAT = Trim(Replace(tmp(5), "BAT:", ""))
'    State = Trim(Replace(tmp(6), "STATE :", ""))
    Exit Sub
'    Debug.Print Acceleration
Spliterr:
    ErrLog.LogPath = App.Path & "\Log\" & Trim(Date) & ".txt"
    ErrLog.WriteErrInfo "�������������ݴ������", "", "����"
End Sub


'******************************************************************************
'** �� �� ����CheckFinish
'** ��    �룺
'** ��    ����
'** ���������������������ж�
'** ȫ�ֱ�����
'** ��    �ߣ���١���Т��
'** ��    �䣺tonylicao@163.com��hexiaoqin027@163.com
'** ��    �ڣ�2009-4-11
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub CheckFinish(ByVal TireNO As Integer)
    If TireNO = 4 Then
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Then

            Else
                lblOption.Caption = "̥ѹ������"
'                Unload Me
                ExitFlag = True
            End If
    Else
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Or TireIDFlag(5) = "" Then

            Else
                lblOption.Caption = "̥ѹ������"
'                Unload Me
                ExitFlag = True
            End If
    End If
End Sub


'******************************************************************************
'** �� �� ����PrintResult
'** ��    �룺
'** ��    ����
'** ���������������������ж�
'** ȫ�ֱ�����
'** ��    �ߣ����
'** ��    �䣺tonylicao@163.com
'** ��    �ڣ�2009-5-21
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.1005
'******************************************************************************
Private Sub PrintResult()
Dim lbl(5) As String
lbl(1) = "��ǰ�֣�"
lbl(2) = "��ǰ�֣�"
lbl(3) = "����֣�"
lbl(4) = "�Һ��֣�"
lbl(5) = "����̥��"
On err GoTo err
Dim J As Integer
    J = 1
    If optTire4.value = True Then
        DataReport1.Sections(1).Controls("lbl5").Visible = False
        For J = 1 To 4
            If TireIDFlag(J) = "" Then
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &HFF&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & "���ϸ�"
            Else
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &H0&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & TireIDFlag(J)
            End If
        Next J
    Else
        For J = 1 To 5
            If TireIDFlag(J) = "" Then
                DataReport1.Sections("section1").Controls("lbl" & J).ForeColor = &HFF&
                DataReport1.Sections("section1").Controls("lbl" & J).Caption = lbl(J) & "���ϸ�"
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
    ErrLog.WriteErrInfo "��ӡģ��", "", "����"
    MsgBox "��ӡʧ�ܣ�", vbExclamation
End Sub


'******************************************************************************
'** �� �� ����SaveToDB
'** ��    �룺
'** ��    ����
'** ����������һ���Խ��������ݴ������ݿ�
'** ȫ�ֱ�����
'** ��    �ߣ���˧
'** ��    �䣺tonylicao@163.com��hexiaoqin027@163.com
'** ��    �ڣ�2011-6-21
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
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
           
'Զ�̿�洢
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
    ErrLog.WriteErrInfo "1���豸SaveToDB", "", "����"
End Sub

'******************************************************************************
'** �� �� ����WriteDataBase
'** ��    �룺
'** ��    ����
'** �������������ݴ������ݿ�
'** ȫ�ֱ�����
'** ��    �ߣ���١���Т��
'** ��    �䣺tonylicao@163.com��hexiaoqin027@163.com
'** ��    �ڣ�2009-4-11
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
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
' ���ؿ�洢
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
            
'Զ�̿�洢
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
    ErrLog.WriteErrInfo "1���豸WriteDataBase", "", "����"
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
