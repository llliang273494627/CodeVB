VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "TPMS快速速在线返修"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15300
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":1CFA
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrSetColor 
      Left            =   2520
      Top             =   2610
   End
   Begin VB.Timer TmrCheckState 
      Enabled         =   0   'False
      Left            =   1860
      Top             =   2610
   End
   Begin VB.TextBox txtVIN 
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
   Begin MSCommLib.MSComm MSScanCom 
      Left            =   3600
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Image Image8 
      Height          =   390
      Left            =   120
      Picture         =   "frmMain.frx":41B24
      Stretch         =   -1  'True
      Top             =   90
      Width           =   360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "胎压初始化返修系统"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000018&
      Height          =   285
      Left            =   570
      TabIndex        =   51
      Top             =   150
      Width           =   2415
   End
   Begin VB.Image Image6 
      Height          =   1335
      Left            =   4080
      Top             =   8400
      Width           =   3960
   End
   Begin VB.Image Image7 
      Height          =   1335
      Left            =   10005
      Top             =   8415
      Width           =   3960
   End
   Begin VB.Image Image5 
      Height          =   1335
      Left            =   10011
      Top             =   4200
      Width           =   3960
   End
   Begin VB.Label lbLRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   7410
      TabIndex        =   35
      Top             =   4830
      Width           =   840
   End
   Begin VB.Image Image4 
      Height          =   1335
      Left            =   4080
      Top             =   4200
      Width           =   3960
   End
   Begin VB.Label lbRFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   13380
      TabIndex        =   50
      Top             =   9075
      Width           =   900
   End
   Begin VB.Label lbRFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12750
      TabIndex        =   49
      Top             =   9405
      Width           =   630
   End
   Begin VB.Label lbRFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   10800
      TabIndex        =   48
      Top             =   9075
      Width           =   720
   End
   Begin VB.Label lbRFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12150
      TabIndex        =   47
      Top             =   9075
      Width           =   510
   End
   Begin VB.Label lbRFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11040
      TabIndex        =   46
      Top             =   9405
      Width           =   720
   End
   Begin VB.Label lbRRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5040
      TabIndex        =   45
      Top             =   9405
      Width           =   690
   End
   Begin VB.Label lbRRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6210
      TabIndex        =   44
      Top             =   9075
      Width           =   510
   End
   Begin VB.Label lbRRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4830
      TabIndex        =   43
      Top             =   9075
      Width           =   750
   End
   Begin VB.Label lbRRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6720
      TabIndex        =   42
      Top             =   9405
      Width           =   570
   End
   Begin VB.Label lbRRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   7380
      TabIndex        =   41
      Top             =   9075
      Width           =   540
   End
   Begin VB.Label lbLFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11010
      TabIndex        =   40
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label lbLFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12180
      TabIndex        =   39
      Top             =   4830
      Width           =   510
   End
   Begin VB.Label lbLFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   10770
      TabIndex        =   38
      Top             =   4830
      Width           =   750
   End
   Begin VB.Label lbLFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12750
      TabIndex        =   37
      Top             =   5160
      Width           =   540
   End
   Begin VB.Label lbLFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   13410
      TabIndex        =   36
      Top             =   4830
      Width           =   930
   End
   Begin VB.Label txtLR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   345
      Left            =   5970
      TabIndex        =   9
      Top             =   4335
      Width           =   1815
   End
   Begin VB.Label lbLRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   4830
      TabIndex        =   33
      Top             =   4830
      Width           =   720
   End
   Begin VB.Label lbLRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6750
      TabIndex        =   34
      Top             =   5160
      Width           =   510
   End
   Begin VB.Label lbLRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6210
      TabIndex        =   32
      Top             =   4830
      Width           =   510
   End
   Begin VB.Label lbLRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5070
      TabIndex        =   31
      Top             =   5160
      Width           =   690
   End
   Begin VB.Label txtRF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   345
      Left            =   11940
      TabIndex        =   30
      Top             =   8610
      Width           =   1815
   End
   Begin VB.Label Label27 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Width           =   1275
   End
   Begin VB.Label Label26 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12750
      TabIndex        =   28
      Top             =   9075
      Width           =   1455
   End
   Begin VB.Label Label25 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11550
      TabIndex        =   27
      Top             =   9075
      Width           =   1275
   End
   Begin VB.Label Label24 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10170
      TabIndex        =   26
      Top             =   9405
      Width           =   1755
   End
   Begin VB.Label Label23 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10170
      TabIndex        =   25
      Top             =   9075
      Width           =   1185
   End
   Begin VB.Image picRF 
      Height          =   420
      Left            =   10200
      Picture         =   "frmMain.frx":4381E
      Top             =   8550
      Width           =   420
   End
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "右前轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10710
      TabIndex        =   24
      Top             =   8595
      Width           =   1215
   End
   Begin VB.Label txtRR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   345
      Left            =   5970
      TabIndex        =   23
      Top             =   8610
      Width           =   1815
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Caption         =   "温度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6750
      TabIndex        =   21
      Top             =   9075
      Width           =   615
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5610
      TabIndex        =   20
      Top             =   9075
      Width           =   675
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Caption         =   "模式:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Top             =   9075
      Width           =   735
   End
   Begin VB.Image picRR 
      Height          =   420
      Left            =   4200
      Picture         =   "frmMain.frx":49A4C
      Top             =   8550
      Width           =   420
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "右后轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4740
      TabIndex        =   17
      Top             =   8595
      Width           =   1215
   End
   Begin VB.Label txtLF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   345
      Left            =   11910
      TabIndex        =   16
      Top             =   4350
      Width           =   1815
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "温度:"
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
      Left            =   12690
      TabIndex        =   14
      Top             =   4830
      Width           =   1455
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11580
      TabIndex        =   13
      Top             =   4830
      Width           =   1215
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10140
      TabIndex        =   12
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "模式:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   10140
      TabIndex        =   11
      Top             =   4830
      Width           =   1185
   End
   Begin VB.Image picLF 
      Height          =   420
      Left            =   10200
      Picture         =   "frmMain.frx":4FC7A
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
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10710
      TabIndex        =   10
      Top             =   4335
      Width           =   1215
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "压力:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
      Caption         =   "温度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6780
      TabIndex        =   7
      Top             =   4830
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "电池:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5610
      TabIndex        =   6
      Top             =   4830
      Width           =   705
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "加速度:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
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
         Size            =   12
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
      Top             =   4830
      Width           =   945
   End
   Begin VB.Image picLR 
      Height          =   420
      Left            =   4200
      Picture         =   "frmMain.frx":55EA8
      Top             =   4305
      Width           =   420
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "左后轮："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4710
      TabIndex        =   3
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Image Image12 
      Height          =   525
      Left            =   810
      Picture         =   "frmMain.frx":5C0D6
      Stretch         =   -1  'True
      Top             =   10740
      Width           =   1680
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
      Caption         =   "等待扫描VIN码，开始测量!"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   585
      Left            =   4590
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   3240
      Width           =   9285
   End
   Begin VB.Image Image15 
      Height          =   390
      Left            =   14040
      Picture         =   "frmMain.frx":62971
      Top             =   105
      Width           =   390
   End
   Begin VB.Image Image14 
      Height          =   405
      Left            =   14640
      Picture         =   "frmMain.frx":6851B
      Top             =   90
      Width           =   435
   End
   Begin VB.Image Image13 
      Height          =   1320
      Left            =   13680
      Picture         =   "frmMain.frx":6E3EC
      Top             =   555
      Width           =   960
   End
   Begin VB.Image Image3 
      Height          =   1320
      Left            =   12690
      Picture         =   "frmMain.frx":75243
      Top             =   555
      Width           =   960
   End
   Begin VB.Image Image2 
      Height          =   1320
      Left            =   11648
      Picture         =   "frmMain.frx":7C144
      Top             =   560
      Width           =   1005
   End
   Begin VB.Image ImgRemoteDBState 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":8343C
      Top             =   6220
      Width           =   420
   End
   Begin VB.Image ImgLocalDBState 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":8966A
      Top             =   5280
      Width           =   420
   End
   Begin VB.Image ImgHDDState 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":8F898
      Top             =   4420
      Width           =   420
   End
   Begin VB.Image ImgNetState 
      Height          =   420
      Left            =   960
      Picture         =   "frmMain.frx":95AC6
      Top             =   3480
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   10640
      Picture         =   "frmMain.frx":9BCF4
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
'关闭指定进程
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 260
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS = &H2&

Dim TireID As String, AirPressure As Double, Temperature As Double, Acceleration As Double, BAT As String, State As String, TirePostion As String
Dim ErrLog As New Clog
Dim DataFlag  As Integer
Dim TireIDFlag(1 To 5) As String
Dim TirePreFlag(1 To 5) As String
Dim TireTempFlag(1 To 5) As String
Dim TireAcSpeedFlag(1 To 5) As String
Dim TireBatFlag(1 To 5) As String
Dim TireStateFlag(1 To 5) As String

Dim ExitFlag As Boolean
Dim Flash As Integer
Dim PrintType As Integer
Dim Flag As Integer
Dim TestState As Integer
Public Reset As Boolean

Private RemoteServerIP As String '远端服务器IP
Private LocalDBDrive As String '本地数据库所在磁盘盘符
Private CheckStateInterval As Integer '检查系统各状态的时间周期
Public DevNum As String '该系统的编号，大线返修为201，返修区为301
Private MM As Integer
Private TestCode As String
Private isFormLoad As Boolean
Public isQuit As Boolean

'信号灯相关控制参数（io信号输出端口）
Public LampYellowPort As Integer
Public LampRedPort As Integer
Public LampGreenPort As Integer
Public LampBuzzerPort As Integer
Public HornPort As Integer

'********************************************************************************
' 常量定义
'********************************************************************************
Const AvailableSpace = 100 '本地数据库所在磁盘最小可用空间(MB)

Private Sub Form_Load()
On Error GoTo LoadErr:
    isFormLoad = True
    
    Set oIOCard = New IOControl.IOCard
    
    AppPath = GetProjectPath()
    LocalDBConnStr = GetIniS("Client", "LocalDBConnStr", "", GetProjectPath() & "setting.ini")
    RemoteDBConnStr = GetIniS("Client", "DSG101DBConnStr", "", GetProjectPath() & "setting.ini")
    LocalDBDrive = GetIniS("Client", "LocalDBDrive", "", GetProjectPath() & "setting.ini")
    CheckStateInterval = GetIniN("App", "CheckStateInterval", 0, AppPath & "setting.ini")
    DevNum = GetIniS("App", "DevNum", "", AppPath & "setting.ini")
    RemoteServerIP = GetIniS("Net", "RemoteServerIP", "", AppPath & "setting.ini")
    
    '加载信号灯及喇叭IO输出端口
    LampYellowPort = CInt(GetIniS("IOPort", "LampYellowPort", "", AppPath & "setting.ini"))
    LampRedPort = CInt(GetIniS("IOPort", "LampRedPort", "", AppPath & "setting.ini"))
    LampGreenPort = CInt(GetIniS("IOPort", "LampGreenPort", "", AppPath & "setting.ini"))
    LampBuzzerPort = CInt(GetIniS("IOPort", "LampBuzzerPort", "", AppPath & "setting.ini"))
    HornPort = CInt(GetIniS("IOPort", "HornPort", "", AppPath & "setting.ini"))

    resetList False

    TmrCheckState.Interval = 1000
    TmrCheckState.Enabled = True
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    
    txtVIN.Text = "123"
    txtVIN_KeyPress (13)
    
    resetList False
    
    isFormLoad = False
    
On Error GoTo comerr
    Dim scanComNo As Integer
    scanComNo = GetIniS("App", "ScanGunComPort", "", GetProjectPath() & "Setting.ini")
    MSScanCom.CommPort = scanComNo
    MSScanCom.InBufferSize = 1024
    MSScanCom.OutBufferSize = 512
    MSScanCom.InBufferCount = 0
    MSScanCom.Settings = "9600,n,8,1"
    MSScanCom.InputMode = comInputModeText
    MSScanCom.RTSEnable = True
    MSScanCom.RThreshold = 1
    MSScanCom.PortOpen = True
    Exit Sub
comerr:
    isFormLoad = False
    ErrLog.WriteErrInfo "打开扫描枪串口", "", "出错！" & Err.Description
    AddMessage "打开扫描枪串口时出错!", True, True
    Exit Sub
LoadErr:
    isFormLoad = False
    ErrLog.WriteErrInfo "加载系统参数配置", "", "出错！" & Err.Description
End Sub

Private Sub Image1_Click()
    FrmHistory.Show 1
End Sub

Private Sub Image13_Click()
    FrmHelp.Show 1
End Sub

Private Sub Image14_Click()
    isQuit = False

    PswMode = "exit"
    FrmPsw.Show 1
    
    If isQuit Then
        Dim x As Form
        
        blnOpenstat = False
        If MSComDEV1.PortOpen = True Then
           MSComDEV1.PortOpen = False
        End If
        
        For Each x In Forms
            Unload x
            Set x = Nothing
        Next
        
        Call CloseAllIO '关闭IO输出
        Set oIOCard = Nothing
        Call KillProcess("TirePressure.exe")
    End If
End Sub

Private Sub Image15_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub Image2_Click()
    FrmShowLog.Show 1
End Sub

Private Sub Image3_Click()
    PswMode = "option"
    FrmPsw.Show 1
End Sub

Private Sub Image4_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If
    
    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 3
        AddMessage "正在检测左后轮......", False, True
        txtLR.BackColor = &HFF&
        DataFlag = 3
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
    End If
End Sub

Private Sub Image5_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 2
        AddMessage "正在检测左前轮......", False, True
        txtLF.BackColor = &HFF&
        DataFlag = 2
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
    End If
End Sub

Private Sub Image6_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 4
        AddMessage "正在检测右后轮......", False, True
        txtRR.BackColor = &HFF&
        DataFlag = 4
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
    End If
End Sub

Private Sub Image7_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("你选择了手动操作", vbYesNo) = vbYes Then
        Flash = 1
        AddMessage "正在检测右前轮......", False, True
        txtRF.BackColor = &HFF&
        DataFlag = 1
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
    End If
End Sub

Private Sub MSComDEV1_OnComm()
    Dim recv() As Byte
    Dim tmp As Variant
    Dim i As Long
    Dim nStatus As Long, n As Long 'nStatus '用于接受状态字

    On Error GoTo EH

    Static OnCommBusy As Boolean

    DoEvents

    Select Case MSComDEV1.CommEvent
    ' Handle each event or error by placing
    ' code below each case statement

    ' 错误
      Case comEventBreak   ' 收到 Break。
        'Debug.Print "收到中断"
        AddMessage "收到中断", True, True
        blnOpenstat = False
        'Unload Me
        Err.Clear
        Exit Sub

      Case comEventCDTO   ' CD (RLSD) 超时。
      Case comEventCTSTO   ' CTS Timeout。
      Case comEventDSRTO   ' DSR Timeout。
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   '数据丢失。
      Case comEventRxOver '接收缓冲区溢出。
        'MSComm1.InBufferCount = 0
        'AddStrToRTB "接收缓冲区溢出 !" + Chr(10), RGB(50, 0, 0)
      Case comEventRxParity ' Parity 错误。
      Case comEventTxFull   '传输缓冲区已满。
        'MsgBox "发送缓冲区已满", vbOKOnly, "警告"
      Case comEventDCB   '获取 DCB] 时意外错误

      '事件
      Case comEvCD   ' CD 线状态变化。
        If blnOpenstat = False Then    '状态5
            MSComDEV1.PortOpen = False
        End If
        'MSComm1.PortOpen = False
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

            DelayTime 100

            tmp = MSComDEV1.Input
            strin = strin & tmp
            tmp = ""

            strin = Replace(strin, Chr(10), "")
            strin = Replace(strin, Chr(13), "")

            'ErrLog.WriteOprInfo "接收到VT60返回的数据:" & strin
            If Right(strin, 3) = ":OK" Or StrConv(Right(strin, 3), vbUpperCase) = "LOW" Or Right(strin, 5) = "DRIVE" Or Right(strin, 5) = "LEARN" Then
            
                Select Case DataFlag
                    Case 1
                    '提示正在检测状态和下一步状态
                        AddMessage "正在检测右前轮......", False, True
                        ErrLog.WriteOprInfo "接收到VT60返回的数据:" & strin
                        Flash = 2
                        Derult = ""
                        Derult = strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(1) = TireID
                        TirePreFlag(1) = AirPressure
                        TireTempFlag(1) = Temperature
                        TireAcSpeedFlag(1) = Acceleration
                        TireBatFlag(1) = BAT
                        TireStateFlag(1) = State
                        
                        '判断是否检测重复
                        If TireIDFlag(1) <> TireIDFlag(2) And TireIDFlag(1) <> TireIDFlag(3) And TireIDFlag(1) <> TireIDFlag(4) Then
                            WriteDataBase ("RF")
                            DataFlag = DataFlag + 1
                            txtRF.BackColor = &HC000&
                            ErrLog.WriteOprInfo "右前轮检测结果:" & TireIDFlag(1)
                            AddMessage "右前轮检测完毕,请将设备移到左前轮", False, False
                            txtLF.BackColor = &HFF&
                            
                            txtRF.Caption = TireID
                            lbRFPre.Caption = AirPressure & "kPa"
                            lbRFTemp.Caption = Temperature & "℃"
                            lbRFAcSpeed.Caption = Acceleration & "g"
                            lbRFBattery.Caption = BAT
                            lbRFMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                        Else
                            TireIDFlag(1) = ""
                            TirePreFlag(1) = ""
                            TireTempFlag(1) = ""
                            TireAcSpeedFlag(1) = ""
                            TireBatFlag(1) = ""
                            TireStateFlag(1) = ""
                            AddMessage "请将手持设备靠近右前轮,重新检测", True, True
                            Flash = 1
                                                        
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                            
                        End If
                    Case 2
                        AddMessage "正在检测左前轮......", False, True
                        ErrLog.WriteOprInfo "接收到VT60返回的数据:" & strin
                        Flash = 3
                        Derult = ""
                        Derult = strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(3) = TireID '&H000000FF& 红色  &H0000FF00& 绿色
                        TirePreFlag(3) = AirPressure
                        TireTempFlag(3) = Temperature
                        TireAcSpeedFlag(3) = Acceleration
                        TireBatFlag(3) = BAT
                        TireStateFlag(3) = State
                        
                        'txtLR.Caption = TireID
                        If TireIDFlag(3) <> TireIDFlag(1) And TireIDFlag(3) <> TireIDFlag(2) And TireIDFlag(3) <> TireIDFlag(4) Then
                            WriteDataBase ("LF")
                            txtLF.BackColor = &HC000&
                            DataFlag = DataFlag + 1
                            ErrLog.WriteOprInfo "左前轮检测结果:" & TireIDFlag(3)
                            AddMessage "左前轮检测完毕,请将设备移到左后轮", False, False
                            txtLR.BackColor = &HFF&
                            
                            txtLF.Caption = TireID
                            lbLFPre.Caption = AirPressure & "kPa"
                            lbLFTemp.Caption = Temperature & "℃"
                            lbLFAcSpeed.Caption = Acceleration & "g"
                            lbLFBattery.Caption = BAT
                            lbLFMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                        Else
                            TireIDFlag(3) = ""
                            TirePreFlag(3) = ""
                            TireTempFlag(3) = ""
                            TireAcSpeedFlag(3) = ""
                            TireBatFlag(3) = ""
                            TireStateFlag(3) = ""
                            
                            AddMessage "请将手持设备靠近左前轮,重新检测", True, True
                            Flash = 2
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                            
                        End If
                    Case 3
                        AddMessage "正在检测左后轮......", True
                        ErrLog.WriteOprInfo "接收到VT60返回的数据:" & strin
                        Flash = 4
                        Derult = ""
                        Derult = strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(4) = TireID
                        TirePreFlag(4) = AirPressure
                        TireTempFlag(4) = Temperature
                        TireAcSpeedFlag(4) = Acceleration
                        TireBatFlag(4) = BAT
                        TireStateFlag(4) = State
                        
                    '判断是否检测重复
                        If TireIDFlag(4) <> TireIDFlag(1) And TireIDFlag(4) <> TireIDFlag(2) And TireIDFlag(4) <> TireIDFlag(3) Then
                            WriteDataBase ("LR")
                            DataFlag = DataFlag + 1
                            txtLR.BackColor = &HC000&
                            ErrLog.WriteOprInfo "左后轮检测结果:" & TireIDFlag(4)
                            AddMessage "左后轮检测完毕,请将设备移到右后轮", False, False
                            txtRR.BackColor = &HFF&
                        
                            txtLR.Caption = TireID
                            lbLRPre.Caption = AirPressure & "kPa"
                            lbLRTemp.Caption = Temperature & "℃"
                            lbLRAcSpeed.Caption = Acceleration & "g"
                            lbLRBattery.Caption = BAT
                            lbLRMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                        Else
                            TireIDFlag(4) = ""
                            TirePreFlag(4) = ""
                            TireTempFlag(4) = ""
                            TireAcSpeedFlag(4) = ""
                            TireBatFlag(4) = ""
                            TireStateFlag(4) = ""
                            
                            AddMessage "请将手持设备靠近左后轮,重新检测", True, False
                            Flash = 3
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                            
                        End If
                    Case 4
                        AddMessage "正在检测右后轮......", False, True
                        ErrLog.WriteOprInfo "接收到VT60返回的数据:" & strin
                        Flash = 5
                        Derult = ""
                        Derult = strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(2) = TireID
                        TirePreFlag(2) = AirPressure
                        TireTempFlag(2) = Temperature
                        TireAcSpeedFlag(2) = Acceleration
                        TireBatFlag(2) = BAT
                        TireStateFlag(2) = State
                        
                        '判断是否检测重复
                        If TireIDFlag(2) <> TireIDFlag(1) And TireIDFlag(2) <> TireIDFlag(3) And TireIDFlag(2) <> TireIDFlag(4) Then
                            WriteDataBase ("RR")
                            txtRR.BackColor = &HC000&
                            
                            txtRR.Caption = TireID
                            lbRRPre.Caption = AirPressure & "kPa"
                            lbRRTemp.Caption = Temperature & "℃"
                            lbRRAcSpeed.Caption = Acceleration & "g"
                            lbRRBattery.Caption = BAT
                            lbRRMdl.Caption = State
                            
                            ErrLog.WriteOprInfo "右后轮检测结果:" & TireIDFlag(2)
                            ErrLog.WriteOprInfo "检测完成！"
                            ErrLog.WriteOprInfo "============================="
                            
                            Call SaveToDB
                            
                            AddMessage "右后轮检测完毕", False, False
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 1500
                            
                            resetList False '这个必需放在SaveToDB之后，因为都重置掉了
                            CheckFinish (4)
                            
                            AddMessage "上次检测结果合格", False, False
                            TmrSetColor.Interval = 0
                        Else
                            TireIDFlag(2) = ""
                            TirePreFlag(2) = ""
                            TireTempFlag(2) = ""
                            TireAcSpeedFlag(2) = ""
                            TireBatFlag(2) = ""
                            TireStateFlag(2) = ""
                            
                            AddMessage "请将手持设备靠近右后轮,重新检测", True, True
                            Flash = 4
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                            
                        End If
                End Select
            End If
      Case comEvSend   ' 传输缓冲区有 Sthreshold 个字符                     '
      Case comEvEOF   ' 输入数据流中发现 EOF 字符
   End Select

   Exit Sub
EH: 'error handler
    ErrLog.WriteErrInfo "解析VT60数据", "", "出错！" & Err.Description
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
            Initstallcom
            If MSComDEV1.PortOpen = True Then
                MSComDEV1.PortOpen = False
            End If
            DoEvents
            MSComDEV1.PortOpen = True
        End If

        lblOption.Caption = ""
        AddMessage "设备打开正常,请将手持设置靠近右前轮", False, False
        txtRF.BackColor = &HFF&
        DataFlag = 1
        Flash = 1
        TmrSetColor.Interval = 700
        
        Exit Sub
StartUpERR:
    ErrLog.WriteErrInfo "1号设备StartUp", "", "出错！" & Err.Description
    AddMessage "打开手持设备串口时出错!", True, True
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
'JCAE PF2-PLATFORM ; C9B371 ;    1  kPa ;  20鳦 ; BAT:OK ; DRIVE
On Err GoTo Spliterr
    Dim tmp() As String
    Dim i As Integer
    tmp() = Split(Strtemp, ";")
    TireID = FormatStrLen(Trim(tmp(1)), 8)
    AirPressure = CDbl(Trim(Replace(tmp(2), "kPa", "")))
    Temperature = CInt(CDbl(Trim(Replace(tmp(3), "鳦", ""))) / 10)
    Acceleration = 0
    BAT = Trim(Replace(tmp(4), "BAT:", ""))
    If UBound(tmp()) > 4 Then
        State = Trim(Replace(tmp(5), " ", ""))
        If State = "DRIVE" Then
            State = "4"
        End If
    Else
        State = "L"
    End If
    Exit Sub
Spliterr:
    ErrLog.WriteErrInfo "传感器其他数据处理错误", "", "出错！" & Err.Description
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
                AddMessage "胎压检测完毕", False, False
                ExitFlag = True
            End If
    Else
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Or TireIDFlag(5) = "" Then

            Else
                AddMessage "胎压检测完毕", False, False
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
'Dim lbl(5) As String
'lbl(1) = "右前轮："
'lbl(2) = "左前轮："
'lbl(3) = "左后轮："
'lbl(4) = "右后轮："
'lbl(5) = "备用胎："
'On err GoTo err
'Dim j As Integer
'    j = 1
'    If optTire4.value = True Then
'        DataReport1.Sections(1).Controls("lbl5").Visible = False
'        For j = 1 To 4
'            If TireIDFlag(j) = "" Then
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &HFF&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & "不合格"
'            Else
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &H0&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & TireIDFlag(j)
'            End If
'        Next j
'    Else
'        For j = 1 To 5
'            If TireIDFlag(j) = "" Then
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &HFF&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & "不合格"
'            Else
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &H0&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & TireIDFlag(j)
'            End If
'        Next j
'    End If
'    DataReport1.Sections("section1").Controls("lblVIN").Caption = "VIN:" & Dev1VIN
'    DataReport1.Sections("section1").Controls("lblDate").Caption = "Date:" & Format(Now, "yyyy-mm-dd")
'    DataReport1.Sections("section1").Controls("lblTime").Caption = "Time:" & Format(Now, "hh:mm:ss")
'    DataReport1.PrintReport False, rptRangeFromTo
'    Exit Sub
'err:
'    ErrLog.WriteErrInfo "打印模块", "", "出错！"
'    MsgBox "打印失败！", vbExclamation
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
    On Error GoTo SaveToDBErr
        
    Dim cnn As ADODB.Connection
    Dim rst As ADODB.Recordset
    
    Set cnn = New ADODB.Connection
    Set rst = New ADODB.Recordset
    
    
    On Error GoTo LocalErr
    cnn.ConnectionTimeout = 2
    cnn.CommandTimeout = 2
    cnn.Open LocalDBConnStr
    
    
    
    
    rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
        rst.AddNew
    End If
    rst.Fields("VIN").value = Dev1VIN
    rst("VIS").value = Right(Dev1VIN, 8)

    rst("ID020").value = TireIDFlag(1)
    rst("ID021").value = TireIDFlag(2)
    rst("ID022").value = TireIDFlag(3)
    rst("ID023").value = TireIDFlag(4)
    
    rst("Mdl020").value = TireStateFlag(1)
    rst("Mdl021").value = TireStateFlag(2)
    rst("Mdl022").value = TireStateFlag(3)
    rst("Mdl023").value = TireStateFlag(4)
    
    rst("Pre020").value = TirePreFlag(1)
    rst("Pre021").value = TirePreFlag(2)
    rst("Pre022").value = TirePreFlag(3)
    rst("Pre023").value = TirePreFlag(4)
    
    rst("Temp020").value = TireTempFlag(1)
    rst("Temp021").value = TireTempFlag(2)
    rst("Temp022").value = TireTempFlag(3)
    rst("Temp023").value = TireTempFlag(4)
    
    rst("Battery020").value = TireBatFlag(1)
    rst("Battery021").value = TireBatFlag(2)
    rst("Battery022").value = TireBatFlag(3)
    rst("Battery023").value = TireBatFlag(4)
    
    rst("AcSpeed020").value = TireAcSpeedFlag(1)
    rst("AcSpeed021").value = TireAcSpeedFlag(2)
    rst("AcSpeed022").value = TireAcSpeedFlag(3)
    rst("AcSpeed023").value = TireAcSpeedFlag(4)
    
    rst("TestTime").value = Now
    rst.Fields("TestState") = TestState
    rst.Fields("UploadSign") = False
    rst.Fields("DownloadSign") = False
    rst.Fields("Dev") = DevNum
    rst.Update
    rst.Close
    cnn.Close
        
LocalErr:
           
  Resume SaveToRemoteDB
           
SaveToRemoteDB:

    '远程库存储
    cnn.ConnectionTimeout = 2
    cnn.CommandTimeout = 2
    
    On Error GoTo RemoteErr
    cnn.Open RemoteDBConnStr


    rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
    If rst.EOF Then
        rst.AddNew
    End If
    rst.Fields("VIN").value = Dev1VIN
    rst("VIS").value = Right(Dev1VIN, 8)

    rst("ID020").value = TireIDFlag(1)
    rst("ID021").value = TireIDFlag(2)
    rst("ID022").value = TireIDFlag(3)
    rst("ID023").value = TireIDFlag(4)

    rst("Mdl020").value = TireStateFlag(1)
    rst("Mdl021").value = TireStateFlag(2)
    rst("Mdl022").value = TireStateFlag(3)
    rst("Mdl023").value = TireStateFlag(4)

    rst("Pre020").value = TirePreFlag(1)
    rst("Pre021").value = TirePreFlag(2)
    rst("Pre022").value = TirePreFlag(3)
    rst("Pre023").value = TirePreFlag(4)

    rst("Temp020").value = TireTempFlag(1)
    rst("Temp021").value = TireTempFlag(2)
    rst("Temp022").value = TireTempFlag(3)
    rst("Temp023").value = TireTempFlag(4)

    rst("Battery020").value = TireBatFlag(1)
    rst("Battery021").value = TireBatFlag(2)
    rst("Battery022").value = TireBatFlag(3)
    rst("Battery023").value = TireBatFlag(4)

    rst("AcSpeed020").value = TireAcSpeedFlag(1)
    rst("AcSpeed021").value = TireAcSpeedFlag(2)
    rst("AcSpeed022").value = TireAcSpeedFlag(3)
    rst("AcSpeed023").value = TireAcSpeedFlag(4)

    rst("TestTime").value = Now
    rst.Fields("TestState") = TestState
    rst.Fields("UploadSign") = False
    rst.Fields("DownloadSign") = False
    rst.Fields("Dev") = DevNum
    rst.Update
    rst.Close
    cnn.Close
    
    Exit Sub
    
    
RemoteErr:

    ErrLog.WriteErrInfo "在将检测数据存入到远程数据库时", "", "出错！" & Err.Description
    Exit Sub
    
    
SaveToDBErr:
    ErrLog.WriteErrInfo "1号设备SaveToDB", "", "出错！" & Err.Description
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
    
    Dim cnn As New ADODB.Connection
    Dim rst As New ADODB.Recordset
    Dim TireField As String
    Dim TempField As String
    Dim AcSpeedField As String
    Dim BatteryField As String
    Dim PreField As String
    Dim MdlField As String
    
    Select Case StrPostion
        Case "RF"
            If TireID <> "00000000" Or TireID <> "" Then
                TestState = TestState + 8
            End If
            TireField = "ID020"
            TempField = "Temp020"
            AcSpeedField = "AcSpeed020"
            BatteryField = "Battery020"
            PreField = "Pre020"
            MdlField = "Mdl020"
        Case "LF"
            If TireID <> "00000000" Or TireID <> "" Then
                TestState = TestState + 4
            End If
            TireField = "ID022"
            TempField = "Temp022"
            AcSpeedField = "AcSpeed022"
            BatteryField = "Battery022"
            PreField = "Pre022"
            MdlField = "Mdl022"
        Case "RR"
            If TireID <> "00000000" Or TireID <> "" Then
                TestState = TestState + 2
            End If
            TireField = "ID021"
            TempField = "Temp021"
            AcSpeedField = "AcSpeed021"
            BatteryField = "Battery021"
            PreField = "Pre021"
            MdlField = "Mdl021"
        Case "LR"
            If TireID <> "00000000" Or TireID <> "" Then
                TestState = TestState + 1
            End If
            TempField = "Temp023"
            AcSpeedField = "AcSpeed023"
            BatteryField = "Battery023"
            PreField = "Pre023"
            MdlField = "Mdl023"
            TireField = "ID023"
        Case "ST"
            TireField = "ID024"
    End Select
         
'On Error GoTo Local_Conn
'' 本地库存储
'        cnn.ConnectionString = LocalDBConnStr
'        cnn.ConnectionTimeout = 2
'        cnn.CommandTimeout = 2
'        cnn.Open
'        rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
'        If rst.EOF Then
'            rst.AddNew
'        End If
'        rst.Fields("VIN").value = Dev1VIN
'        rst("VIS").value = Right(Dev1VIN, 8)
'        rst(TireField).value = TireID
'        rst(TempField).value = Temperature
'        rst(AcSpeedField).value = Acceleration
'        rst(BatteryField).value = BAT
'        rst(PreField).value = AirPressure
'        rst(MdlField).value = State
'        rst("TestTime").value = Now
'        rst.Fields("TestState") = TestState
'        rst.Fields("UploadSign") = False
'        rst.Fields("DownloadSign") = False
'        rst.Fields("Dev") = DevNum
'        rst.Update
'        rst.Close
'        cnn.Close
'
'Local_Conn:
'    Resume SaveToRemoteDB
'
'SaveToRemoteDB:
'
'    '远程库存储
'    On Error GoTo Remote_Conn
'        cnn.ConnectionString = RemoteDBConnStr
'        cnn.ConnectionTimeout = 2
'        cnn.CommandTimeout = 2
'        cnn.Open
'        rst.Open "select  * from ""T_Result"" where ""VIN""='" & Dev1VIN & "' ", cnn, adOpenDynamic, adLockOptimistic
'        If rst.EOF Then
'            rst.AddNew
'        End If
'        rst.Fields("VIN").value = Dev1VIN
'        rst("VIS").value = Right(Dev1VIN, 8)
'        rst(TireField).value = TireID
'        rst(TempField).value = Temperature
'        rst(AcSpeedField).value = Acceleration
'        rst(BatteryField).value = BAT
'        rst(PreField).value = AirPressure
'        rst(MdlField).value = State
'        rst("TestTime").value = Now
'        rst.Fields("TestState") = TestState
'        rst.Fields("UploadSign") = False
'        rst.Fields("DownloadSign") = False
'        rst.Fields("Dev") = DevNum
'        rst.Update
'        rst.Close
'        cnn.Close
'
'        Exit Sub
'Remote_Conn:
'    ErrLog.WriteErrInfo "在将检测数据存入到远程数据库时", "", "出错！" & Err.Description
    Exit Sub
WriteDataBaseErr:
    ErrLog.WriteErrInfo "1号设备WriteDataBase", "", "出错！" & Err.Description
End Sub

Private Sub MSScanCom_OnComm()
On Error GoTo Err
    DelayTime 100
    Dim tmp As Variant
    Dim strin As String
    
    tmp = MSScanCom.Input
    If tmp = "" Then Exit Sub
    strin = strin & tmp
    tmp = ""
    txtVIN.Text = strin
    TestCode = strin
    Call txtVIN_KeyPress(13)
    Exit Sub
Err:
    ErrLog.WriteErrInfo "扫描枪扫描条码", "", "出错！" & Err.Description
End Sub

' Timer 检测系统状态
' 如本地数据库硬盘容量，网络状态等
Private Sub TmrCheckState_Timer()
On Error Resume Next

    MM = MM + 1
    If MM < CheckStateInterval Then
        Exit Sub
    End If

    If DataFlag = 0 Then '程序不在测量状态
        '///// 检查本地数据库硬盘容量
        DoEvents
        If GetHDDState(LocalDBDrive, AvailableSpace) = 1 Then 'Normal
            ImgHDDState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        Else
            ImgHDDState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
        End If
        
        '///// 检查网络状态
        If Ping(RemoteServerIP) Then 'Normal
            ImgNetState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        Else
            ImgNetState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
        End If
        
        '///// 检查远端数据库连接状态
        CheckRemoteDbState
        
        '//// 检查本地数据库连接状态
        CheckLocalDbState
    End If
    
    MM = 0
End Sub
'***************************************************************************
' 检查远端数据库连接状态
' 调用位置:TmrCheckState_Timer
'***************************************************************************
Private Sub CheckRemoteDbState()
On Error GoTo CheckRemoteDbStateErr
    Dim objConn As Connection
    
    Set objConn = New Connection
    objConn.ConnectionTimeout = 2
    objConn.Open RemoteDBConnStr
    If objConn.State = adStateOpen Then 'Normal
        ImgRemoteDBState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        objConn.Close
    Else
        ImgRemoteDBState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
    End If

    Set objConn = Nothing
    Exit Sub
CheckRemoteDbStateErr:
    Set objConn = Nothing
    ImgRemoteDBState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
End Sub

'***************************************************************************
' 检查本地数据库连接状态
' 调用位置:TmrCheckState_Timer
'***************************************************************************
Private Sub CheckLocalDbState()
On Error GoTo CheckLocalDbStateErr
    Dim objConn As Connection
    
    Set objConn = New Connection
    objConn.ConnectionTimeout = 2
    objConn.Open LocalDBConnStr
    If objConn.State = adStateOpen Then 'Normal
        ImgLocalDBState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        objConn.Close
    Else
        ImgLocalDBState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
    End If

    Set objConn = Nothing
    Exit Sub
CheckLocalDbStateErr:
    Set objConn = Nothing
    ImgLocalDBState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
End Sub
Private Sub TmrSetColor_Timer()
On Error GoTo Err
    Select Case Flash
        Case 1
            If Flag = 0 Then
                txtRF.BackColor = &HFFFFFF
                Flag = 1
            ElseIf Flag = 1 Then
                txtRF.BackColor = &HFF&
                Flag = 0
            End If
            
            If txtLF.Caption <> "" Then
                txtLF.BackColor = &HFF00& '绿色
            Else
                txtLF.BackColor = &HFFFFFF  '白色
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '绿色
            Else
                txtLR.BackColor = &HFFFFFF  '白色
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '绿色
            Else
                txtRR.BackColor = &HFFFFFF  '白色
            End If
        Case 2
            If Flag = 0 Then
                txtLF.BackColor = &HFFFFFF
                Flag = 1
            ElseIf Flag = 1 Then
                txtLF.BackColor = &HFF&
                Flag = 0
            End If
            
            If txtRF.Caption <> "" Then
                txtRF.BackColor = &HFF00& '绿色
            Else
                txtRF.BackColor = &HFFFFFF  '白色
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '绿色
            Else
                txtLR.BackColor = &HFFFFFF  '白色
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '绿色
            Else
                txtRR.BackColor = &HFFFFFF  '白色
            End If
        Case 3
            If Flag = 0 Then
                txtLR.BackColor = &HFFFFFF
                Flag = 1
            ElseIf Flag = 1 Then
                txtLR.BackColor = &HFF&
                Flag = 0
            End If
            
            If txtLF.Caption <> "" Then
                txtLF.BackColor = &HFF00& '绿色
            Else
                txtLF.BackColor = &HFFFFFF  '白色
            End If
            
            If txtRF.Caption <> "" Then
                txtRF.BackColor = &HFF00& '绿色
            Else
                txtRF.BackColor = &HFFFFFF  '白色
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '绿色
            Else
                txtRR.BackColor = &HFFFFFF  '白色
            End If
        Case 4
            If Flag = 0 Then
                txtRR.BackColor = &HFFFFFF
                Flag = 1
            ElseIf Flag = 1 Then
                txtRR.BackColor = &HFF&
                Flag = 0
            End If
            
            If txtLF.Caption <> "" Then
                txtLF.BackColor = &HFF00& '绿色
            Else
                txtLF.BackColor = &HFFFFFF  '白色
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '绿色
            Else
                txtLR.BackColor = &HFFFFFF  '白色
            End If
                        
            If txtRF.Caption <> "" Then
                txtRF.BackColor = &HFF00& '绿色
            Else
                txtRF.BackColor = &HFFFFFF  '白色
            End If
    End Select
    Exit Sub
Err:
    ErrLog.WriteErrInfo "在设置检测状态颜色时", "", "出错！" & Err.Description
End Sub

Private Sub txtVIN_KeyPress(KeyAscii As Integer)
On Error GoTo Err
If KeyAscii = 13 Then
        TestCode = Replace(txtVIN.Text, Chr(10), "")
        TestCode = Replace(TestCode, Chr(13), "")
        If Trim(txtVIN.Text) = "" Then
            Exit Sub
        End If
        
        If Len(TestCode) <> 17 And Len(TestCode) <> 26 Then
            AddMessage "录入的条码长度不合法！", True, True
            txtVIN.Text = ""
            TestCode = ""
            
            If isFormLoad = False Then
                UseIOPort LampBuzzerPort, 500
            End If
            
            Exit Sub
        Else
            If txtVIN.Text = "R010000000000000C" Then
                resetList True
                ErrLog.WriteOprInfo "扫描复位条码，系统被复位"
            Else
                If Len(TestCode) = 26 Then
                   If StrConv(Right(Left(TestCode, 24), 1), vbUpperCase) <> "D" Then
                        AddMessage "该车辆未装配DSG传感器！", True, True
                        txtVIN.Text = ""
                        TestCode = ""
                        
                        UseIOPort LampBuzzerPort, 500
                        
                        Exit Sub
                   Else
                        TestCode = Right(Left(TestCode, 18), 17)
                        txtVIN.Text = TestCode
                   End If
                End If
                
                txtVIN.SelStart = 0
                txtVIN.SelLength = Len(txtVIN.Text)
                txtVIN.SetFocus
                
                oIOCard.OutputController LampYellowPort, True '黄色表示开始测量
                
                ErrLog.WriteOprInfo "扫描VIN码：" & TestCode
                Dev1VIN = TestCode
                BeginTestFlow TestCode
            End If
        End If
    End If
    Exit Sub
Err:
    ErrLog.WriteErrInfo "在扫描条码时", "", "出错！" & Err.Description
End Sub
'显示系统信息
Public Sub AddMessage(txt As String, Optional isAlert As Boolean = False, Optional isWriteLog As Boolean = False)

    lblOption.Caption = txt
    If isAlert Then
        lblOption.ForeColor = &HFF&
    Else
        lblOption.ForeColor = &H80000002
    End If
    
    If isWriteLog Then
        ErrLog.WriteOprInfo txt
    End If
End Sub

Public Sub resetList(Optional isCloseComPort As Boolean = False)
On Error GoTo Err:
    txtVIN.Text = ""
    lblOption.Caption = "等待扫描VIN码，开始测量!"
    lblOption.ForeColor = &H80000002
    TmrSetColor.Interval = 0
    DataFlag = 0
    Flash = 0
    
    TireIDFlag(1) = ""
    TireIDFlag(2) = ""
    TireIDFlag(3) = ""
    TireIDFlag(4) = ""
    TireIDFlag(5) = ""
    
    TirePreFlag(1) = ""
    TirePreFlag(2) = ""
    TirePreFlag(3) = ""
    TirePreFlag(4) = ""
    TirePreFlag(5) = ""
    
    TireTempFlag(1) = ""
    TireTempFlag(2) = ""
    TireTempFlag(3) = ""
    TireTempFlag(4) = ""
    TireTempFlag(5) = ""
    
    TireAcSpeedFlag(1) = ""
    TireAcSpeedFlag(2) = ""
    TireAcSpeedFlag(3) = ""
    TireAcSpeedFlag(4) = ""
    TireAcSpeedFlag(5) = ""
    
    TireBatFlag(1) = ""
    TireBatFlag(2) = ""
    TireBatFlag(3) = ""
    TireBatFlag(4) = ""
    TireBatFlag(5) = ""
    
    TireStateFlag(1) = ""
    TireStateFlag(2) = ""
    TireStateFlag(3) = ""
    TireStateFlag(4) = ""
    TireStateFlag(5) = ""
    
    TestState = 0
    
    picLF.Picture = LoadPicture(AppPath & "PIC\blue.jpg")
    picLR.Picture = LoadPicture(AppPath & "PIC\blue.jpg")
    picRF.Picture = LoadPicture(AppPath & "PIC\blue.jpg")
    picRR.Picture = LoadPicture(AppPath & "PIC\blue.jpg")

    txtLR.Caption = ""
    lbLRMdl.Caption = ""
    lbLRPre.Caption = ""
    lbLRTemp.Caption = ""
    lbLRBattery.Caption = ""
    lbLRAcSpeed.Caption = ""

    txtLF.Caption = ""
    lbLFMdl.Caption = ""
    lbLFPre.Caption = ""
    lbLFTemp.Caption = ""
    lbLFBattery.Caption = ""
    lbLFAcSpeed.Caption = ""

    txtRR.Caption = ""
    lbRRMdl.Caption = ""
    lbRRPre.Caption = ""
    lbRRTemp.Caption = ""
    lbRRBattery.Caption = ""
    lbRRAcSpeed.Caption = ""

    txtRF.Caption = ""
    lbRFMdl.Caption = ""
    lbRFPre.Caption = ""
    lbRFTemp.Caption = ""
    lbRFBattery.Caption = ""
    lbRFAcSpeed.Caption = ""
    
    txtRF.BackColor = &H80000005
    txtRR.BackColor = &H80000005
    txtLF.BackColor = &H80000005
    txtLR.BackColor = &H80000005
    
    If isCloseComPort Then
        blnOpenstat = False
        If MSComDEV1.PortOpen = True Then
           MSComDEV1.PortOpen = False
        End If
    End If
    
    CloseAllIO '关闭IO输出
    
    Exit Sub
Err:
    ErrLog.WriteErrInfo "在重置系统状态时", "", "出错！" & Err.Description
End Sub

Public Sub BeginTestFlow(VIN As String)
    StartUp
End Sub

Public Sub Initstallcom()
    Dim blueComNo As Integer
    On Error GoTo comerr
    blueComNo = GetIniS("App", "BlueToothComPort", "", GetProjectPath() & "Setting.ini")
    MSComDEV1.CommPort = blueComNo
    MSComDEV1.InBufferSize = 1024
    MSComDEV1.OutBufferSize = 1024
    MSComDEV1.InBufferCount = 0
    MSComDEV1.Settings = "57600,n,8,1"
    MSComDEV1.InputMode = comInputModeText
    MSComDEV1.RTSEnable = True
    MSComDEV1.RThreshold = 1
    strin = ""
    blnOpenstat = True
    
    Exit Sub
comerr:
    blnOpenstat = False
    lblOption.ForeColor = &HFF&
    lblOption.Caption = "配置串口失败"
    ErrLog.WriteErrInfo "手持设备配置串口时出错", "", "出错！"
    AddMessage "手持设备配置串口时出错!", True, True
    txtVIN = ""
End Sub

Public Function FormatStrLen(str As String, strLen As Integer) As String
    Dim i As Integer
    Dim tmpLen As Integer
    
On Error GoTo Err:
    If Len(str) < strLen Then
        tmpLen = strLen - Len(str)
        For i = 1 To tmpLen
            str = "0" & str
        Next i
    ElseIf Len(str) > strLen Then
        str = Mid(str, 1, strLen)
    End If
    
    FormatStrLen = str
    Exit Function
Err:
    FormatStrLen = str
    Exit Function
End Function

Public Function CloseAllIO()
On Error Resume Next
    oIOCard.OutputController LampYellowPort, False '关闭黄灯
    oIOCard.OutputController LampRedPort, False '关闭红灯
    oIOCard.OutputController LampGreenPort, False '关闭绿灯
    oIOCard.OutputController LampBuzzerPort, False '关闭蜂鸣
    oIOCard.OutputController HornPort, False '关闭喇叭
End Function

Public Function UseIOPort(portNum As Integer, keepTime As Long)
On Error Resume Next
    oIOCard.OutputController portNum, True
    DelayTime keepTime
    oIOCard.OutputController portNum, False
End Function

'关闭指定名称的进程
Public Sub KillProcess(sProcess As String)
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
    tPE.dwSize = Len(tPE)
    lNextProcess = Process32First(lSnapShot, tPE)
    Do While lNextProcess
    If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
    Dim lProcess As Long
    Dim lExitCode As Long
    lProcess = OpenProcess(1, False, tPE.th32ProcessID)
    TerminateProcess lProcess, lExitCode
    CloseHandle lProcess
    End If
    lNextProcess = Process32Next(lSnapShot, tPE)
    Loop
    CloseHandle (lSnapShot)
    End If
End Sub
