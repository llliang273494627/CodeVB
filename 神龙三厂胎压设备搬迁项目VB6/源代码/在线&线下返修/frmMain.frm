VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "TPMS���������߷���"
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
   StartUpPosition =   3  '����ȱʡ
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
      Caption         =   "̥ѹ��ʼ������ϵͳ"
      BeginProperty Font 
         Name            =   "������"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǰ�֣�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�Һ��֣�"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
      Caption         =   "ѹ��:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�:"
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
      Left            =   12690
      TabIndex        =   14
      Top             =   4830
      Width           =   1455
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "���:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "��ǰ�֣�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ѹ��:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�¶�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "���ٶ�:"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "����֣�"
      BeginProperty Font 
         Name            =   "����"
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
      Caption         =   "�ȴ�ɨ��VIN�룬��ʼ����!"
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "����"
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
'�ر�ָ������
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

Private RemoteServerIP As String 'Զ�˷�����IP
Private LocalDBDrive As String '�������ݿ����ڴ����̷�
Private CheckStateInterval As Integer '���ϵͳ��״̬��ʱ������
Public DevNum As String '��ϵͳ�ı�ţ����߷���Ϊ201��������Ϊ301
Private MM As Integer
Private TestCode As String
Private isFormLoad As Boolean
Public isQuit As Boolean

'�źŵ���ؿ��Ʋ�����io�ź�����˿ڣ�
Public LampYellowPort As Integer
Public LampRedPort As Integer
Public LampGreenPort As Integer
Public LampBuzzerPort As Integer
Public HornPort As Integer

'********************************************************************************
' ��������
'********************************************************************************
Const AvailableSpace = 100 '�������ݿ����ڴ�����С���ÿռ�(MB)

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
    
    '�����źŵƼ�����IO����˿�
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
    ErrLog.WriteErrInfo "��ɨ��ǹ����", "", "����" & Err.Description
    AddMessage "��ɨ��ǹ����ʱ����!", True, True
    Exit Sub
LoadErr:
    isFormLoad = False
    ErrLog.WriteErrInfo "����ϵͳ��������", "", "����" & Err.Description
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
        
        Call CloseAllIO '�ر�IO���
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
    
    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 3
        AddMessage "���ڼ�������......", False, True
        txtLR.BackColor = &HFF&
        DataFlag = 3
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
    End If
End Sub

Private Sub Image5_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 2
        AddMessage "���ڼ����ǰ��......", False, True
        txtLF.BackColor = &HFF&
        DataFlag = 2
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
    End If
End Sub

Private Sub Image6_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 4
        AddMessage "���ڼ���Һ���......", False, True
        txtRR.BackColor = &HFF&
        DataFlag = 4
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
    End If
End Sub

Private Sub Image7_Click()
    If DataFlag <= 0 Then
        Exit Sub
    End If

    If MsgBox("��ѡ�����ֶ�����", vbYesNo) = vbYes Then
        Flash = 1
        AddMessage "���ڼ����ǰ��......", False, True
        txtRF.BackColor = &HFF&
        DataFlag = 1
        
        CloseAllIO
        oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
    End If
End Sub

Private Sub MSComDEV1_OnComm()
    Dim recv() As Byte
    Dim tmp As Variant
    Dim i As Long
    Dim nStatus As Long, n As Long 'nStatus '���ڽ���״̬��

    On Error GoTo EH

    Static OnCommBusy As Boolean

    DoEvents

    Select Case MSComDEV1.CommEvent
    ' Handle each event or error by placing
    ' code below each case statement

    ' ����
      Case comEventBreak   ' �յ� Break��
        'Debug.Print "�յ��ж�"
        AddMessage "�յ��ж�", True, True
        blnOpenstat = False
        'Unload Me
        Err.Clear
        Exit Sub

      Case comEventCDTO   ' CD (RLSD) ��ʱ��
      Case comEventCTSTO   ' CTS Timeout��
      Case comEventDSRTO   ' DSR Timeout��
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   '���ݶ�ʧ��
      Case comEventRxOver '���ջ����������
        'MSComm1.InBufferCount = 0
        'AddStrToRTB "���ջ�������� !" + Chr(10), RGB(50, 0, 0)
      Case comEventRxParity ' Parity ����
      Case comEventTxFull   '���仺����������
        'MsgBox "���ͻ���������", vbOKOnly, "����"
      Case comEventDCB   '��ȡ DCB] ʱ�������

      '�¼�
      Case comEvCD   ' CD ��״̬�仯��
        If blnOpenstat = False Then    '״̬5
            MSComDEV1.PortOpen = False
        End If
        'MSComm1.PortOpen = False
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

            DelayTime 100

            tmp = MSComDEV1.Input
            strin = strin & tmp
            tmp = ""

            strin = Replace(strin, Chr(10), "")
            strin = Replace(strin, Chr(13), "")

            'ErrLog.WriteOprInfo "���յ�VT60���ص�����:" & strin
            If Right(strin, 3) = ":OK" Or StrConv(Right(strin, 3), vbUpperCase) = "LOW" Or Right(strin, 5) = "DRIVE" Or Right(strin, 5) = "LEARN" Then
            
                Select Case DataFlag
                    Case 1
                    '��ʾ���ڼ��״̬����һ��״̬
                        AddMessage "���ڼ����ǰ��......", False, True
                        ErrLog.WriteOprInfo "���յ�VT60���ص�����:" & strin
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
                        
                        '�ж��Ƿ����ظ�
                        If TireIDFlag(1) <> TireIDFlag(2) And TireIDFlag(1) <> TireIDFlag(3) And TireIDFlag(1) <> TireIDFlag(4) Then
                            WriteDataBase ("RF")
                            DataFlag = DataFlag + 1
                            txtRF.BackColor = &HC000&
                            ErrLog.WriteOprInfo "��ǰ�ּ����:" & TireIDFlag(1)
                            AddMessage "��ǰ�ּ�����,�뽫�豸�Ƶ���ǰ��", False, False
                            txtLF.BackColor = &HFF&
                            
                            txtRF.Caption = TireID
                            lbRFPre.Caption = AirPressure & "kPa"
                            lbRFTemp.Caption = Temperature & "��"
                            lbRFAcSpeed.Caption = Acceleration & "g"
                            lbRFBattery.Caption = BAT
                            lbRFMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                        Else
                            TireIDFlag(1) = ""
                            TirePreFlag(1) = ""
                            TireTempFlag(1) = ""
                            TireAcSpeedFlag(1) = ""
                            TireBatFlag(1) = ""
                            TireStateFlag(1) = ""
                            AddMessage "�뽫�ֳ��豸������ǰ��,���¼��", True, True
                            Flash = 1
                                                        
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                            
                        End If
                    Case 2
                        AddMessage "���ڼ����ǰ��......", False, True
                        ErrLog.WriteOprInfo "���յ�VT60���ص�����:" & strin
                        Flash = 3
                        Derult = ""
                        Derult = strin
                        strin = ""
                        Call SplitData(Derult)
                        Derult = ""
                        TireIDFlag(3) = TireID '&H000000FF& ��ɫ  &H0000FF00& ��ɫ
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
                            ErrLog.WriteOprInfo "��ǰ�ּ����:" & TireIDFlag(3)
                            AddMessage "��ǰ�ּ�����,�뽫�豸�Ƶ������", False, False
                            txtLR.BackColor = &HFF&
                            
                            txtLF.Caption = TireID
                            lbLFPre.Caption = AirPressure & "kPa"
                            lbLFTemp.Caption = Temperature & "��"
                            lbLFAcSpeed.Caption = Acceleration & "g"
                            lbLFBattery.Caption = BAT
                            lbLFMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                        Else
                            TireIDFlag(3) = ""
                            TirePreFlag(3) = ""
                            TireTempFlag(3) = ""
                            TireAcSpeedFlag(3) = ""
                            TireBatFlag(3) = ""
                            TireStateFlag(3) = ""
                            
                            AddMessage "�뽫�ֳ��豸������ǰ��,���¼��", True, True
                            Flash = 2
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                            
                        End If
                    Case 3
                        AddMessage "���ڼ�������......", True
                        ErrLog.WriteOprInfo "���յ�VT60���ص�����:" & strin
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
                        
                    '�ж��Ƿ����ظ�
                        If TireIDFlag(4) <> TireIDFlag(1) And TireIDFlag(4) <> TireIDFlag(2) And TireIDFlag(4) <> TireIDFlag(3) Then
                            WriteDataBase ("LR")
                            DataFlag = DataFlag + 1
                            txtLR.BackColor = &HC000&
                            ErrLog.WriteOprInfo "����ּ����:" & TireIDFlag(4)
                            AddMessage "����ּ�����,�뽫�豸�Ƶ��Һ���", False, False
                            txtRR.BackColor = &HFF&
                        
                            txtLR.Caption = TireID
                            lbLRPre.Caption = AirPressure & "kPa"
                            lbLRTemp.Caption = Temperature & "��"
                            lbLRAcSpeed.Caption = Acceleration & "g"
                            lbLRBattery.Caption = BAT
                            lbLRMdl.Caption = State
                            
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 5000
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                        Else
                            TireIDFlag(4) = ""
                            TirePreFlag(4) = ""
                            TireTempFlag(4) = ""
                            TireAcSpeedFlag(4) = ""
                            TireBatFlag(4) = ""
                            TireStateFlag(4) = ""
                            
                            AddMessage "�뽫�ֳ��豸���������,���¼��", True, False
                            Flash = 3
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                            
                        End If
                    Case 4
                        AddMessage "���ڼ���Һ���......", False, True
                        ErrLog.WriteOprInfo "���յ�VT60���ص�����:" & strin
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
                        
                        '�ж��Ƿ����ظ�
                        If TireIDFlag(2) <> TireIDFlag(1) And TireIDFlag(2) <> TireIDFlag(3) And TireIDFlag(2) <> TireIDFlag(4) Then
                            WriteDataBase ("RR")
                            txtRR.BackColor = &HC000&
                            
                            txtRR.Caption = TireID
                            lbRRPre.Caption = AirPressure & "kPa"
                            lbRRTemp.Caption = Temperature & "��"
                            lbRRAcSpeed.Caption = Acceleration & "g"
                            lbRRBattery.Caption = BAT
                            lbRRMdl.Caption = State
                            
                            ErrLog.WriteOprInfo "�Һ��ּ����:" & TireIDFlag(2)
                            ErrLog.WriteOprInfo "�����ɣ�"
                            ErrLog.WriteOprInfo "============================="
                            
                            Call SaveToDB
                            
                            AddMessage "�Һ��ּ�����", False, False
                            oIOCard.OutputController LampYellowPort, False
                            UseIOPort LampGreenPort, 1500
                            
                            resetList False '����������SaveToDB֮����Ϊ�����õ���
                            CheckFinish (4)
                            
                            AddMessage "�ϴμ�����ϸ�", False, False
                            TmrSetColor.Interval = 0
                        Else
                            TireIDFlag(2) = ""
                            TirePreFlag(2) = ""
                            TireTempFlag(2) = ""
                            TireAcSpeedFlag(2) = ""
                            TireBatFlag(2) = ""
                            TireStateFlag(2) = ""
                            
                            AddMessage "�뽫�ֳ��豸�����Һ���,���¼��", True, True
                            Flash = 4
                            
                            CloseAllIO
                            oIOCard.OutputController LampRedPort, True
                            oIOCard.OutputController HornPort, True
                            DelayTime 500
                            oIOCard.OutputController HornPort, False
                            DelayTime 3500
                            oIOCard.OutputController LampRedPort, False
                            oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                            
                        End If
                End Select
            End If
      Case comEvSend   ' ���仺������ Sthreshold ���ַ�                     '
      Case comEvEOF   ' �����������з��� EOF �ַ�
   End Select

   Exit Sub
EH: 'error handler
    ErrLog.WriteErrInfo "����VT60����", "", "����" & Err.Description
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
        AddMessage "�豸������,�뽫�ֳ����ÿ�����ǰ��", False, False
        txtRF.BackColor = &HFF&
        DataFlag = 1
        Flash = 1
        TmrSetColor.Interval = 700
        
        Exit Sub
StartUpERR:
    ErrLog.WriteErrInfo "1���豸StartUp", "", "����" & Err.Description
    AddMessage "���ֳ��豸����ʱ����!", True, True
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
'JCAE PF2-PLATFORM ; C9B371 ;    1  kPa ;  20�C ; BAT:OK ; DRIVE
On Err GoTo Spliterr
    Dim tmp() As String
    Dim i As Integer
    tmp() = Split(Strtemp, ";")
    TireID = FormatStrLen(Trim(tmp(1)), 8)
    AirPressure = CDbl(Trim(Replace(tmp(2), "kPa", "")))
    Temperature = CInt(CDbl(Trim(Replace(tmp(3), "�C", ""))) / 10)
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
    ErrLog.WriteErrInfo "�������������ݴ������", "", "����" & Err.Description
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
                AddMessage "̥ѹ������", False, False
                ExitFlag = True
            End If
    Else
            If TireIDFlag(1) = "" Or TireIDFlag(2) = "" Or TireIDFlag(3) = "" Or TireIDFlag(4) = "" Or TireIDFlag(5) = "" Then

            Else
                AddMessage "̥ѹ������", False, False
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
'Dim lbl(5) As String
'lbl(1) = "��ǰ�֣�"
'lbl(2) = "��ǰ�֣�"
'lbl(3) = "����֣�"
'lbl(4) = "�Һ��֣�"
'lbl(5) = "����̥��"
'On err GoTo err
'Dim j As Integer
'    j = 1
'    If optTire4.value = True Then
'        DataReport1.Sections(1).Controls("lbl5").Visible = False
'        For j = 1 To 4
'            If TireIDFlag(j) = "" Then
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &HFF&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & "���ϸ�"
'            Else
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &H0&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & TireIDFlag(j)
'            End If
'        Next j
'    Else
'        For j = 1 To 5
'            If TireIDFlag(j) = "" Then
'                DataReport1.Sections("section1").Controls("lbl" & j).ForeColor = &HFF&
'                DataReport1.Sections("section1").Controls("lbl" & j).Caption = lbl(j) & "���ϸ�"
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
'    ErrLog.WriteErrInfo "��ӡģ��", "", "����"
'    MsgBox "��ӡʧ�ܣ�", vbExclamation
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

    'Զ�̿�洢
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

    ErrLog.WriteErrInfo "�ڽ�������ݴ��뵽Զ�����ݿ�ʱ", "", "����" & Err.Description
    Exit Sub
    
    
SaveToDBErr:
    ErrLog.WriteErrInfo "1���豸SaveToDB", "", "����" & Err.Description
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
'' ���ؿ�洢
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
'    'Զ�̿�洢
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
'    ErrLog.WriteErrInfo "�ڽ�������ݴ��뵽Զ�����ݿ�ʱ", "", "����" & Err.Description
    Exit Sub
WriteDataBaseErr:
    ErrLog.WriteErrInfo "1���豸WriteDataBase", "", "����" & Err.Description
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
    ErrLog.WriteErrInfo "ɨ��ǹɨ������", "", "����" & Err.Description
End Sub

' Timer ���ϵͳ״̬
' �籾�����ݿ�Ӳ������������״̬��
Private Sub TmrCheckState_Timer()
On Error Resume Next

    MM = MM + 1
    If MM < CheckStateInterval Then
        Exit Sub
    End If

    If DataFlag = 0 Then '�����ڲ���״̬
        '///// ��鱾�����ݿ�Ӳ������
        DoEvents
        If GetHDDState(LocalDBDrive, AvailableSpace) = 1 Then 'Normal
            ImgHDDState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        Else
            ImgHDDState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
        End If
        
        '///// �������״̬
        If Ping(RemoteServerIP) Then 'Normal
            ImgNetState.Picture = LoadPicture(AppPath & "PIC\green.jpg")
        Else
            ImgNetState.Picture = LoadPicture(AppPath & "PIC\red.jpg")
        End If
        
        '///// ���Զ�����ݿ�����״̬
        CheckRemoteDbState
        
        '//// ��鱾�����ݿ�����״̬
        CheckLocalDbState
    End If
    
    MM = 0
End Sub
'***************************************************************************
' ���Զ�����ݿ�����״̬
' ����λ��:TmrCheckState_Timer
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
' ��鱾�����ݿ�����״̬
' ����λ��:TmrCheckState_Timer
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
                txtLF.BackColor = &HFF00& '��ɫ
            Else
                txtLF.BackColor = &HFFFFFF  '��ɫ
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '��ɫ
            Else
                txtLR.BackColor = &HFFFFFF  '��ɫ
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '��ɫ
            Else
                txtRR.BackColor = &HFFFFFF  '��ɫ
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
                txtRF.BackColor = &HFF00& '��ɫ
            Else
                txtRF.BackColor = &HFFFFFF  '��ɫ
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '��ɫ
            Else
                txtLR.BackColor = &HFFFFFF  '��ɫ
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '��ɫ
            Else
                txtRR.BackColor = &HFFFFFF  '��ɫ
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
                txtLF.BackColor = &HFF00& '��ɫ
            Else
                txtLF.BackColor = &HFFFFFF  '��ɫ
            End If
            
            If txtRF.Caption <> "" Then
                txtRF.BackColor = &HFF00& '��ɫ
            Else
                txtRF.BackColor = &HFFFFFF  '��ɫ
            End If
                        
            If txtRR.Caption <> "" Then
                txtRR.BackColor = &HFF00& '��ɫ
            Else
                txtRR.BackColor = &HFFFFFF  '��ɫ
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
                txtLF.BackColor = &HFF00& '��ɫ
            Else
                txtLF.BackColor = &HFFFFFF  '��ɫ
            End If
            
            If txtLR.Caption <> "" Then
                txtLR.BackColor = &HFF00& '��ɫ
            Else
                txtLR.BackColor = &HFFFFFF  '��ɫ
            End If
                        
            If txtRF.Caption <> "" Then
                txtRF.BackColor = &HFF00& '��ɫ
            Else
                txtRF.BackColor = &HFFFFFF  '��ɫ
            End If
    End Select
    Exit Sub
Err:
    ErrLog.WriteErrInfo "�����ü��״̬��ɫʱ", "", "����" & Err.Description
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
            AddMessage "¼������볤�Ȳ��Ϸ���", True, True
            txtVIN.Text = ""
            TestCode = ""
            
            If isFormLoad = False Then
                UseIOPort LampBuzzerPort, 500
            End If
            
            Exit Sub
        Else
            If txtVIN.Text = "R010000000000000C" Then
                resetList True
                ErrLog.WriteOprInfo "ɨ�踴λ���룬ϵͳ����λ"
            Else
                If Len(TestCode) = 26 Then
                   If StrConv(Right(Left(TestCode, 24), 1), vbUpperCase) <> "D" Then
                        AddMessage "�ó���δװ��DSG��������", True, True
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
                
                oIOCard.OutputController LampYellowPort, True '��ɫ��ʾ��ʼ����
                
                ErrLog.WriteOprInfo "ɨ��VIN�룺" & TestCode
                Dev1VIN = TestCode
                BeginTestFlow TestCode
            End If
        End If
    End If
    Exit Sub
Err:
    ErrLog.WriteErrInfo "��ɨ������ʱ", "", "����" & Err.Description
End Sub
'��ʾϵͳ��Ϣ
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
    lblOption.Caption = "�ȴ�ɨ��VIN�룬��ʼ����!"
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
    
    CloseAllIO '�ر�IO���
    
    Exit Sub
Err:
    ErrLog.WriteErrInfo "������ϵͳ״̬ʱ", "", "����" & Err.Description
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
    lblOption.Caption = "���ô���ʧ��"
    ErrLog.WriteErrInfo "�ֳ��豸���ô���ʱ����", "", "����"
    AddMessage "�ֳ��豸���ô���ʱ����!", True, True
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
    oIOCard.OutputController LampYellowPort, False '�رջƵ�
    oIOCard.OutputController LampRedPort, False '�رպ��
    oIOCard.OutputController LampGreenPort, False '�ر��̵�
    oIOCard.OutputController LampBuzzerPort, False '�رշ���
    oIOCard.OutputController HornPort, False '�ر�����
End Function

Public Function UseIOPort(portNum As Integer, keepTime As Long)
On Error Resume Next
    oIOCard.OutputController portNum, True
    DelayTime keepTime
    oIOCard.OutputController portNum, False
End Function

'�ر�ָ�����ƵĽ���
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
