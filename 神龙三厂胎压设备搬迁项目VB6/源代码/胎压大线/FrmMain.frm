VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "̥ѹ����ʼ��ϵͳ"
   ClientHeight    =   11520
   ClientLeft      =   1845
   ClientTop       =   1470
   ClientWidth     =   15360
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmMain.frx":1CFA
   ScaleHeight     =   12214.47
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command12 
      Caption         =   "ʮ������ת����"
      Height          =   675
      Left            =   7260
      TabIndex        =   96
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Command7"
      Height          =   495
      Left            =   9840
      TabIndex        =   95
      Top             =   7440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DoEvents"
      Height          =   495
      Left            =   8160
      TabIndex        =   94
      Top             =   7440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�������빤λ"
      Height          =   405
      Left            =   3300
      TabIndex        =   93
      Top             =   4200
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   360
      Picture         =   "FrmMain.frx":40DF9
      ScaleHeight     =   420
      ScaleWidth      =   645
      TabIndex        =   92
      Top             =   11040
      Width           =   645
   End
   Begin VB.Timer Timer_PrintError 
      Enabled         =   0   'False
      Left            =   1860
      Top             =   4380
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����VT520����"
      Height          =   435
      Left            =   1740
      TabIndex        =   90
      Top             =   9930
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Timer Timer_DataSync 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   2400
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ϵͳ��������"
      Height          =   405
      Left            =   3300
      TabIndex        =   49
      Top             =   3600
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command6 
      Caption         =   "��������"
      Height          =   465
      Left            =   1740
      TabIndex        =   48
      Top             =   9390
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton Command5 
      Caption         =   "�������"
      Height          =   465
      Left            =   60
      TabIndex        =   47
      Top             =   9915
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton Command11 
      Caption         =   "�����"
      Height          =   405
      Left            =   3300
      TabIndex        =   46
      Top             =   6120
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command10 
      Caption         =   "�Һ���"
      Height          =   405
      Left            =   3300
      TabIndex        =   45
      Top             =   5640
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command9 
      Caption         =   "��ǰ��"
      Height          =   405
      Left            =   3300
      TabIndex        =   44
      Top             =   5160
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.CommandButton Command8 
      Caption         =   "��ǰ��"
      Height          =   405
      Left            =   3300
      TabIndex        =   43
      Top             =   4680
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.TextBox txtInputVIN 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   435
      Left            =   0
      TabIndex        =   42
      Text            =   "�ֹ�¼��VIN���س�ȷ��"
      Top             =   1140
      Width           =   3345
   End
   Begin VB.CommandButton Command14 
      Caption         =   "�������"
      Height          =   495
      Left            =   7830
      TabIndex        =   41
      Top             =   2730
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command17 
      Caption         =   "ɨ��̥ѹ��"
      Height          =   495
      Left            =   7830
      TabIndex        =   40
      Top             =   2190
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7800
      TabIndex        =   39
      Text            =   "LMGDK1G87B1S00037"
      Top             =   1740
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.ListBox List1 
      Height          =   2760
      Left            =   12720
      TabIndex        =   38
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtVin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H007B3C08&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   2580
      Locked          =   -1  'True
      MaxLength       =   17
      TabIndex        =   37
      Top             =   1140
      Width           =   12735
   End
   Begin VB.Timer Timer_StatusQuery 
      Interval        =   1000
      Left            =   1800
      Top             =   1890
   End
   Begin VB.ListBox ListMsg 
      Height          =   1500
      Left            =   3900
      TabIndex        =   29
      Top             =   9150
      Width           =   11055
   End
   Begin VB.TextBox txtRF 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12390
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   7410
      Width           =   2235
   End
   Begin VB.PictureBox picRF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   11460
      Picture         =   "FrmMain.frx":419B4
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   25
      Top             =   7410
      Width           =   420
   End
   Begin VB.TextBox txtRR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   7410
      Width           =   2235
   End
   Begin VB.PictureBox picRR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4110
      Picture         =   "FrmMain.frx":4402C
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   22
      Top             =   7410
      Width           =   420
   End
   Begin VB.TextBox txtLF 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   12360
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   2340
      Width           =   2235
   End
   Begin VB.PictureBox picLF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   11490
      Picture         =   "FrmMain.frx":466A4
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   19
      Top             =   2340
      Width           =   420
   End
   Begin VB.TextBox txtLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   2340
      Width           =   2235
   End
   Begin VB.PictureBox picLR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4110
      Picture         =   "FrmMain.frx":48D1C
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   16
      Top             =   2340
      Width           =   420
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4B394
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   15
      Top             =   8520
      Width           =   420
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4BA8E
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   14
      Top             =   7470
      Width           =   420
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4C188
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   13
      Top             =   6300
      Width           =   420
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4C882
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   12
      Top             =   5070
      Width           =   420
   End
   Begin VB.PictureBox PicNet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4CF7C
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   11
      Top             =   3870
      Width           =   420
   End
   Begin VB.PictureBox PicInd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   570
      Picture         =   "FrmMain.frx":4D676
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   10
      Top             =   2760
      Width           =   420
   End
   Begin VB.PictureBox picCommandReset 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   580
      Left            =   9570
      Picture         =   "FrmMain.frx":4DD70
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   7
      Top             =   495
      Width           =   1565
   End
   Begin VB.PictureBox picCommandConifg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   580
      Left            =   8010
      Picture         =   "FrmMain.frx":4F322
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   6
      Top             =   495
      Width           =   1560
   End
   Begin VB.PictureBox picCommandOut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   580
      Left            =   6450
      Picture         =   "FrmMain.frx":509DA
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   5
      Top             =   495
      Width           =   1565
   End
   Begin VB.PictureBox picCommandLog 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   580
      Left            =   4890
      Picture         =   "FrmMain.frx":51EBC
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   4
      Top             =   495
      Width           =   1565
   End
   Begin VB.PictureBox picCommandHis 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   580
      Left            =   3315
      Picture         =   "FrmMain.frx":533DA
      ScaleHeight     =   585
      ScaleWidth      =   1560
      TabIndex        =   3
      Top             =   495
      Width           =   1565
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   14250
      Picture         =   "FrmMain.frx":5493B
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   2
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox picExit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   14745
      Picture         =   "FrmMain.frx":54DC6
      ScaleHeight     =   360
      ScaleWidth      =   495
      TabIndex        =   1
      Top             =   0
      Width           =   495
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   0
      Picture         =   "FrmMain.frx":553E1
      ScaleHeight     =   450
      ScaleWidth      =   4485
      TabIndex        =   0
      Top             =   30
      Width           =   4485
   End
   Begin MSCommLib.MSComm MSComVIN 
      Left            =   2340
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSCommBT 
      Left            =   3000
      Top             =   1860
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "������"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   1080
      TabIndex        =   91
      Top             =   11100
      Width           =   1410
   End
   Begin VB.Label lbRFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   13110
      TabIndex        =   89
      Top             =   8130
      Width           =   1410
   End
   Begin VB.Label lbRFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11970
      TabIndex        =   88
      Top             =   8130
      Width           =   510
   End
   Begin VB.Label lbRFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   11970
      TabIndex        =   87
      Top             =   7890
      Width           =   540
   End
   Begin VB.Label lbRFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12900
      TabIndex        =   86
      Top             =   7890
      Width           =   930
   End
   Begin VB.Label lbRFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   14250
      TabIndex        =   85
      Top             =   7890
      Width           =   930
   End
   Begin VB.Label lbRRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6900
      TabIndex        =   84
      Top             =   7860
      Width           =   930
   End
   Begin VB.Label lbRRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5550
      TabIndex        =   83
      Top             =   7860
      Width           =   930
   End
   Begin VB.Label lbRRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4590
      TabIndex        =   82
      Top             =   7860
      Width           =   540
   End
   Begin VB.Label lbRRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4590
      TabIndex        =   81
      Top             =   8100
      Width           =   510
   End
   Begin VB.Label lbRRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5760
      TabIndex        =   80
      Top             =   8100
      Width           =   1410
   End
   Begin VB.Label lbLFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   14280
      TabIndex        =   79
      Top             =   2790
      Width           =   930
   End
   Begin VB.Label lbLFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12930
      TabIndex        =   78
      Top             =   2790
      Width           =   930
   End
   Begin VB.Label lbLFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12000
      TabIndex        =   77
      Top             =   2790
      Width           =   540
   End
   Begin VB.Label lbLFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   12000
      TabIndex        =   76
      Top             =   3030
      Width           =   510
   End
   Begin VB.Label lbLFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   13140
      TabIndex        =   75
      Top             =   3030
      Width           =   1410
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4080
      TabIndex        =   74
      Top             =   2790
      Width           =   720
   End
   Begin VB.Label lbLRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5790
      TabIndex        =   73
      Top             =   3030
      Width           =   1410
   End
   Begin VB.Label lbLRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4650
      TabIndex        =   72
      Top             =   3030
      Width           =   510
   End
   Begin VB.Label lbLRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   4650
      TabIndex        =   71
      Top             =   2790
      Width           =   540
   End
   Begin VB.Label lbLRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   5580
      TabIndex        =   70
      Top             =   2790
      Width           =   930
   End
   Begin VB.Label lbLRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6930
      TabIndex        =   69
      Top             =   2790
      Width           =   930
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11400
      TabIndex        =   68
      Top             =   7890
      Width           =   1140
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "ѹ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   12330
      TabIndex        =   67
      Top             =   7890
      Width           =   1200
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "�¶ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13680
      TabIndex        =   66
      Top             =   7890
      Width           =   1200
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "���ٶȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   12330
      TabIndex        =   65
      Top             =   8130
      Width           =   1200
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "��أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11400
      TabIndex        =   64
      Top             =   8130
      Width           =   1200
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4050
      TabIndex        =   63
      Top             =   7860
      Width           =   1140
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "ѹ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4980
      TabIndex        =   62
      Top             =   7860
      Width           =   1200
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "�¶ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6330
      TabIndex        =   61
      Top             =   7860
      Width           =   1200
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "���ٶȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4980
      TabIndex        =   60
      Top             =   8100
      Width           =   1200
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "��أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4050
      TabIndex        =   59
      Top             =   8100
      Width           =   1200
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "ģʽ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11430
      TabIndex        =   58
      Top             =   2790
      Width           =   1140
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "ѹ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   12360
      TabIndex        =   57
      Top             =   2790
      Width           =   1200
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "�¶ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   13680
      TabIndex        =   56
      Top             =   2790
      Width           =   1200
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "���ٶȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   12360
      TabIndex        =   55
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "��أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11430
      TabIndex        =   54
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "��أ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4080
      TabIndex        =   53
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "���ٶȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4980
      TabIndex        =   52
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "�¶ȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   6330
      TabIndex        =   51
      Top             =   2790
      Width           =   690
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ѹ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   4980
      TabIndex        =   50
      Top             =   2790
      Width           =   1200
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   2475
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   327682
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   2400
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   28
      ImageHeight     =   28
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5B0D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5BA5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5C3DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5CD5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5D6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmMain.frx":5E062
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   "��������״̬�쳣"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Left            =   720
      TabIndex        =   36
      Top             =   9000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "�Ҳ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   375
      Left            =   1065
      TabIndex        =   35
      Top             =   8595
      Width           =   1575
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   375
      Left            =   1065
      TabIndex        =   34
      Top             =   7545
      Width           =   1575
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "���ݿ�Ӳ������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   360
      Left            =   1065
      TabIndex        =   33
      Top             =   6375
      Width           =   2175
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SPPV���ݿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   375
      Left            =   1065
      TabIndex        =   32
      Top             =   5130
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "�������ݿ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   375
      Left            =   1065
      TabIndex        =   31
      Top             =   3945
      Width           =   1575
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004E4E4E&
      Height          =   360
      Left            =   1065
      TabIndex        =   30
      Top             =   2835
      Width           =   1335
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "�人�к�ɽ����褶�·��԰·��ȹ���A��2318��    �绰��027-87775236"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7050
      TabIndex        =   28
      Top             =   11130
      Width           =   7875
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   12015
      TabIndex        =   27
      Top             =   6960
      Width           =   2520
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�Һ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   4500
      TabIndex        =   24
      Top             =   6960
      Width           =   2520
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "��ǰ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   11985
      TabIndex        =   21
      Top             =   1890
      Width           =   2520
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   4500
      TabIndex        =   18
      Top             =   1890
      Width           =   2520
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "״̬����:"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   210
      TabIndex        =   9
      Top             =   1980
      Width           =   2175
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "̥ѹ��ʼ��ϵͳ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   660
      Width           =   2805
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'** �ļ�����FrmMain.frm
'** ��  Ȩ��CopyRight (c)
'** �����ˣ�yangshuai
'** ��  �䣺shuaigoplay@live.cn
'** ��  �ڣ�2009-2-27
'** �޸��ˣ�
'** ��  �ڣ�
'** ��  ����DSG��̥���������ϵͳ������
'** ��  ����1.0
'******************************************************************************

Option Explicit

Dim tmpTime As String
'[2011-7-12 16:54:02] osensor0 - ---True
'[2011-7-12 16:54:10] osensor1 - ---True
Dim Step1Time As Integer
'[2011-7-12 16:54:26] osensor2 - ---True
'[2011-7-12 16:54:28] osensor3 - ---True
'[2011-7-12 16:54:35] osensor4 - ---True
'[2011-7-12 16:54:37] osensor2 - ---False
Dim Step2Time As Integer
'[2011-7-12 16:54:52] osensor5 - ---True
'[2011-7-12 16:55:03] osensor5 - ---False
Dim Step3Time As Integer
'[2011-7-12 16:55:23] osensor2 - ---True
'[2011-7-12 16:55:34] osensor2 - ---False
Dim Step4Time As Integer
'[2011-7-12 16:55:39] osensor0 - ---False
'[2011-7-12 16:55:47] osensor1 - ---False
'[2011-7-12 16:55:48] osensor5 - ---True
'[2011-7-12 16:55:59] osensor5 - ---False
'[2011-7-12 16:56:05] osensor3 - ---False
'[2011-7-12 16:56:12] osensor4 - ---False
Dim osen0Time As String

Private WithEvents osensor0  As CSensor
Attribute osensor0.VB_VarHelpID = -1
Private WithEvents osensor1  As CSensor
Attribute osensor1.VB_VarHelpID = -1
Private WithEvents osensor2  As CSensor
Attribute osensor2.VB_VarHelpID = -1
Private WithEvents osensor3  As CSensor
Attribute osensor3.VB_VarHelpID = -1
Private WithEvents osensor4  As CSensor
Attribute osensor4.VB_VarHelpID = -1
Private WithEvents osensor5  As CSensor
Attribute osensor5.VB_VarHelpID = -1
Private WithEvents oRDCommand As CSensor
Attribute oRDCommand.VB_VarHelpID = -1

'����״̬
Private gCancel As Boolean
Dim nn As Integer   '��չʱ�Ӽ���
Dim mm As Integer   '��չʱ�Ӽ���
Dim HH As Integer   '��չʱ�Ӽ���
Public TimerN As Integer    '�Ų�����ͬ������
Public TimerStatus As Integer    '״̬�������

'״̬����
Public DBPosition As String     '���ݿ�洢���̷�
Public SpaceAvailable As Long       '���ÿռ�澯��ֵ


Private firstFlag As Boolean
Private secondFlag As Boolean

Private WithEvents osensorCommand  As CSensor
Attribute osensorCommand.VB_VarHelpID = -1
Private WithEvents osensorLine  As CSensor
Attribute osensorLine.VB_VarHelpID = -1
Private car As CCar
Private TestCode As String
Private VINCode As String
Public MTOCCode As String
Dim inputCode As Dictionary '����洢����
Public TestStateFlag As Integer
Dim barCodeFlag As Boolean
Dim sensorFlag As Boolean
Dim sensorControlFlag As Boolean
Dim testEndDelyed As Boolean
Dim isInTesting As Boolean '�Ƿ����ڼ����̥������ Add by ZCJ 2012-07-09

'TestStateFlag��ʶ�÷���
'-1=��ʾ5�ڱ���ɹ����3���֣�ǰ���ǲ�����û��ɨ�������룬ɨ���״̬����0
'0=vin�Ѿ�������Խ���׼��DSG���
'1=��ǰ�ֲ����ɹ�
'2=��ǰ�ֲ����ɹ�
'3=�Һ��ֲ����ɹ�
'4=����ֲ����ɹ�
'5=����ɹ�
'9998=δװ��DSG
'9999=�ȴ�����

Public BreakFlag As Boolean
'BreakFlag = False  'ϵͳ������������ϵͳ��������
'sensorFlag = True  '��������
'barCodeFlag = True '�൱��ɨ��ǿ��¼������

'����VT520�������
Private Sub Command1_Click()
    Dim tmp As String
    tmp = "FF 03 1A 00 00 01 00 01 00 00 00 E8 03 00 00 E0 2E 17 00 00 00 00 00 3F CC 47 0D 42 41 47 43"

    Dim m_TirePreResult As String

    m_TirePreResult = CLng("&H000003E8") / 300

    Dim Temp As String
    Temp = CLng("&H0017")
    Temp = Val("&H46")

End Sub

Private Sub Command12_Click()
'   Dim A As Integer
'   A = CLong("&H8H")
End Sub

'�������
Private Sub Command14_Click()
    'Call DSGTestEnd
    Dim mtoc As String
    Dim tmpCar As CCar
    Set tmpCar = New CCar
    'mtoc = tmpCar.GetMtocFromVinColl("11")
    tmpCar.VINCode = "11"
    tmpCar.Save
End Sub
'ɨ������
Private Sub Command17_Click()
    BreakFlag = False
    TestCode = Text2.text
    If Left(TestCode, 17) = "R010000000000000C" Then '��������
        LogWritter "0ɨ����������"
        resetList
        Exit Sub
    End If
    If Left(TestCode, 17) = "R020000000000000C" Then 'ǿ����������
        LogWritter "ɨ��ǿ����������"
        barCodeFlag = True
        Exit Sub
    End If
    Debug.Print TestCode
    Call txtVIN_KeyPress(13)
End Sub
'�������빤λ
Private Sub Command2_Click()
    If inputCode.Count <> 0 Then
    '�ٴ�����DSGStart
        Call Me.DSGTestStart(CStr(inputCode(inputCode.Keys(0))))
    End If
End Sub

'ϵͳ����
Private Sub Command3_Click()

If BreakFlag Then
    osensorCommand_onChange True    'ϵͳ����
Else
    osensorCommand_onChange False   '����ϵͳ
End If
'    Dim Result As Boolean
'    Dim arr() As String
'    arr = Split(mdlValue, ",")
'    Result = judgeMdlIsOK("1", arr)
End Sub

Private Sub Command4_Click()

'    oRVT520.ResetResult
'    oRVT520.Start "Comm"
'
'    For i = 0 To 60
'        oRVT520.ReadResult
'        tmpID = oRVT520.TireIDResult
'        If tmpID <> "00000000" And Trim(tmpID) <> "" Then
'            Exit For
'        End If
'    Next i
    
End Sub

'�������Ų����У��൱��ɨ��ǿ��¼������
Private Sub Command5_Click()
    barCodeFlag = True
End Sub
'����������
Private Sub Command6_Click()
    sensorControlFlag = False
End Sub

Private Sub Command7_Click()
    Dim A As Integer
    Do While A < 10000
        A = A + 1
    Loop
    DelayTime 2000
    Do While A < 10000
        A = A + 1
    Loop
End Sub

'��ǰ��(����ʱ��)
Private Sub Command8_Click()

'    If DateDiff("s", tmpTime, Now) <= Step1Time Then
'        MsgBox ("��Ӧʱ��δ�ﵽҪ��!")
'        Exit Sub
'    Else
'        tmpTime = Now
'    End If

    'BreakFlag = False  'ϵͳ����
    'sensorFlag = True  '��������
    TestStateFlag = 0
    Dim tmpID As String
    Dim i As Long
    If TestStateFlag = 0 Then
        '�������̣����빤λ
        '�����ǰ��

        TestStateFlag = 1
        updateState "state", CStr(TestStateFlag)
        isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ�����ǰ��
        AddMessage "���ڼ����ǰ�֡���"
        LogWritter "��ʼ��һ�μ����ǰ�֡���"
        oRVT520.ResetResult
        oRVT520.Start "Comm"

'        For i = 0 To 20
'            oRVT520.ReadResult
'            tmpID = oRVT520.TireIDResult
'            If tmpID <> "00000000" And Trim(tmpID) <> "" Then
'                Exit For
'            End If
'        Next i
'        If tmpID = "00000000" Or Trim(tmpID) = "" Then '�ұ�û�в⵽�ز�һ��
'            LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
'            oRVT520.ResetResult
'            oRVT520.Start "Comm"
'            For i = 0 To 20
'                oRVT520.ReadResult
'                tmpID = oRVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
'                    Exit For
'                End If
'            Next i
'        End If

        For i = 0 To 6
            oRVT520.ReadResult
            tmpID = oRVT520.TireIDResult
            If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                Exit For
            End If
        Next i
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Then '�ڶ��β���
            LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                    Exit For
                End If
            Next i
        End If

        If tmpID = "00000000" Or Trim(tmpID) = "" Then '�����β���
            LogWritter "��ʼ�����μ����ǰ�֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Then '���Ĵβ���
            LogWritter "��ʼ���Ĵμ����ǰ�֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                    Exit For
                End If
            Next i
        End If

        If tmpID = "00000000" Or Trim(tmpID) = "" Then '����β���
            LogWritter "��ʼ����μ����ǰ�֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                    Exit For
                End If
            Next i
        End If
        
        isInTesting = False 'Add by ZCJ 2012-07-09 ��ǰ�ּ�����

        car.TireRFID = tmpID
        LogWritter "��ǰ�ּ�����ݣ�" & oRVT520.Result
        car.TireRFMdl = oRVT520.TireMdlResult
        car.TireRFPre = oRVT520.TirePreResult
        car.TireRFTemp = oRVT520.TireTempResult
        car.TireRFBattery = oRVT520.TireBatteryResult
        car.TireRFAcSpeed = oRVT520.TireAcSpeedResult

        updateState "dsgrf", tmpID
        updateState "mdlrf", car.TireRFMdl
        updateState "prerf", car.TireRFPre
        updateState "temprf", car.TireRFTemp
        updateState "batteryrf", car.TireRFBattery
        updateState "acspeedrf", car.TireRFAcSpeed

        '��ǰ�ּ�����
        setFrm TestStateFlag
    End If
End Sub
'��ǰ��(����ʱ��)
Private Sub Command9_Click()

'    If DateDiff("s", tmpTime, Now) <= Step2Time Then
'        MsgBox ("��Ӧʱ��δ�ﵽҪ��!")
'        Exit Sub
'    Else
'        tmpTime = Now
'    End If

    TestStateFlag = 1
    Dim tmpID As String
    Dim i As Long

    If TestStateFlag = 1 Then
        TestStateFlag = 2
        updateState "state", CStr(TestStateFlag)
        isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ�����ǰ��
        AddMessage "���ڼ����ǰ�֡���"
        LogWritter "��ʼ��һ�μ����ǰ�֡���"
        oLVT520.ResetResult
        oLVT520.Start "Comm"
        
'        For i = 0 To 40
'            oLVT520.ReadResult
'            tmpID = oLVT520.TireIDResult
'            If tmpID <> "00000000" And Trim(tmpID) <> "" Then
'                Exit For
'            End If
'        Next i
'        If tmpID = "00000000" Or Trim(tmpID) = "" Then '���û�в⵽�ز�һ��
'            LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
'            oLVT520.ResetResult
'            oLVT520.Start "Comm"
'            For i = 0 To 40
'                oLVT520.ReadResult
'                tmpID = oLVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
'                    Exit For
'                End If
'            Next i
'        End If

        For i = 0 To 6
            oLVT520.ReadResult
            tmpID = oLVT520.TireIDResult
            'If tmpID <> "00000000" And Trim(tmpID) <> "" Then
            If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                Exit For
            End If
        Next i
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '�ڶ��β���
            LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '�����β���
            LogWritter "��ʼ�����μ����ǰ�֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '���Ĵβ���
            LogWritter "��ʼ���Ĵμ����ǰ�֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '����β���
            LogWritter "��ʼ����μ����ǰ�֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                    Exit For
                End If
            Next i
        End If
        
        isInTesting = False 'Add by ZCJ 2012-07-09 ��ǰ�ּ�����

        car.TireLFID = tmpID
        LogWritter "��ǰ�ּ�����ݣ�" & oLVT520.Result
        car.TireLFMdl = oLVT520.TireMdlResult
        car.TireLFPre = oLVT520.TirePreResult
        car.TireLFTemp = oLVT520.TireTempResult
        car.TireLFBattery = oLVT520.TireBatteryResult
        car.TireLFAcSpeed = oLVT520.TireAcSpeedResult

        updateState "dsglf", tmpID
        updateState "mdllf", car.TireLFMdl
        updateState "prelf", car.TireLFPre
        updateState "templf", car.TireLFTemp
        updateState "batterylf", car.TireLFBattery
        updateState "acspeedlf", car.TireLFAcSpeed

        '��ǰ�ּ�����
        setFrm TestStateFlag
    End If
End Sub
'�Һ���(����ʱ��)
Private Sub Command10_Click()

'    If DateDiff("s", tmpTime, Now) <= Step3Time Then
'        MsgBox ("��Ӧʱ��δ�ﵽҪ��!")
'        Exit Sub
'    Else
'        tmpTime = Now
'    End If


    TestStateFlag = 2
    Dim tmpID As String
    Dim i As Long
    If TestStateFlag = 2 Then

        TestStateFlag = 3
        updateState "state", CStr(TestStateFlag)
        isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ����Һ���
        AddMessage "���ڼ���Һ��֡���"
        LogWritter "��ʼ��һ�μ���Һ��֡���"
        oRVT520.ResetResult
        oRVT520.Start "Comm"

'        For i = 0 To 40
'            oRVT520.ReadResult
'            tmpID = oRVT520.TireIDResult
'            If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
'                Exit For
'            End If
'        Next i
'        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then   '�ұ�û�в⵽�ز�һ��
'            LogWritter "��ʼ�ڶ��μ���Һ��֡���"
'            oRVT520.ResetResult
'            oRVT520.Start "Comm"
'            For i = 0 To 40
'                oRVT520.ReadResult
'                tmpID = oRVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
'                    Exit For
'                End If
'            Next i
'        End If
'        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then   '�ұ�û�в⵽�ز�һ��
'            LogWritter "��ʼ�����μ���Һ��֡���"
'            oRVT520.ResetResult
'            oRVT520.Start "Comm"
'            For i = 0 To 40
'                oRVT520.ReadResult
'                tmpID = oRVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
'                    Exit For
'                End If
'            Next i
'        End If

        For i = 0 To 6
            oRVT520.ReadResult
            tmpID = oRVT520.TireIDResult
            'If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
            If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                Exit For
            End If
        Next i
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '�ڶ��β���
            LogWritter "��ʼ�ڶ��μ���Һ��֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '�����β���
            LogWritter "��ʼ�����μ���Һ��֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '���Ĵβ���
            LogWritter "��ʼ���Ĵμ���Һ��֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '����β���
            LogWritter "��ʼ����μ���Һ��֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"
            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                    Exit For
                End If
            Next i
        End If
        
        isInTesting = False 'Add by ZCJ 2012-07-09 �Һ��ּ�����

        car.TireRRID = tmpID
        LogWritter "�Һ��ּ�����ݣ�" & oRVT520.Result
        car.TireRRMdl = oRVT520.TireMdlResult
        car.TireRRPre = oRVT520.TirePreResult
        car.TireRRTemp = oRVT520.TireTempResult
        car.TireRRBattery = oRVT520.TireBatteryResult
        car.TireRRAcSpeed = oRVT520.TireAcSpeedResult

        updateState "dsgrr", tmpID
        updateState "mdlrr", car.TireRRMdl
        updateState "prerr", car.TireRRPre
        updateState "temprr", car.TireRRTemp
        updateState "batteryrr", car.TireRRBattery
        updateState "acspeedrr", car.TireRRAcSpeed

        TestStateFlag = 3 '�Һ��ּ�����
        updateState "state", CStr(TestStateFlag)
        setFrm TestStateFlag
    End If
End Sub
'�����(����ʱ��)
Private Sub Command11_Click()

'    If DateDiff("s", tmpTime, Now) <= Step4Time Then
'        MsgBox ("��Ӧʱ��δ�ﵽҪ��!")
'        Exit Sub
'    Else
'        tmpTime = Now
'    End If


    TestStateFlag = 3
    Dim tmpID As String
    Dim i As Long
    If TestStateFlag = 3 Then

        isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ��������
        AddMessage "���ڼ������֡���"
        LogWritter "��ʼ��һ�μ������֡���"
        oLVT520.ResetResult
        oLVT520.Start "Comm"

'        For i = 0 To 40
'            oLVT520.ReadResult
'            tmpID = oLVT520.TireIDResult
'            If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID Then
'                Exit For
'            End If
'        Next i
'        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Then '���û�в⵽�ز�һ��
'            LogWritter "��ʼ�ڶ��μ������֡���"
'            oLVT520.ResetResult
'            oLVT520.Start "Comm"
'            For i = 0 To 40
'                oLVT520.ReadResult
'                tmpID = oLVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID Then
'                    Exit For
'                End If
'            Next i
'        End If
'        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Then '���û�в⵽�ز�һ��
'            LogWritter "��ʼ�����μ������֡���"
'            oLVT520.ResetResult
'            oLVT520.Start "Comm"
'            For i = 0 To 40
'                oLVT520.ReadResult
'                tmpID = oLVT520.TireIDResult
'                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID Then
'                    Exit For
'                End If
'            Next i
'        End If

        For i = 0 To 6
            oLVT520.ReadResult
            tmpID = oLVT520.TireIDResult
            'If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID Then
            If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                Exit For
            End If
        Next i
        'If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID '�ڶ��β���
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '�ڶ��β���
            LogWritter "��ʼ�ڶ��μ������֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '�����β���
            LogWritter "��ʼ�����μ������֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                    Exit For
                End If
            Next i
        End If
        
        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '���Ĵβ���
            LogWritter "��ʼ���Ĵμ������֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                    Exit For
                End If
            Next i
        End If

        If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '����β���
            LogWritter "��ʼ����μ������֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                    Exit For
                End If
            Next i
        End If

        isInTesting = False 'Add by ZCJ 2012-07-09 ����ּ�����

        car.TireLRID = tmpID
        LogWritter "����ּ�����ݣ�" & oLVT520.Result
        car.TireLRMdl = oLVT520.TireMdlResult
        car.TireLRPre = oLVT520.TirePreResult
        car.TireLRTemp = oLVT520.TireTempResult
        car.TireLRBattery = oLVT520.TireBatteryResult
        car.TireLRAcSpeed = oLVT520.TireAcSpeedResult

        updateState "dsglr", tmpID
        updateState "mdllr", car.TireLRMdl
        updateState "prelr", car.TireLRPre
        updateState "templr", car.TireLRTemp
        updateState "batterylr", car.TireLRBattery
        updateState "acspeedlr", car.TireLRAcSpeed

        TestStateFlag = 4 '���ּ�����
        updateState "state", CStr(TestStateFlag)
        setFrm TestStateFlag

        If TestStateFlag = 4 Then
            LogWritter "�����ɣ�"

            car.Save
            If car.GetTestState = 15 Then
'����ָ����Χ�򱨾�
'                car.CheckResultIsOverStandard
'                If car.IsOverStandard Then
'                     Call printErrResult(car)
'                Else
'                    flashLamp Lamp_YellowFlash_IOPort
'                End If
            Else
                flashBuzzerLamp Lamp_RedLight_IOPort
                AddMessage "����������ظ�ֵ��", True
                LogWritter "����������ظ�ֵ��������ӡ��"
                If car.printFlag And car.LastCar.GetTestState <> 15 Then
                    Call printErrResult(car.LastCar)
                End If

                Call printErrResult(car)
            End If
            DSGTestEnd
        ElseIf TestStateFlag = 9994 Then
            DSGTestEnd
        End If

    End If
End Sub
'******************************************************************************
'** �� �� ����Form_Load
'** ��    �룺
'** ��    ����
'** �����������������ʱ����Ӧ
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub Form_Load()
    
    'Add by ZCJ 2012-07-09 ��ʼ������״̬
    isInTesting = False
    osen0Time = ""
    'Add by ZCJ 2012-07-09 ��ʼ�����ʱ��
    tmpTime = DateAdd("s", -30, Now)
    
    barCodeFlag = False
    frmInfo.Show
    initFrom True
    Dim testFlag As Boolean
    TestStateFlag = readState("state")
    testFlag = readState("test")    '�Ƿ��DSG

    TimerN = getConfigValue("T_RunParam", "Timer", "TimerDataSync")     '�Ų�����ͬ������
    TimerStatus = getConfigValue("T_RunParam", "Timer", "TimerStatus")  'ϵͳ״̬���������
    DBPosition = getConfigValue("T_RunParam", "Status", "DBPosition")   '���ݿ������̷�
    SpaceAvailable = getConfigValue("T_RunParam", "Status", "SpaceAvailable")   '���ݿ�����Ӳ�̿��ÿռ�����

    '�����DSGϵͳ����δ�����ɣ��ȼ����Ѽ���˵�����
    If testFlag And TestStateFlag <> 9999 Then
        Set car = getRunStateCar
        Me.txtVIN.text = car.VINCode
    End If
    '����Ѽ����ɣ�������ݿ��м���VIN
    If TestStateFlag > 9000 And TestStateFlag < 9999 Or TestStateFlag = -1 Then
       Me.txtVIN.text = readState("vin")
    End If
    frmInfo.labNow.Caption = Right(Me.txtVIN.text, 8)
    If Me.txtVIN.text <> "" Then
        frmInfo.labVin = Me.txtVIN.text
    End If
    setFrm TestStateFlag

    Step1Time = 4 '8
    Step2Time = 13 '17
    Step3Time = 13 '17
    Step4Time = 14 '18

    updateState "state", CStr(TestStateFlag)
    '������󼯺�
    Set inputCode = New Dictionary
        
    'Modiy by ZCJ 2012-07-09 �������¼��ƶ����˴�
    Set osensorCommand = sensorCommand      '�����¼�
    osensorCommand_onChange sensorCommand.state
    
    '������
    Set osensor0 = sensor0
    Set osensor1 = sensor1
    Set osensor2 = sensor2
    Set osensor3 = sensor3
    Set osensor4 = sensor4
    Set osensor5 = sensor5
    Set osensorLine = sensorLine            'ͣ���¼�
    Set oRDCommand = rdResetCommandS        'ϵͳ��λ�¼�
    DelayTime 1000

    sensorFlag = osensorLine.state
    sensorControlFlag = False   '������״̬,False��ʾû����
    testEndDelyed = False   '�˱�ʾ��TestStateFlag=-1����ʹ��

    initDictionary
    iniListInput
    flashLamp Lamp_GreenLight_IOPort
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call setWirledComScan     '��ʼ��ɨ��ǹ�Ĵ���
    Call setWirlessComScan
End Sub

'�رճ����ȹرյ��������ͷŴ���
Private Sub Form_Unload(Cancel As Integer)
    Call closeAll
    Dim X As Form

    For Each X In Forms
        Unload X
    Next
End Sub

'��������ǹͨ��
Private Sub MSCommBT_OnComm()
On Error GoTo MSCommBT_OnComm_Err
    If BreakFlag Then Exit Sub
    DelayTime 100
    Dim tmp As Variant
    Dim strin As String
    tmp = MSCommBT.Input
    If tmp = "" Then Exit Sub
    strin = strin & tmp
    TestCode = strin
    If Left(TestCode, 17) = "R010000000000000C" Then '��������
        LogWritter "0ɨ����������"
        resetList
        Exit Sub
    End If
    If Left(TestCode, 17) = "R020000000000000C" Then 'ǿ����������
        LogWritter "ɨ��ǿ����������"
        barCodeFlag = True
        Exit Sub
    End If
    Debug.Print TestCode
    tmp = ""
    Call txtVIN_KeyPress(13)
    Exit Sub
MSCommBT_OnComm_Err:
    LogWritter "����ɨ��ǹͨ�Ŵ���" & Err.Description
End Sub
'�������ϵĸ�λ��ť�¼�
Private Sub oRDCommand_onChange(state As Boolean)
    If state Then
        If BreakFlag Then Exit Sub
        LogWritter "ϵͳ����λ"
        resetList
    End If
End Sub
'0�Ŵ�����
Private Sub osensor0_onChange(state As Boolean)
    SensorLogWritter "osensor0----" + CStr(state)
    If BreakFlag Then Exit Sub
    
    If osen0Time <> "" Then
        If DateDiff("s", osen0Time, Now) <= 3 Then
            SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor0�¼�δ��Ӧ."
            Exit Sub
        Else
            osen0Time = Now
        End If
    Else
        osen0Time = Now
    End If
    
    If state = True Then
        '�������빤λ��һ����ʶ
        firstFlag = True
        flashLamp Lamp_YellowFlash_IOPort
    ElseIf secondFlag And osensor4.state Then
        If TestStateFlag < 10 And TestStateFlag <> 3 And TestStateFlag <> 0 And TestStateFlag <> -1 Then
        'If TestStateFlag < 10 And TestStateFlag <> 1 And TestStateFlag <> 3 And TestStateFlag <> 0 Then
            LogWritter "�����ɣ�"

            car.Save
            If car.GetTestState = 15 Then
'                car.CheckResultIsOverStandard
'                If car.IsOverStandard Then
'                     Call printErrResult(car)
'                End If
            Else
                flashBuzzerLamp Lamp_RedLight_IOPort
                AddMessage "����������ظ�ֵ��", True
                LogWritter "����������ظ�ֵ��������ӡ��"
                If car.printFlag And car.LastCar.GetTestState <> 15 Then
                    Call printErrResult(car.LastCar)
                End If
                Call printErrResult(car)
            End If
            AddMessage "��ע������Ƿ���ȷ", True
            LogWritter "���ְ�̨������"
            DSGTestEnd

            DelayTime 5000
            oIOCard.OutputController rdOutput, False
            oIOCard.OutputController Lamp_RedLight_IOPort, False
            oIOCard.OutputController Lamp_GreenLight_IOPort, True
        ElseIf TestStateFlag > 9990 And TestStateFlag <> 9995 And TestStateFlag <> 9999 And TestStateFlag <> -1 Then
        'ElseIf TestStateFlag > 9990 And TestStateFlag <> 9998 And TestStateFlag <> 9997 And TestStateFlag <> 9995 And TestStateFlag <> 9999 Then
            AddMessage "��ע������Ƿ���ȷ", True
            LogWritter "���ְ�̨������"
            DSGTestEnd

        End If
    End If

End Sub
'1�Ŵ�����
Private Sub osensor1_onChange(state As Boolean)
    SensorLogWritter "osensor1----" + CStr(state)
    If BreakFlag Then Exit Sub

    secondFlag = state
    If Not firstFlag Then
        '�����쳣����
    End If

    If firstFlag And secondFlag Then
        '�������繤λ�ȴ���ʼ����
        firstFlag = False
        'secondFlag = False
        If inputCode.Count <> 0 Then
        '�ٴ�����DSGStart
            Call Me.DSGTestStart(CStr(inputCode(inputCode.Keys(0))))
            tmpTime = Now
        End If

    End If
End Sub
'2�Ŵ�����
Private Sub osensor2_onChange(state As Boolean)
SensorLogWritter "osensor2----" + CStr(state)

On Error Resume Next
    If BreakFlag Then Exit Sub
    '��������ֹͣ������Ӧֹͣ��ʱ���˳�����
    If Not sensorFlag And sensorControlFlag Then
        SensorLogWritter "������ֹͣ�¼�δ��Ӧ"
        Exit Sub
    End If
    
    'Add by ZCJ 2012-08-09 �����ڼ��ʱ���˳�
    If isInTesting Then Exit Sub
    
    Dim tmpID As String
    Dim i As Long
    DelayTime 800
    If osensor1.state And osensor0.state And osensor2.state = state Then
        If TestStateFlag = 0 Then
            '�������̣����빤λ
            '�����ǰ��

            If DateDiff("s", tmpTime, Now) <= Step1Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor2�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If


            TestStateFlag = 1
            updateState "state", CStr(TestStateFlag)
            
            isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ�����ǰ��
            
            AddMessage "���ڼ����ǰ�֡���"
            LogWritter "��ʼ��һ�μ����ǰ�֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"

            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                    Exit For
                End If
            Next i
            
            LogWritter "��һ�μ�����ݣ�" & tmpID
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Then '�ڶ��β���
                LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�ڶ��μ�����ݣ�" & tmpID
                
            End If

            If tmpID = "00000000" Or Trim(tmpID) = "" Then '�����β���
                LogWritter "��ʼ�����μ����ǰ�֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�����μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Then '���Ĵβ���
                LogWritter "��ʼ���Ĵμ����ǰ�֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                        Exit For
                    End If
                Next i
                
                LogWritter "���Ĵμ�����ݣ�" & tmpID
                
            End If

            If tmpID = "00000000" Or Trim(tmpID) = "" Then '����β���
                LogWritter "��ʼ����μ����ǰ�֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                        Exit For
                    End If
                Next i
                
                LogWritter "����μ�����ݣ�" & tmpID
                
            End If
            
            isInTesting = False 'Add by ZCJ 2012-07-09 ��ǰ�ּ�����

            car.TireRFID = tmpID
            LogWritter "��ǰ�ּ�����ݣ�" & oRVT520.Result
            car.TireRFMdl = oRVT520.TireMdlResult
            car.TireRFPre = oRVT520.TirePreResult
            car.TireRFTemp = oRVT520.TireTempResult
            car.TireRFBattery = oRVT520.TireBatteryResult
            car.TireRFAcSpeed = oRVT520.TireAcSpeedResult

            updateState "dsgrf", tmpID
            updateState "mdlrf", car.TireRFMdl
            updateState "prerf", car.TireRFPre
            updateState "temprf", car.TireRFTemp
            updateState "batteryrf", car.TireRFBattery
            updateState "acspeedrf", car.TireRFAcSpeed

            'ǰ�ּ�����
            setFrm TestStateFlag

        ElseIf TestStateFlag = 2 Then
            '����Һ���
            If DateDiff("s", tmpTime, Now) <= Step3Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor5�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If
            TestStateFlag = 3 '���ּ�����
            updateState "state", CStr(TestStateFlag)
            
            isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ����Һ���
            
            AddMessage "���ڼ���Һ��֡���"
            LogWritter "��ʼ��һ�μ���Һ��֡���"
            oRVT520.ResetResult
            oRVT520.Start "Comm"

            For i = 0 To 6
                oRVT520.ReadResult
                tmpID = oRVT520.TireIDResult
                'If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                    Exit For
                End If
            Next i
            
            LogWritter "��һ�μ�����ݣ�" & tmpID
            
            'If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then   '�ڶ��β���
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '�ڶ��β���
                LogWritter "��ʼ�ڶ��μ���Һ��֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�ڶ��μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '�����β���
                LogWritter "��ʼ�����μ���Һ��֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�����μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '���Ĵβ���
                LogWritter "��ʼ���Ĵμ���Һ��֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "���Ĵμ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireLFID Then   '����β���
                LogWritter "��ʼ����μ���Һ��֡���"
                oRVT520.ResetResult
                oRVT520.Start "Comm"
                For i = 0 To 6
                    oRVT520.ReadResult
                    tmpID = oRVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireLFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "����μ�����ݣ�" & tmpID
                
            End If

            isInTesting = False 'Add by ZCJ 2012-07-09 �Һ��ּ�����

            car.TireRRID = tmpID
            LogWritter "�Һ��ּ�����ݣ�" & oRVT520.Result
            car.TireRRMdl = oRVT520.TireMdlResult
            car.TireRRPre = oRVT520.TirePreResult
            car.TireRRTemp = oRVT520.TireTempResult
            car.TireRRBattery = oRVT520.TireBatteryResult
            car.TireRRAcSpeed = oRVT520.TireAcSpeedResult

            updateState "dsgrr", tmpID
            updateState "mdlrr", car.TireRRMdl
            updateState "prerr", car.TireRRPre
            updateState "temprr", car.TireRRTemp
            updateState "batteryrr", car.TireRRBattery
            updateState "acspeedrr", car.TireRRAcSpeed

            setFrm TestStateFlag
        ElseIf TestStateFlag = 9998 Then
            '����DSG�ĳ�

            If DateDiff("s", tmpTime, Now) <= Step1Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor2�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If


            TestStateFlag = TestStateFlag - 1
            updateState "state", CStr(TestStateFlag)
            setFrm TestStateFlag
        ElseIf TestStateFlag = 9996 Then
            If DateDiff("s", tmpTime, Now) <= Step3Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor2�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If


            TestStateFlag = TestStateFlag - 1
            updateState "state", CStr(TestStateFlag)
            setFrm TestStateFlag
        End If
        
        isInTesting = False 'Add by ZCJ 2012-07-09 ��ʼ����̥���״̬
    Else
        
    End If
End Sub
'������3
Private Sub osensor3_onChange(state As Boolean)
    SensorLogWritter "osensor3----" + CStr(state)
End Sub
'������4
Private Sub osensor4_onChange(state As Boolean)
    SensorLogWritter "osensor4----" + CStr(state)
End Sub
'������5
Private Sub osensor5_onChange(state As Boolean)
SensorLogWritter "osensor5----" + CStr(state)

    On Error Resume Next
    If BreakFlag Then Exit Sub
    If Not sensorFlag And sensorControlFlag Then
        SensorLogWritter "������ֹͣ�¼�δ��Ӧ"
        Exit Sub
    End If
    
    'Add by ZCJ 2012-08-09 �����ڼ��ʱ���˳�
    If isInTesting Then Exit Sub
    
    Dim tmpID As String
    Dim i As Long
    DelayTime 800
    If osensor3.state And osensor4.state And osensor5.state = state Then
        If TestStateFlag = 1 Then
            '�������̣����빤λ
            '�����ǰ��

            If DateDiff("s", tmpTime, Now) <= Step2Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor5�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If

            TestStateFlag = 2
            updateState "state", CStr(TestStateFlag)
            
            isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ�����ǰ��
            
            AddMessage "���ڼ����ǰ�֡���"
            LogWritter "��ʼ��һ�μ����ǰ�֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"
            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                'If tmpID <> "00000000" And Trim(tmpID) <> "" Then
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                    Exit For
                End If
            Next i
            
            LogWritter "��һ�μ�����ݣ�" & tmpID
            
            'If tmpID = "00000000" Or Trim(tmpID) = "" Then '�ڶ��β���
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '�ڶ��β���
                LogWritter "��ʼ�ڶ��μ����ǰ�֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�ڶ��μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '�����β���
                LogWritter "��ʼ�����μ����ǰ�֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�����μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '���Ĵβ���
                LogWritter "��ʼ���Ĵμ����ǰ�֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "���Ĵμ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireRFID Then '����β���
                LogWritter "��ʼ����μ����ǰ�֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireRFID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "����μ�����ݣ�" & tmpID
                
            End If
            
            isInTesting = False 'Add by ZCJ 2012-07-09 ��ǰ�ּ�����

            car.TireLFID = tmpID
            LogWritter "��ǰ�ּ�����ݣ�" & oLVT520.Result
            car.TireLFMdl = oLVT520.TireMdlResult
            car.TireLFPre = oLVT520.TirePreResult
            car.TireLFTemp = oLVT520.TireTempResult
            car.TireLFBattery = oLVT520.TireBatteryResult
            car.TireLFAcSpeed = oLVT520.TireAcSpeedResult

            updateState "dsglf", tmpID
            updateState "mdllf", car.TireLFMdl
            updateState "prelf", car.TireLFPre
            updateState "templf", car.TireLFTemp
            updateState "batterylf", car.TireLFBattery
            updateState "acspeedlf", car.TireLFAcSpeed

             'ǰ�ּ�����
            setFrm TestStateFlag
        ElseIf TestStateFlag = 3 Then
            '��������
            If DateDiff("s", tmpTime, Now) <= Step4Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor5�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If

            TestStateFlag = 4
            updateState "state", CStr(TestStateFlag)
            
            isInTesting = True 'Add by ZCJ 2012-07-09 ��ʼ��������
            
            AddMessage "���ڼ������֡���"
            LogWritter "��ʼ��һ�μ������֡���"
            oLVT520.ResetResult
            oLVT520.Start "Comm"

            For i = 0 To 6
                oLVT520.ReadResult
                tmpID = oLVT520.TireIDResult
                'If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID Then
                If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                    Exit For
                End If
            Next i
            
            LogWritter "��һ�μ�����ݣ�" & tmpID
            
            'If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Then '�ڶ��β���
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then
                LogWritter "��ʼ�ڶ��μ������֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�ڶ��μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '�����β���
                LogWritter "��ʼ�����μ������֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "�����μ�����ݣ�" & tmpID
                
            End If
            
            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '���Ĵβ���
                LogWritter "��ʼ���Ĵμ������֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                        Exit For
                    End If
                Next i
                
                LogWritter "���Ĵμ�����ݣ�" & tmpID
                
            End If

            If tmpID = "00000000" Or Trim(tmpID) = "" Or Trim(tmpID) = car.TireLFID Or Trim(tmpID) = car.TireRFID Or Trim(tmpID) = car.TireRRID Then '����β���
                LogWritter "��ʼ����μ������֡���"
                oLVT520.ResetResult
                oLVT520.Start "Comm"
                For i = 0 To 6
                    oLVT520.ReadResult
                    tmpID = oLVT520.TireIDResult
                    If tmpID <> "00000000" And Trim(tmpID) <> "" And Trim(tmpID) <> car.TireLFID And Trim(tmpID) <> car.TireRFID And Trim(tmpID) <> car.TireRRID Then
                        Exit For
                    End If
                Next i
                
                
                LogWritter "����μ�����ݣ�" & tmpID
                
            End If

            isInTesting = False 'Add by ZCJ 2012-07-09 ����ּ�����

            car.TireLRID = tmpID
            LogWritter "����ּ�����ݣ�" & oLVT520.Result
            car.TireLRMdl = oLVT520.TireMdlResult
            car.TireLRPre = oLVT520.TirePreResult
            car.TireLRTemp = oLVT520.TireTempResult
            car.TireLRBattery = oLVT520.TireBatteryResult
            car.TireLRAcSpeed = oLVT520.TireAcSpeedResult

            updateState "dsglr", tmpID
            updateState "mdllr", car.TireLRMdl
            updateState "prelr", car.TireLRPre
            updateState "templr", car.TireLRTemp
            updateState "batterylr", car.TireLRBattery
            updateState "acspeedlr", car.TireLRAcSpeed

            '���ּ�����
            setFrm TestStateFlag
            DelayTime 200 '������ڽ�����ʾ0.2��
        ElseIf TestStateFlag = 9997 Then
            '����DSG�ĳ�
            If DateDiff("s", tmpTime, Now) <= Step2Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor5�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If
            TestStateFlag = TestStateFlag - 1
            updateState "state", CStr(TestStateFlag)
            setFrm TestStateFlag
        ElseIf TestStateFlag = 9995 Then
            '����DSG�ĳ�
            If DateDiff("s", tmpTime, Now) <= Step4Time Then
                SensorLogWritter "��Ӧʱ��δ�ﵽҪ��osensor5�¼�δ��Ӧ."
                Exit Sub
            Else
                tmpTime = Now
            End If
            TestStateFlag = TestStateFlag - 1
            updateState "state", CStr(TestStateFlag)
            setFrm TestStateFlag
        End If

        If TestStateFlag = 4 Then
            LogWritter "�����ɣ�"

            car.Save
            If car.GetTestState = 15 Then
'                car.CheckResultIsOverStandard
'                If car.IsOverStandard Then
'                     Call printErrResult(car)
'                Else
                    flashLamp Lamp_YellowFlash_IOPort
                    'flashLamp Lamp_GreenFlash_IOPort
'                End If
            Else
                flashBuzzerLamp Lamp_RedLight_IOPort
                AddMessage "����������ظ�ֵ��", True
                LogWritter "����������ظ�ֵ��������ӡ��"
                If car.printFlag And car.LastCar.GetTestState <> 15 Then
                    Call printErrResult(car.LastCar)
                End If

                Call printErrResult(car)
            End If
            DSGTestEnd

            DelayTime 5000
            oIOCard.OutputController rdOutput, False
            oIOCard.OutputController Lamp_RedLight_IOPort, False
            oIOCard.OutputController Lamp_GreenLight_IOPort, True
        ElseIf TestStateFlag = 9994 Then
            'oIOCard.OutputController rdOutput, True
            DSGTestEnd
        End If

    Else

    End If

End Sub
'���������¼�
Private Sub osensorCommand_onChange(state As Boolean)
    SensorLogWritter "osensorCommand----" + CStr(state)
    BreakFlag = Not state
    If state Then
'        If lineCommandFlag Then
'            oIOCard.OutputController sensorLinePort, True
'        End If

        AddMessage "ϵͳ�ѽ�����", True
        setFrm TestStateFlag
        LogWritter "ϵͳ�ѽ�����"
        Timer_PrintError.Interval = 1000
    Else
'        If lineCommandFlag Then
'            oIOCard.OutputController sensorLinePort, False
'        End If

        AddMessage "ϵͳ�ѱ��������������", True
        LogWritter "ϵͳ��������"
        Timer_PrintError.Interval = 0
    End If
End Sub
'ͣ���¼�
Private Sub osensorLine_onChange(state As Boolean)
    SensorLogWritter "sensorLine----" + CStr(state)
    sensorFlag = state
End Sub

Private Sub Timer_PrintError_Timer()
On Error GoTo Err
    HH = HH + 1

    If HH < 5 Then
        Exit Sub
    End If
    
    'Call printErrCode
    Call printErrCodeAuto
    
    HH = 0
    Exit Sub
Err:
    LogWritter "printErrCode timer error"
    HH = 0
    Exit Sub
End Sub

Private Sub txtInputVIN_GotFocus()
    txtInputVIN.text = ""
End Sub

Private Sub txtInputVIN_KeyPress(KeyAscii As Integer)
    If BreakFlag Then Exit Sub
    Dim tmp As String
    If KeyAscii = 13 Then '�س�����
        tmp = txtInputVIN.text
        
        If tmp = "" Then Exit Sub
        TestCode = tmp
        If Left(TestCode, 17) = "R010000000000000C" Then
            LogWritter "1ɨ����������"
            resetList
            txtInputVIN.text = "�ֹ�¼��VIN���س�ȷ��"
            Exit Sub
        End If
        If Left(TestCode, 17) = "R020000000000000C" Then
            barCodeFlag = True
            txtInputVIN.text = "�ֹ�¼��VIN���س�ȷ��"
            Exit Sub
        End If
    
        Debug.Print TestCode
        Call txtVIN_KeyPress(13)
        txtInputVIN.text = "�ֹ�¼��VIN���س�ȷ��"
    End If
End Sub

Private Sub txtInputVIN_LostFocus()
    txtInputVIN.text = "�ֹ�¼��VIN���س�ȷ��"
End Sub

'����ɨ��������Ϣ
Private Sub txtVIN_KeyPress(KeyAscii As Integer)
    
    Dim tmpCode As String, tmpKey As String
    tmpCode = TestCode
    tmpKey = Mid(tmpCode, 2, 17)
    
    If BreakFlag Then Exit Sub
    If KeyAscii = 13 Then


    TestCode = Trim(TestCode)
    TestCode = Replace(TestCode, Chr(10), "")
    TestCode = Replace(TestCode, Chr(13), "")
    LogWritter "************************************************************"
    LogWritter "ɨ�����룺" & TestCode
    LogWritter "************************************************************"
        If Len(TestCode) = 26 Then
            If isCheckAllQueue Then
                If frmInfo.ListInput.ListCount <> 0 And barCodeFlag = False Then
                    If frmInfo.labNext.Caption <> Right(tmpKey, 8) Then
                        AddMessage "��ע���ɨ������Ϣ�Ƿ���ȷ", True
                        flashBuzzerLamp Lamp_RedLight_IOPort
                        LogWritter "��ɨ������ƥ��,������������"
                        DelayTime 2000
                        oIOCard.OutputController Lamp_RedLight_IOPort, False
                        oIOCard.OutputController rdOutput, False
                        If TestStateFlag = 9999 Or TestStateFlag = -1 Then
                            oIOCard.OutputController Lamp_GreenLight_IOPort, True
                        Else
                            oIOCard.OutputController Lamp_YellowFlash_IOPort, True
                        End If
                        Exit Sub
                    End If
                End If
            End If
            If barCodeFlag Then
                barCodeFlag = False
            End If
            If inputCode.Exists(tmpKey) Then
                Exit Sub
            End If
                    
            inputCode.Add tmpKey, tmpCode
            insertColl tmpCode
            LogWritter tmpKey & "����ɨ�����"
            Me.List1.AddItem tmpKey
            frmInfo.ListOutput.AddItem Right(tmpKey, 8)
            setFrm TestStateFlag
            initDictionary
            If inputCode.Count = 1 Then
                txtVIN.text = CStr(Mid(inputCode(inputCode.Keys(0)), 2, 17))
                frmInfo.labVin.Caption = txtVIN.text
                updateState "test", "False"
                updateState "vin", txtVIN.text
                TestStateFlag = -1
                updateState "state", -1
                AddMessage "�ȴ�ɨ�賵�����빤λ!"
            End If
            iniListInput
            flashLamp Lamp_GreenFlash_IOPort
            DelayTime 1000
            flashLamp Lamp_GreenLight_IOPort
            If TestStateFlag = 9999 Or TestStateFlag = -1 Then
                oIOCard.OutputController Lamp_GreenLight_IOPort, True
            Else
                oIOCard.OutputController Lamp_GreenLight_IOPort, False
                oIOCard.OutputController Lamp_YellowFlash_IOPort, True
            End If
        Else
            AddMessage "��ע��ɨ�����볤���Ƿ���ȷ", True
            flashBuzzerLamp Lamp_RedLight_IOPort
            LogWritter "���볤�Ȳ���ȷ,������������!"
            DelayTime 2000
            oIOCard.OutputController Lamp_RedLight_IOPort, False
            oIOCard.OutputController rdOutput, False
            If TestStateFlag = 9999 Or TestStateFlag = -1 Then
                oIOCard.OutputController Lamp_GreenLight_IOPort, True
            Else
                oIOCard.OutputController Lamp_GreenLight_IOPort, False
                oIOCard.OutputController Lamp_YellowFlash_IOPort, True
            End If
        End If

    End If
End Sub

'******************************************************************************
'** �� �� ����DSGTestStart
'** ��    �룺
'** ��    ����
'** ����������DSG���Կ�ʼ
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Sub DSGTestStart(vin As String)

    isInTesting = False 'Add by ZCJ 2012-07-09 ��ʼ����̥���״̬

    If TestStateFlag <> 9999 Then
        If TestStateFlag <> -1 Then
            '����������������
            Exit Sub
        End If
    End If

    txtVIN.text = Mid(vin, 2, 17)
    frmInfo.labVin.Caption = txtVIN.text
    frmInfo.labNow.Caption = Right(txtVIN.text, 8)
    LogWritter "============================================================"
    LogWritter txtVIN.text & "��ʼ����!"
    If hasDSG(vin) Then
        LogWritter "������ͨ��,��ʼDSG���!"
        updateState "test", "True"
        updateState "vin", txtVIN.text
        Set car = New CCar
        car.VINCode = txtVIN.text
        TestStateFlag = 0
        setFrm TestStateFlag
        updateState "state", CStr(TestStateFlag)
        If osensor1.state Then
            osensor1_onChange True
        End If
    Else
        LogWritter "����δװ��DSG,ֱ��ͨ��!"
        updateState "test", "False"
        updateState "vin", txtVIN.text

        TestStateFlag = 9998
        setFrm TestStateFlag
        updateState "state", CStr(TestStateFlag)
    End If
End Sub
'�������
Public Sub DSGTestEnd()
On Error GoTo END_ERR

    isInTesting = False 'Add by ZCJ 2012-07-09 ��ʼ����̥���״̬

    testEndDelyed = True
    TestStateFlag = 9999
    resetState
    LogWritter txtVIN.text & "�������!"
    LogWritter "============================================================"

    txtVIN.text = ""
    frmInfo.labNow.Caption = ""
    frmInfo.labVin.Caption = "̥ѹ����ʼ��ϵͳ"

    setFrm TestStateFlag
    LogWritter CStr(inputCode.Keys(0)) & "�˳�ɨ�����!"
    delColl CStr(inputCode.Keys(0))
    inputCode.Remove inputCode.Keys(0)
    If inputCode.Count <> 0 Then
        updateState "vin", CStr(inputCode.Keys(0))
        TestStateFlag = -1
        updateState "state", CStr(TestStateFlag)
        If hasDSG(CStr(inputCode(inputCode.Keys(0)))) Then
            updateState "test", "True"
        Else
            updateState "test", "False"
        End If
    End If

    DelayTime 3000
    testEndDelyed = False
    flashLamp Lamp_GreenLight_IOPort

    iniListInput
    initDictionary

    If inputCode.Count <> 0 Then
        '�ٴ�����DSGStart
        Call Me.DSGTestStart(CStr(inputCode(inputCode.Keys(0))))
    Else
        LogWritter "ɨ������г�����Ϊ��"
    End If

    Exit Sub
END_ERR:
    LogWritter Err.Description
End Sub
'�ڽ�������ʾ��⵽�Ĵ�������Ϣ
Public Sub showDSGInfo(str As String, text As String, model As String, pressure As String, temperature As String, battery As String, acSpeed As String, imgName As String)
    On Error Resume Next
    Dim Result As Boolean
    Dim mdlArr() As String
    
    FrmMain.Controls("txt" & str).text = text
    FrmMain.Controls("pic" & str).Picture = LoadPicture(App.Path & "\img\" & imgName)
    frmInfo.Controls("txt" & str).text = text
    frmInfo.Controls("pic" & str).Picture = LoadPicture(App.Path & "\img\" & imgName)
    FrmMain.Controls("lb" & str & "Mdl").Caption = model
    frmInfo.Controls("lb" & str & "Mdl").Caption = model

    mdlArr = Split(mdlValue, ",")
    Result = judgeMdlIsOK(model, mdlArr)
    If Result Then
        FrmMain.Controls("lb" & str & "Mdl").ForeColor = &HFF0000
        frmInfo.Controls("lb" & str & "Mdl").ForeColor = &HFF0000
    Else
        FrmMain.Controls("lb" & str & "Mdl").ForeColor = &HFF&
        frmInfo.Controls("lb" & str & "Mdl").ForeColor = &HFF&
    End If
    FrmMain.Controls("lb" & str & "Mdl").Caption = model
    frmInfo.Controls("lb" & str & "Mdl").Caption = model
    

    Result = judgeResultIsOK(pressure, preMinValue, preMaxValue)
    If Result Then
        FrmMain.Controls("lb" & str & "Pre").ForeColor = &HFF0000
        frmInfo.Controls("lb" & str & "Pre").ForeColor = &HFF0000
    Else
        FrmMain.Controls("lb" & str & "Pre").ForeColor = &HFF&
        frmInfo.Controls("lb" & str & "Pre").ForeColor = &HFF&
    End If
    If pressure <> "" Then
        FrmMain.Controls("lb" & str & "Pre").Caption = pressure & "kPa"
        frmInfo.Controls("lb" & str & "Pre").Caption = pressure & "kPa"
    Else
        FrmMain.Controls("lb" & str & "Pre").Caption = ""
        frmInfo.Controls("lb" & str & "Pre").Caption = ""
    End If



    Result = judgeResultIsOK(temperature, tempMinValue, tempMaxValue)
    If Result Then
        FrmMain.Controls("lb" & str & "Temp").ForeColor = &HFF0000
        frmInfo.Controls("lb" & str & "Temp").ForeColor = &HFF0000
    Else
        FrmMain.Controls("lb" & str & "Temp").ForeColor = &HFF&
        frmInfo.Controls("lb" & str & "Temp").ForeColor = &HFF&
    End If
    If temperature <> "" Then
        FrmMain.Controls("lb" & str & "Temp").Caption = temperature & "��"
        frmInfo.Controls("lb" & str & "Temp").Caption = temperature & "��"
    Else
        FrmMain.Controls("lb" & str & "Temp").Caption = ""
        frmInfo.Controls("lb" & str & "Temp").Caption = ""
    End If


    If battery = "OK" Then
        FrmMain.Controls("lb" & str & "Battery").ForeColor = &HFF0000
        frmInfo.Controls("lb" & str & "Battery").ForeColor = &HFF0000
    Else
        FrmMain.Controls("lb" & str & "Battery").ForeColor = &HFF&
        frmInfo.Controls("lb" & str & "Battery").ForeColor = &HFF&
    End If
    FrmMain.Controls("lb" & str & "Battery").Caption = battery
    frmInfo.Controls("lb" & str & "Battery").Caption = battery



    Result = judgeResultIsOK(acSpeed, acSpeedMinValue, acSpeedMaxValue)
    If Result Then
        FrmMain.Controls("lb" & str & "AcSpeed").ForeColor = &HFF0000
        frmInfo.Controls("lb" & str & "AcSpeed").ForeColor = &HFF0000
    Else
        FrmMain.Controls("lb" & str & "AcSpeed").ForeColor = &HFF&
        frmInfo.Controls("lb" & str & "AcSpeed").ForeColor = &HFF&
    End If
    If acSpeed <> "" Then
        FrmMain.Controls("lb" & str & "AcSpeed").Caption = acSpeed & "g"
        frmInfo.Controls("lb" & str & "AcSpeed").Caption = acSpeed & "g"
    Else
        FrmMain.Controls("lb" & str & "AcSpeed").Caption = ""
        frmInfo.Controls("lb" & str & "AcSpeed").Caption = ""
    End If
End Sub

'��������ǹ������Ϣ����
Public Sub setWirledComScan()
On Error GoTo Err
    MSComVIN.CommPort = WirledCodeGun_PortNum
    MSComVIN.InBufferSize = 1024
    MSComVIN.OutBufferSize = 512
    MSComVIN.InBufferCount = 0
    MSComVIN.Settings = WirledCodeGun_Settings
    MSComVIN.InputMode = comInputModeText
    MSComVIN.RTSEnable = True
    MSComVIN.RThreshold = 1
    MSComVIN.PortOpen = True
    Exit Sub
Err:
    LogWritter "��������ǹ�������ô���" & Err.Description
End Sub
'��������ǹ������Ϣ����
Public Sub setWirlessComScan()
On Error GoTo Err
    MSCommBT.CommPort = WirlessCodeGun_PortNum
    MSCommBT.InBufferSize = 1024
    MSCommBT.OutBufferSize = 512
    MSCommBT.InBufferCount = 0
    MSCommBT.Settings = WirlessCodeGun_Settings
    MSCommBT.InputMode = comInputModeText
    MSCommBT.RTSEnable = True
    MSCommBT.RThreshold = 1
    MSCommBT.PortOpen = True
    Exit Sub
Err:
    LogWritter "��������ǹ�������ô���" & Err.Description
End Sub
'��ʾ��ǰ�ļ��״̬
Public Sub setFrm(state As Integer)
    If state = -1 Then
        AddMessage "�ȴ�ɨ�賵�����빤λ!"
        initFrom False
    ElseIf state = 9999 Then
        AddMessage "�ȴ�ɨ��VIN����ʼ����!"
        initFrom True
    ElseIf state > 9000 And state < 9999 Then
        AddMessage "����δװ��DSG��������ֱ��ͨ��!"
        Select Case state
        Case 9997
            AddMessage "δװ��DSG:��ǰ����ͨ����������"
        Case 9996
            AddMessage "δװ��DSG:��ǰ����ͨ����������"
        Case 9995
            AddMessage "δװ��DSG:�Һ�����ͨ����������"
        Case 9994
            AddMessage "δװ��DSG:�������ͨ����������"
        End Select

    Else
        Select Case state

        Case 0
            AddMessage "����ɨ��ͨ���ȴ��������빤λ,��ʼ����!"
            LogWritter "����ɨ��ͨ���ȴ��������빤λ,��ʼ����!"
            initFrom False
        Case 1
            If car.TireRFID <> "00000000" And Trim(car.TireRFID) <> "" Then
                showDSGInfo "RF", car.TireRFID, car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Green1.jpg"
                LogWritter "��ǰ�ּ������" & car.TireRFID
                AddMessage "��ǰ�ּ�����"
            Else
                'Modiy by ZCJ 2012=07-09 ���������ڼ����̥��״̬����
                If isInTesting = True Then
                    AddMessage "���ڼ����ǰ�֡���"
                Else
                    showDSGInfo "RF", "���ʧ��", car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Red1.jpg"
                    LogWritter "��ǰ�ּ��ʧ��"
                    AddMessage "��ǰ�ּ��ʧ��", True
                End If
            End If

        Case 2
            If car.TireRFID <> "00000000" And Trim(car.TireRFID) <> "" Then
                showDSGInfo "RF", car.TireRFID, car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "RF", "���ʧ��", car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Red1.jpg"
            End If
            If car.TireLFID <> "00000000" And Trim(car.TireLFID) <> "" Then
                showDSGInfo "LF", car.TireLFID, car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Green1.jpg"
                LogWritter "��ǰ�ּ������" & car.TireLFID
                AddMessage "��ǰ�ּ�����"
            Else
                'Modiy by ZCJ 2012=07-09 ���������ڼ����̥��״̬����
                If isInTesting = True Then
                    AddMessage "���ڼ����ǰ�֡���"
                Else
                    showDSGInfo "LF", "���ʧ��", car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Red1.jpg"
                    LogWritter "��ǰ�ּ��ʧ��"
                    AddMessage "��ǰ�ּ��ʧ��", True
                End If
            End If

        Case 3
            If car.TireRFID <> "00000000" And Trim(car.TireRFID) <> "" Then
                showDSGInfo "RF", car.TireRFID, car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "RF", "���ʧ��", car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Red1.jpg"
            End If
            If car.TireLFID <> "00000000" And Trim(car.TireLFID) <> "" Then
                showDSGInfo "LF", car.TireLFID, car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "LF", "���ʧ��", car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Red1.jpg"
            End If
            If car.TireRRID <> "00000000" And Trim(car.TireRRID) <> "" Then
                showDSGInfo "RR", car.TireRRID, car.TireRRMdl, car.TireRRPre, car.TireRRTemp, car.TireRRBattery, car.TireRRAcSpeed, "Green1.jpg"
                LogWritter "�Һ��ּ������" & car.TireRRID
                AddMessage "�Һ��ּ�����"
            Else
                'Modiy by ZCJ 2012=07-09 ���������ڼ����̥��״̬����
                If isInTesting = True Then
                    AddMessage "���ڼ���Һ��֡���"
                Else
                    showDSGInfo "RR", "���ʧ��", car.TireRRMdl, car.TireRRPre, car.TireRRTemp, car.TireRRBattery, car.TireRRAcSpeed, "Red1.jpg"
                    LogWritter "�Һ��ּ��ʧ��"
                    AddMessage "�Һ��ּ��ʧ��", True
                End If
            End If

        Case 4
            If car.TireRFID <> "00000000" And Trim(car.TireRFID) <> "" Then
                showDSGInfo "RF", car.TireRFID, car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "RF", "���ʧ��", car.TireRFMdl, car.TireRFPre, car.TireRFTemp, car.TireRFBattery, car.TireRFAcSpeed, "Red1.jpg"
            End If
            If car.TireLFID <> "00000000" And Trim(car.TireLFID) <> "" Then
                showDSGInfo "LF", car.TireLFID, car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "LF", "���ʧ��", car.TireLFMdl, car.TireLFPre, car.TireLFTemp, car.TireLFBattery, car.TireLFAcSpeed, "Red1.jpg"
            End If
            If car.TireRRID <> "00000000" And Trim(car.TireRRID) <> "" Then
                showDSGInfo "RR", car.TireRRID, car.TireRRMdl, car.TireRRPre, car.TireRRTemp, car.TireRRBattery, car.TireRRAcSpeed, "Green1.jpg"
            Else
                showDSGInfo "RR", "���ʧ��", car.TireRRMdl, car.TireRRPre, car.TireRRTemp, car.TireRRBattery, car.TireRRAcSpeed, "Red1.jpg"
            End If
            If car.TireLRID <> "00000000" And Trim(car.TireLRID) <> "" Then
                showDSGInfo "LR", car.TireLRID, car.TireLRMdl, car.TireLRPre, car.TireLRTemp, car.TireLRBattery, car.TireLRAcSpeed, "Green1.jpg"
                LogWritter "����ּ������" & car.TireLRID
                AddMessage "����ּ�����"
            Else
                'Modiy by ZCJ 2012=07-09 ���������ڼ����̥��״̬����
                If isInTesting = True Then
                    AddMessage "���ڼ������֡���"
                Else
                    showDSGInfo "LR", "���ʧ��", car.TireLRMdl, car.TireLRPre, car.TireLRTemp, car.TireLRBattery, car.TireLRAcSpeed, "Red1.jpg"
                    LogWritter "����ּ��ʧ��"
                    AddMessage "����ּ��ʧ��", True
                End If
            End If

        End Select
    End If

End Sub
'��������ɨ��ǹ��ɨ����Ϣ
Private Sub MSComVIN_OnComm()
If BreakFlag Then Exit Sub
    DelayTime 100
    Dim tmp As Variant
    Dim strin As String
    tmp = MSComVIN.Input
    If tmp = "" Then Exit Sub
    strin = strin & tmp
    TestCode = strin
    If Left(TestCode, 17) = "R010000000000000C" Then
        LogWritter "1ɨ����������"
        resetList
        Exit Sub
    End If
    If Left(TestCode, 17) = "R020000000000000C" Then
        barCodeFlag = True
        Exit Sub
    End If

    Debug.Print TestCode
    tmp = ""
    Call txtVIN_KeyPress(13)

End Sub
'��ʼ��ɨ�������Ϣ
Public Sub initDictionary()
On Error Resume Next

    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select vin from vincoll order by id asc")
    inputCode.RemoveAll
    Me.List1.Clear
    frmInfo.ListOutput.Clear
    Do While Not rs.EOF
        inputCode.Add Mid(rs("vin").value, 2, 17), rs("vin").value
        Me.List1.AddItem Mid(rs("vin").value, 2, 17)
        frmInfo.ListOutput.AddItem Right(Mid(rs("vin").value, 2, 17), 8)
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub
'��ʼ���Ų�������Ϣ
Public Sub iniListInput()
On Error Resume Next
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim tmpStr As String
    Dim flag As Boolean
    Dim tmpVIN As String

    cnn.Open DBCnnStr
    If Me.txtVIN.text <> "" Then
        tmpVIN = Me.txtVIN.text
    Else
        tmpVIN = readState("vin")
    End If
    Set rs = cnn.Execute("select uw5anoseq from vinlist where vin = '" & tmpVIN & "' order by uw5anoseq desc limit 1")
    If rs.EOF Then

        If Me.txtVIN.text <> "" Then
            Exit Sub
        Else
            tmpStr = "999999999"
        End If
    Else
        tmpStr = rs(0)
    End If
    If TestStateFlag = 9999 And Me.txtVIN.text = "" Then
        Set rs = cnn.Execute("select vin from  vinlist where uw5anoseq > '" & tmpStr & "'  order by uw5anoseq asc limit 8")
    Else
        Set rs = cnn.Execute("select vin from  vinlist where uw5anoseq >= '" & tmpStr & "'  order by uw5anoseq asc limit 8")
    End If
    frmInfo.ListInput.Clear

    flag = False
    Do While Not rs.EOF
        frmInfo.ListInput.AddItem Right(rs(0), 8)

        If flag Then
            frmInfo.labNext.Caption = Right(rs(0), 8)
            flag = False
        End If
        If inputCode.Count <> 0 Then
            If rs(0) = inputCode.Keys(inputCode.Count - 1) Then
                flag = True
            End If
        End If
        rs.MoveNext
    Loop
    If inputCode.Count = 0 Then
         frmInfo.labNext.Caption = Right(frmInfo.ListInput.List(0), 8)
    End If
    cnn.Close
    Set cnn = Nothing
End Sub
'ϵͳ���ã�����λ
Public Sub resetList()
If BreakFlag Then Exit Sub

    VINCode = "" 'Add by ZCJ 2012-12-08
    MTOCCode = "InitMTOCCode" 'Add by ZCJ 2012-12-08

    delallColl
    initDictionary

    If testEndDelyed = False And TestStateFlag <> -1 Then
        TestStateFlag = 9999
    End If
    If TestStateFlag <> -1 Then
        resetState
        LogWritter txtVIN.text & "�������!"
        LogWritter "============================================================"
    End If
    txtVIN.text = ""
    
    setFrm 9999
    updateState "state", CStr(TestStateFlag) 'Add by ZCJ 20121207
    frmInfo.labNow.Caption = ""

    iniListInput
    
    Call closeAll
    oIOCard.OutputController Lamp_GreenLight_IOPort, True
    oIOCard.OutputController Lamp_Buzzer_IOPort, False '�رշ���
End Sub
'��������ƶ�
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Y > 0 And Y < 496 Then
        Dim ReturnVal As Long
        X = ReleaseCapture()
        ReturnVal = SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub
'��С������
Private Sub Picture1_Click()
    Me.WindowState = vbMinimized
End Sub
'�˳�ϵͳ
Private Sub picExit_Click()
    Dim msgR As Integer
    msgR = MsgBox("�Ƿ��˳�̥ѹ��ʼ��ϵͳ��", vbYesNo, "ϵͳ��ʾ")
    If msgR = 7 Then Exit Sub
    Dim X As Form
    For Each X In Forms
        Unload X
        Set X = Nothing
    Next
    oIOCard.OutputController Lamp_Buzzer_IOPort, False '�رշ���
    Call closeAll
    Call KillProcess("DSGTest.exe")
End Sub
'�����������رյ������������ߣ��κε�����������Ҫ�ȵ��ø÷���
Public Sub closeAll()
    'oIOCard.OutputController Lamp_Buzzer_IOPort, False '�رշ���
    oIOCard.OutputController Lamp_GreenLight_IOPort, False '�ر���ɫ
    oIOCard.OutputController Lamp_GreenFlash_IOPort, False '�ر���ɫ��˸
    oIOCard.OutputController Lamp_YellowLight_IOPort, False '�رջ�ɫ
    oIOCard.OutputController Lamp_YellowFlash_IOPort, False '�رջ�ɫ��˸
    oIOCard.OutputController Lamp_RedLight_IOPort, False '�رպ�ɫ
    oIOCard.OutputController Lamp_RedFlash_IOPort, False '�رպ�ɫ��˸
End Sub
'������������ʷ��¼��ѯ
Private Sub picCommandHis_Click()
    frmHistory.Show
End Sub
'������������־��ѯ
Private Sub picCommandLog_Click()
    frmShowLog.Show
End Sub
'�������������ݵ���
Private Sub picCommandOut_Click()
    frmDateZone.Show
End Sub
'����������ϵͳ����
Private Sub picCommandConifg_Click()
    frmPSW.Show
End Sub
'����������ϵͳ��λ
Private Sub picCommandReset_Click()
    If BreakFlag Then Exit Sub
    LogWritter "ϵͳ����λ"
    resetList

    Call closeAll
    oIOCard.OutputController Lamp_Buzzer_IOPort, False '�رշ���
    flashLamp Lamp_GreenFlash_IOPort '�̵�
End Sub
'����������״̬���
Private Sub Timer_StatusQuery_Timer()
On Error Resume Next
    'Exit Sub
    mm = mm + 1
    If mm < TimerStatus Then
        Exit Sub
    End If

    '���ListMsg������
    Do While ListMsg.ListCount > 20
        ListMsg.RemoveItem 0
    Loop

    If TestStateFlag <= 5 Then
        mm = 0
        Exit Sub
    End If

    '��ѯӲ�̿ռ�״̬
    HDDStateQuery
    '��ѯ����������״̬
    TSStateQuery
    '��ѯ����״̬
    NetStateQuery

    mm = 0
End Sub
'������������ѯӲ�̿ռ�״̬
Private Sub HDDStateQuery()
    DoEvents
    If GetHDDState(DBPosition, SpaceAvailable) = 1 Then
        FrmMain.Picture9.Picture = LoadPicture(App.Path & "\img\Green.jpg")
        frmInfo.Picture9.Picture = LoadPicture(App.Path & "\img\Green.jpg")
    Else
        FrmMain.Picture9.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.Picture9.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter DBPosition & "Ӳ�̿��ÿռ䲻��" & CStr(Format(SpaceAvailable / 1024, "##.#")) & "G"
        AddMessage "Ӳ�̿��ÿռ䲻��", True
        'flashBuzzerLamp Lamp_RedLight_IOPort
'        DelayTime 2000
'        oIOCard.OutputController Lamp_RedLight_IOPort, False
'        oIOCard.OutputController rdOutput, False
'        oIOCard.OutputController Lamp_GreenFlash_IOPort, True
    End If

End Sub
'������������ѯ����������״̬
Private Sub TSStateQuery()
    On Error GoTo Error
    DoEvents

    If TestStateFlag <= 5 Then
        Exit Sub
    End If

    oRVT520.ResetResult
    If oRVT520.status = 3 Then
        FrmMain.Picture8.Picture = LoadPicture(App.Path & "\img\Green.jpg")
        frmInfo.Picture8.Picture = LoadPicture(App.Path & "\img\Green.jpg")
    Else
        FrmMain.Picture8.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.Picture8.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter "�Ҳ����������"
        AddMessage "�Ҳ����������", True
        'flashBuzzerLamp Lamp_RedLight_IOPort
'        DelayTime 2000
'        oIOCard.OutputController Lamp_RedLight_IOPort, False
'        oIOCard.OutputController rdOutput, False
'        oIOCard.OutputController Lamp_GreenFlash_IOPort, True
    End If

    oLVT520.ResetResult
    If oLVT520.status = 3 Then
        FrmMain.Picture7.Picture = LoadPicture(App.Path & "\img\Green.jpg")
        frmInfo.Picture7.Picture = LoadPicture(App.Path & "\img\Green.jpg")
    Else
        FrmMain.Picture7.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.Picture7.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter "������������"
        AddMessage "������������", True
        'flashBuzzerLamp Lamp_RedLight_IOPort
'        DelayTime 2000
'        oIOCard.OutputController Lamp_RedLight_IOPort, False
'        oIOCard.OutputController rdOutput, False
'        oIOCard.OutputController Lamp_GreenFlash_IOPort, True
    End If

    Exit Sub
Error:
    LogWritter "��ѯ������״̬����"
End Sub
'������������ѯ����״̬
Private Sub NetStateQuery()
    On Error GoTo Error

    Dim objConn As Connection
    Dim objConnMES As Connection

    DoEvents

    '̽�鱾�����ݿ����״̬
    Set objConn = New Connection
    objConn.ConnectionTimeout = 2
    objConn.Open DBCnnStr
    If objConn.state = adStateOpen Then
        FrmMain.PicNet.Picture = LoadPicture(App.Path & "\img\Green.jpg")
        frmInfo.PicNet.Picture = LoadPicture(App.Path & "\img\Green.jpg")
'            LogWritter "MES���ݿ���������"
        objConn.Close
    Else
        FrmMain.PicNet.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.PicNet.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter "�������ݿ������쳣"
        AddMessage "�������ݿ������쳣", True
        'flashBuzzerLamp Lamp_RedLight_IOPort
'        DelayTime 2000
'        oIOCard.OutputController Lamp_RedLight_IOPort, False
'        oIOCard.OutputController rdOutput, False
'        oIOCard.OutputController Lamp_GreenFlash_IOPort, True
    End If

    Set objConn = Nothing

    If Ping(MES_IP) Then
        FrmMain.PicInd.Picture = LoadPicture(App.Path & "\img\Green.jpg")
        frmInfo.PicInd.Picture = LoadPicture(App.Path & "\img\Green.jpg")
'        LogWritter "��������"

        '̽��MES����״̬
        On Error GoTo ErrMES

        Set objConnMES = New Connection
        objConnMES.ConnectionTimeout = 3
        DoEvents
        objConnMES.Open MESCnnStr
        If objConnMES.state = adStateOpen Then
            FrmMain.Picture6.Picture = LoadPicture(App.Path & "\img\Green.jpg")
            frmInfo.Picture6.Picture = LoadPicture(App.Path & "\img\Green.jpg")
'            LogWritter "MES���ݿ���������"
            objConnMES.Close
        Else
            FrmMain.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
            frmInfo.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
            LogWritter "MES���ݿ������쳣"
            AddMessage "MES���ݿ������쳣", True
            'flashBuzzerLamp Lamp_RedLight_IOPort
'            DelayTime 2000
'            oIOCard.OutputController Lamp_RedLight_IOPort, False
'            oIOCard.OutputController rdOutput, False
'            oIOCard.OutputController Lamp_GreenFlash_IOPort, True
        End If

    Else
        FrmMain.PicInd.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.PicInd.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter "�����쳣"
        AddMessage "�����쳣", True
        'flashBuzzerLamp Lamp_RedLight_IOPort
'        DelayTime 2000
'        oIOCard.OutputController Lamp_RedLight_IOPort, False
'        oIOCard.OutputController rdOutput, False
'        oIOCard.OutputController Lamp_GreenFlash_IOPort, True
        FrmMain.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        frmInfo.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
        LogWritter "MES���ݿ������쳣"
    End If

    Set objConnMES = Nothing

    Exit Sub
ErrMES:
    FrmMain.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
    frmInfo.Picture6.Picture = LoadPicture(App.Path & "\img\Red.jpg")
    LogWritter "MES���ݿ������쳣"
    Set objConnMES = Nothing
    Exit Sub
Error:
    LogWritter "���������ݿ�״̬̽����̳���" & Err.Description
End Sub
'������ϵͳͬ���Ų�������Ϣ
Private Sub Timer_DataSync_Timer()
On Error GoTo Err
    nn = nn + 1

    If nn < TimerN Then
        Exit Sub
    End If

    If TestStateFlag <= 5 Then
        nn = 0
        Exit Sub
    End If

    If Not Ping(MES_IP) Then
        nn = 0
        Exit Sub
    End If

    Dim objConn As Connection
    Dim objConnMES As Connection
    Dim objRs As Recordset
    Dim objTmpRs As Recordset
    Dim objRsMES As Recordset
    Dim strSQL As String

    LogWritter "�����Զ�ͬ���Ų���������"

    On Error GoTo ErrMES
    '�ȶ�ȡMES�ϵ�����
    Set objConnMES = New Connection
    Set objRsMES = New Recordset
    objConnMES.ConnectionTimeout = 3
    DoEvents
    objConnMES.Open MESCnnStr
    If objConnMES.state <> adStateOpen Then
        LogWritter "MES���ݿ�����ʧ�ܣ��޷�ͬ������"
        Set objConnMES = Nothing
        Exit Sub
    End If
    strSQL = "select * from mesprd.IF_VEHICLE_TPMS_INFO where tpms_process=0 order by pa_off_seq asc"
    objRsMES.Open strSQL, objConnMES, adOpenKeyset, adLockOptimistic

    '�򿪱������ݿ�����
    Set objConn = New Connection
    Set objRs = New Recordset
    objConn.ConnectionTimeout = 2
    objConn.Open DBCnnStr

    strSQL = "select * from vinlist"
    objRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
    DoEvents
    Set objTmpRs = New Recordset
    Do While Not objRsMES.EOF              '---���������

        strSQL = "select * from vinlist where vin='" & objRsMES("vin") & "'"
        objTmpRs.Open strSQL, objConn, adOpenStatic, adLockOptimistic
        If objTmpRs.EOF Then
            objRs.AddNew
            objRs("vin") = objRsMES("vin")
            objRs!mtoc = objRsMES!mtoc
            objRs!pa_off_seq = objRsMES!pa_off_seq
            objRs!pa_off_time = objRsMES!pa_off_time
            objRs!createtime = Now()
            objRs.Update
        Else
            objTmpRs!mtoc = objRsMES!mtoc
            objTmpRs!pa_off_seq = objRsMES!pa_off_seq
            objTmpRs!pa_off_time = objRsMES!pa_off_time
            objTmpRs!createtime = Now()
            objTmpRs.Update
        End If

        '����MESϵͳ�����ر�ʶ
        strSQL = "update mesprd.IF_VEHICLE_TPMS_INFO set tpms_process=1 where vin='" & objRsMES("vin") & "'"
        objConnMES.Execute strSQL

        objRsMES.MoveNext
        objTmpRs.Close
    Loop
    objRs.Close
    objRsMES.Close
    objConn.Close
    objConnMES.Close
    Set objRs = Nothing
    Set objTmpRs = Nothing
    Set objRsMES = Nothing
    Set objConn = Nothing
    Set objConnMES = Nothing

    LogWritter "�Ų���������ͬ�����"

    nn = 0
    Exit Sub
ErrMES:
    LogWritter "MES���ݿ�����ʧ�ܣ��޷�ͬ������"
    Set objConnMES = Nothing
    nn = 0
    Exit Sub
Err:
    LogWritter "����ͬ�����̳���"
    nn = 0
End Sub

'��ʾϵͳ��Ϣ
Public Sub AddMessage(txt As String, Optional isAlert As Boolean = False)

    Me.ListMsg.AddItem "[" & Now & "]" & txt
    If isAlert Then
        frmInfo.txtInfo.ForeColor = &HFF&
        frmInfo.txtInfo.text = txt
    Else
        frmInfo.txtInfo.ForeColor = &H80000002
        frmInfo.txtInfo.text = txt
    End If
    Me.ListMsg.ListIndex = Me.ListMsg.ListCount - 1
End Sub
'��ʼ�����������
Private Sub initFrom(isInitVin As Boolean)
    FrmMain.picLF.Picture = FrmMain.ImageList.ListImages(6).Picture
    frmInfo.picLF.Picture = frmInfo.ImageList.ListImages(6).Picture
    FrmMain.picLR.Picture = FrmMain.ImageList.ListImages(6).Picture
    frmInfo.picLR.Picture = frmInfo.ImageList.ListImages(6).Picture
    FrmMain.picRF.Picture = FrmMain.ImageList.ListImages(6).Picture
    frmInfo.picRF.Picture = frmInfo.ImageList.ListImages(6).Picture
    FrmMain.picRR.Picture = FrmMain.ImageList.ListImages(6).Picture
    frmInfo.picRR.Picture = frmInfo.ImageList.ListImages(6).Picture

    FrmMain.txtLR.text = ""
    FrmMain.lbLRMdl.Caption = ""
    FrmMain.lbLRPre.Caption = ""
    FrmMain.lbLRTemp.Caption = ""
    FrmMain.lbLRBattery.Caption = ""
    FrmMain.lbLRAcSpeed.Caption = ""

    frmInfo.txtLR.text = ""
    frmInfo.lbLRMdl.Caption = ""
    frmInfo.lbLRPre.Caption = ""
    frmInfo.lbLRTemp.Caption = ""
    frmInfo.lbLRBattery.Caption = ""
    frmInfo.lbLRAcSpeed.Caption = ""

    FrmMain.txtLF.text = ""
    FrmMain.lbLFMdl.Caption = ""
    FrmMain.lbLFPre.Caption = ""
    FrmMain.lbLFTemp.Caption = ""
    FrmMain.lbLFBattery.Caption = ""
    FrmMain.lbLFAcSpeed.Caption = ""

    frmInfo.txtLF.text = ""
    frmInfo.lbLFMdl.Caption = ""
    frmInfo.lbLFPre.Caption = ""
    frmInfo.lbLFTemp.Caption = ""
    frmInfo.lbLFBattery.Caption = ""
    frmInfo.lbLFAcSpeed.Caption = ""

    FrmMain.txtRR.text = ""
    FrmMain.lbRRMdl.Caption = ""
    FrmMain.lbRRPre.Caption = ""
    FrmMain.lbRRTemp.Caption = ""
    FrmMain.lbRRBattery.Caption = ""
    FrmMain.lbRRAcSpeed.Caption = ""

    frmInfo.txtRR.text = ""
    frmInfo.lbRRMdl.Caption = ""
    frmInfo.lbRRPre.Caption = ""
    frmInfo.lbRRTemp.Caption = ""
    frmInfo.lbRRBattery.Caption = ""
    frmInfo.lbRRAcSpeed.Caption = ""

    FrmMain.txtRF.text = ""
    FrmMain.lbRFMdl.Caption = ""
    FrmMain.lbRFPre.Caption = ""
    FrmMain.lbRFTemp.Caption = ""
    FrmMain.lbRFBattery.Caption = ""
    FrmMain.lbRFAcSpeed.Caption = ""

    frmInfo.txtRF.text = ""
    frmInfo.lbRFMdl.Caption = ""
    frmInfo.lbRFPre.Caption = ""
    frmInfo.lbRFTemp.Caption = ""
    frmInfo.lbRFBattery.Caption = ""
    frmInfo.lbRFAcSpeed.Caption = ""

    If isInitVin Then
        txtVIN.text = ""
        frmInfo.labVin.Caption = "̥ѹ����ʼ��ϵͳ"
    End If
End Sub
