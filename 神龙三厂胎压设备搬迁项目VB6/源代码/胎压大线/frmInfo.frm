VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.4#0"; "COMCTL32.OCX"
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC.ocx"
Begin VB.Form frmInfo 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   11520
   ClientLeft      =   18960
   ClientTop       =   0
   ClientWidth     =   19200
   LinkTopic       =   "Form1"
   Picture         =   "frmInfo.frx":0000
   ScaleHeight     =   11520
   ScaleWidth      =   19200
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   100
      Picture         =   "frmInfo.frx":469FC
      ScaleHeight     =   450
      ScaleWidth      =   4485
      TabIndex        =   75
      Top             =   0
      Width           =   4485
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   150
      Picture         =   "frmInfo.frx":4C6F3
      ScaleHeight     =   420
      ScaleWidth      =   645
      TabIndex        =   74
      Top             =   11040
      Width           =   645
   End
   Begin VB.PictureBox PicInd 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1800
      Picture         =   "frmInfo.frx":4D2AE
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   25
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox PicNet 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3960
      Picture         =   "frmInfo.frx":4D9A8
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   24
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   8880
      Picture         =   "frmInfo.frx":4E0A2
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   23
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   14520
      Picture         =   "frmInfo.frx":4E79C
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   22
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   12000
      Picture         =   "frmInfo.frx":4EE96
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   21
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   6360
      Picture         =   "frmInfo.frx":4F590
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   20
      Top             =   10410
      Width           =   420
   End
   Begin VB.PictureBox picRF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   15240
      Picture         =   "frmInfo.frx":4FC8A
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   18
      Top             =   8820
      Width           =   420
   End
   Begin VB.TextBox txtRF 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16455
      TabIndex        =   17
      Top             =   8820
      Width           =   2085
   End
   Begin VB.PictureBox picRR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7830
      Picture         =   "frmInfo.frx":52302
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   15
      Top             =   8820
      Width           =   420
   End
   Begin VB.TextBox txtRR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9030
      TabIndex        =   14
      Top             =   8820
      Width           =   2085
   End
   Begin VB.PictureBox picLF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   15270
      Picture         =   "frmInfo.frx":5497A
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   12
      Top             =   3930
      Width           =   420
   End
   Begin VB.TextBox txtLF 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   16545
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3930
      Width           =   1995
   End
   Begin VB.PictureBox picLR 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   7860
      Picture         =   "frmInfo.frx":56FF2
      ScaleHeight     =   420
      ScaleWidth      =   420
      TabIndex        =   9
      Top             =   3930
      Width           =   420
   End
   Begin VB.TextBox txtLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9150
      TabIndex        =   8
      Top             =   3930
      Width           =   1965
   End
   Begin VB.TextBox txtInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1050
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2220
      Width           =   18465
   End
   Begin VB.ListBox ListInput 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   33.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5460
      ItemData        =   "frmInfo.frx":5966A
      Left            =   3960
      List            =   "frmInfo.frx":5966C
      TabIndex        =   1
      Top             =   4410
      Width           =   3165
   End
   Begin VB.ListBox ListOutput 
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   33.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   5460
      ItemData        =   "frmInfo.frx":5966E
      Left            =   360
      List            =   "frmInfo.frx":59670
      TabIndex        =   0
      Top             =   4410
      Width           =   3180
   End
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   9120
      Top             =   22800
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "华信数据"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Left            =   870
      TabIndex        =   73
      Top             =   11100
      Width           =   1410
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "武汉市洪山区珞瑜东路佳园路光谷国际A座2318室    电话：027-87775236"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10950
      TabIndex        =   72
      Top             =   11130
      Width           =   7875
   End
   Begin VB.Label lbRFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   18120
      TabIndex        =   71
      Top             =   9330
      Width           =   930
   End
   Begin VB.Label lbRFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   16740
      TabIndex        =   70
      Top             =   9330
      Width           =   930
   End
   Begin VB.Label lbRFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   15750
      TabIndex        =   69
      Top             =   9330
      Width           =   540
   End
   Begin VB.Label lbRFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   15750
      TabIndex        =   68
      Top             =   9570
      Width           =   510
   End
   Begin VB.Label lbRFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   16950
      TabIndex        =   67
      Top             =   9570
      Width           =   1410
   End
   Begin VB.Label lbRRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9630
      TabIndex        =   66
      Top             =   9570
      Width           =   1410
   End
   Begin VB.Label lbRRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8400
      TabIndex        =   65
      Top             =   9570
      Width           =   510
   End
   Begin VB.Label lbRRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8400
      TabIndex        =   64
      Top             =   9330
      Width           =   540
   End
   Begin VB.Label lbRRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9390
      TabIndex        =   63
      Top             =   9330
      Width           =   930
   End
   Begin VB.Label lbRRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   10830
      TabIndex        =   62
      Top             =   9330
      Width           =   930
   End
   Begin VB.Label lbLFAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   16980
      TabIndex        =   61
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Label lbLFBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   15750
      TabIndex        =   60
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label lbLFMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   15750
      TabIndex        =   59
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label lbLFPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   16770
      TabIndex        =   58
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label lbLFTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   18120
      TabIndex        =   57
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label Label39 
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7830
      TabIndex        =   56
      Top             =   4440
      Width           =   720
   End
   Begin VB.Label lbLRTemp 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   10740
      TabIndex        =   55
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label lbLRPre 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9360
      TabIndex        =   54
      Top             =   4440
      Width           =   930
   End
   Begin VB.Label lbLRMdl 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8400
      TabIndex        =   53
      Top             =   4440
      Width           =   540
   End
   Begin VB.Label lbLRBattery 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   8400
      TabIndex        =   52
      Top             =   4680
      Width           =   510
   End
   Begin VB.Label lbLRAcSpeed 
      BackStyle       =   0  'Transparent
      Caption         =   "123"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   9570
      TabIndex        =   51
      Top             =   4680
      Width           =   1410
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   15210
      TabIndex        =   50
      Top             =   9330
      Width           =   1140
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "压力："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   16170
      TabIndex        =   49
      Top             =   9330
      Width           =   1200
   End
   Begin VB.Label Label30 
      BackStyle       =   0  'Transparent
      Caption         =   "温度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   17550
      TabIndex        =   48
      Top             =   9330
      Width           =   1200
   End
   Begin VB.Label Label29 
      BackStyle       =   0  'Transparent
      Caption         =   "加速度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   16170
      TabIndex        =   47
      Top             =   9570
      Width           =   1200
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "电池："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   15210
      TabIndex        =   46
      Top             =   9570
      Width           =   1200
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7830
      TabIndex        =   45
      Top             =   9330
      Width           =   1140
   End
   Begin VB.Label Label26 
      BackStyle       =   0  'Transparent
      Caption         =   "压力："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8820
      TabIndex        =   44
      Top             =   9330
      Width           =   1200
   End
   Begin VB.Label Label25 
      BackStyle       =   0  'Transparent
      Caption         =   "温度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10260
      TabIndex        =   43
      Top             =   9330
      Width           =   1200
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "加速度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8820
      TabIndex        =   42
      Top             =   9570
      Width           =   1200
   End
   Begin VB.Label Label23 
      BackStyle       =   0  'Transparent
      Caption         =   "电池："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7830
      TabIndex        =   41
      Top             =   9570
      Width           =   1200
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "模式："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   15180
      TabIndex        =   40
      Top             =   4440
      Width           =   1140
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
      Caption         =   "压力："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   16170
      TabIndex        =   39
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "温度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   17550
      TabIndex        =   38
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "加速度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   16170
      TabIndex        =   37
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "电池："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   15180
      TabIndex        =   36
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "压力："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8790
      TabIndex        =   35
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "温度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   10170
      TabIndex        =   34
      Top             =   4440
      Width           =   1200
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "加速度："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   8790
      TabIndex        =   33
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "电池："
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   7830
      TabIndex        =   32
      Top             =   4680
      Width           =   1200
   End
   Begin VB.Label labNow 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1305
      Left            =   9480
      TabIndex        =   2
      Top             =   5370
      Width           =   6705
   End
   Begin ComctlLib.ImageList ImageList 
      Left            =   870
      Top             =   720
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
            Picture         =   "frmInfo.frx":59672
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInfo.frx":59FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInfo.frx":5A976
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInfo.frx":5B2F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInfo.frx":5BC7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmInfo.frx":5C5FC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "网络连接"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   2295
      TabIndex        =   31
      Top             =   10485
      Width           =   1335
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "本地数据库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   4455
      TabIndex        =   30
      Top             =   10485
      Width           =   1575
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "数据库硬盘容量"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   9375
      TabIndex        =   29
      Top             =   10485
      Width           =   2175
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "右侧控制器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15015
      TabIndex        =   28
      Top             =   10485
      Width           =   1575
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "左侧控制器"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12495
      TabIndex        =   27
      Top             =   10485
      Width           =   1575
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SPPV数据库"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   6885
      TabIndex        =   26
      Top             =   10485
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "右前"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   15705
      TabIndex        =   19
      Top             =   8835
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "右后"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   8250
      TabIndex        =   16
      Top             =   8835
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "左前"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   15735
      TabIndex        =   13
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "左后"
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   8340
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "排产队列"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   4080
      TabIndex        =   6
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "扫描队列"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Label labVin 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "胎压检测初始化系统"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   18150
   End
   Begin VB.Label labNext 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "黑体"
         Size            =   72
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1335
      Left            =   9480
      TabIndex        =   3
      Top             =   6720
      Width           =   6735
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



