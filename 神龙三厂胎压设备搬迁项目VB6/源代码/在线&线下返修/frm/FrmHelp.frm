VERSION 5.00
Begin VB.Form FrmHelp 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   Picture         =   "FrmHelp.frx":0000
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "系统：设置系统参数（需要密码验证）"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   2400
      Width           =   8535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "日志：查询系统日志"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   3
      Top             =   1560
      Width           =   7695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "历史：查询历史数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   8535
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   8520
      Picture         =   "FrmHelp.frx":13294
      Top             =   120
      Width           =   285
   End
   Begin VB.Label LbClose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "关闭"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   120
      Picture         =   "FrmHelp.frx":1370C
      Top             =   4440
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9120
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "帮助"
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
      TabIndex        =   0
      Top             =   25
      Width           =   7575
   End
End
Attribute VB_Name = "FrmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub LbClose_Click()
    Unload Me
End Sub
