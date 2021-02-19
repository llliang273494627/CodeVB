VERSION 5.00
Begin VB.Form FrmMsg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   Picture         =   "FrmMsg.frx":0000
   ScaleHeight     =   5025
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrClose 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8400
      Top             =   4320
   End
   Begin VB.Label LbOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3830
      TabIndex        =   2
      Top             =   3770
      Width           =   1430
   End
   Begin VB.Image ImgClose 
      Height          =   285
      Left            =   8530
      Picture         =   "FrmMsg.frx":13294
      Top             =   120
      Width           =   285
   End
   Begin VB.Label LbMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   8535
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
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
      Left            =   100
      TabIndex        =   0
      Top             =   60
      Width           =   7575
   End
End
Attribute VB_Name = "FrmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'**** File Name：FrmMsg.frm
'**** Function：用于弹出消息
'**** Author: Jiangxueqiao
'**** Email: jiangxueqiao@foxmail.com
'**** Time: 2011/12/09
'*****************************************************************************

Dim CloseTime As Integer '窗体自动关闭时间

Private Sub Form_Load()
    CloseTime = GetIniN("App", "PopMsgStayTime", 0, AppPath & "setting.ini")
    If CloseTime = 0 Then
        LbOK.Caption = "确定"
    Else
        LbOK.Caption = "确定 (" & Trim(Str(CloseTime)) & ")"
        TmrClose.Enabled = True
    End If
End Sub

Private Sub ImgClose_Click()
    Unload Me
End Sub

Private Sub LbOK_Click()
    Unload Me
End Sub

Private Sub TmrClose_Timer()
    CloseTime = CloseTime - 1
    LbOK.Caption = "确定 (" & Trim(Str(CloseTime)) & ")"
    If CloseTime = 0 Then
        Unload Me
    End If
End Sub
