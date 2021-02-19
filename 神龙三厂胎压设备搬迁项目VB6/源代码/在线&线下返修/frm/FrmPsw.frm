VERSION 5.00
Begin VB.Form FrmPsw 
   BorderStyle     =   0  'None
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   Picture         =   "FrmPsw.frx":0000
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Timer TmrClose 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   4320
   End
   Begin VB.TextBox TxtPsw 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1680
      Width           =   6975
   End
   Begin VB.Label LbInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   480
      TabIndex        =   4
      Top             =   2520
      Width           =   8295
   End
   Begin VB.Image ImgClose 
      Height          =   285
      Left            =   8520
      Picture         =   "FrmPsw.frx":13294
      Top             =   120
      Width           =   285
   End
   Begin VB.Label LbOK 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ��"
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
      TabIndex        =   3
      Top             =   3770
      Width           =   1425
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "����������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
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
Attribute VB_Name = "FrmPsw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************
'**** File Name��FrmPsw.frm
'**** Function�����ڵ������������
'**** Author: Jiangxueqiao
'**** Email: jiangxueqiao@foxmail.com
'**** Time: 2011/12/09
'*****************************************************************************

Dim CloseTime As Integer '�����Զ��ر�ʱ��
Const ConCloseTime = 10
Private Sub Form_Load()
    LbTitle.Caption = AppName
    CloseTime = ConCloseTime '�Ի����Զ��ر�ʱ��̶�Ϊ10��
End Sub

Private Sub ImgClose_Click()
    Unload Me
End Sub



Private Sub LbOK_Click()
    Dim x As Form
    Dim strpsw As String
    
    strpsw = GetIniS("App", "Psw", "", AppPath & "Setting.ini")
    '///// ��������87775236�����ں���ά��
    If Trim(TxtPsw.Text) = "87775236" Or strpsw = Trim(TxtPsw.Text) Then
        If PswMode = "exit" Then '�˳�
'            Call frmMain.CloseAllIO '�ر�IO���
'            Set oIOCard = Nothing
'            For Each x In Forms
'                Unload x
'            Next

            frmMain.isQuit = True
            Unload Me
            End
        ElseIf PswMode = "option" Then 'ϵͳ����
            FrmOption.Show 1
            Unload Me
        End If
    Else
        If CloseTime > 0 Then '���ڶԻ����Զ��ر�ʱ��̶����˴��ж���ʱ���ã�������չ
            TmrClose.Enabled = True
            LbInfo.Caption = "�������,�������" & Trim(str(CloseTime)) & "���رգ�"
        Else
            TmrClose.Enabled = False
            LbInfo.Caption = "�������!"
        End If
    End If
    PswMode = ""
End Sub


Private Sub TxtPsw_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LbOK_Click
    Else
        LbInfo.Caption = ""
        TmrClose.Enabled = False
        CloseTime = ConCloseTime
    End If
End Sub

Private Sub TmrClose_Timer()
    CloseTime = CloseTime - 1
    LbInfo.Caption = "�������,�������" & Trim(str(CloseTime)) & "���رգ�"
    If CloseTime = 0 Then
        Unload Me
    End If
End Sub

