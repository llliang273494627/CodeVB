VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmDateZone 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ѡ�񵼳�ʱ��"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmDateZone.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4680
   StartUpPosition =   2  '��Ļ����
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   600
      Top             =   2160
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1950
      Width           =   915
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "����"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   1950
      Width           =   915
   End
   Begin MSComCtl2.DTPicker dtpLow 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   840
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   39871
   End
   Begin MSComCtl2.DTPicker dtpHigh 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1380
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   556
      _Version        =   393216
      Format          =   21430273
      CurrentDate     =   39871
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "ѡ�񵼳�����ֹ����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   90
      Width           =   4635
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000009&
      Caption         =   "ֹ����"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   540
      TabIndex        =   3
      Top             =   1440
      Width           =   915
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "������"
      BeginProperty Font 
         Name            =   "����_GB2312"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   540
      TabIndex        =   2
      Top             =   870
      Width           =   885
   End
End
Attribute VB_Name = "frmDateZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'** �ļ�����frmDateZone.frm
'** ��  Ȩ��CopyRight (c) 2008-2010 �人��������ϵͳ���޹�˾
'** �����ˣ�yangshuai
'** ��  �䣺shuaigoplay@live.cn
'** ��  �ڣ�2009-2-27
'** �޸��ˣ�
'** ��  �ڣ�
'** ��  ����ʱ��ѡ��Ի���
'** ��  ����1.0
'******************************************************************************

Option Explicit
'******************************************************************************
'** �� �� ����cmdCancel
'** ��    �룺
'** ��    ����
'** ����������ȡ����ť�¼���Ӧ
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'******************************************************************************
'** �� �� ����cmdSaveAs_Click
'** ��    �룺
'** ��    ����
'** ����������������ť�¼���Ӧ
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Private Sub cmdSaveAs_Click()
    Dim sqlText As String
    Dim lowDate
    Dim highDate
    
    lowDate = CDate(Me.dtpLow.value)
    highDate = CDate(Me.dtpHigh.value)
    
    If lowDate > highDate Then
        MsgBox " "
        Exit Sub
    End If
    
    sqlText = "select ""VIN"",""VIS"",""ID020"" as ""��ǰ��ID"",""Mdl020"" as ""��ǰ��ģʽ"",""Pre020"" as ""��ǰ��ѹ��"",""Temp020"" as ""��ǰ���¶�"",""Battery020"" as ""��ǰ�ֵ��"",""AcSpeed020"" as ""��ǰ�ּ��ٶ�"" ,""ID021"" as ""�Һ���ID"",""Mdl021"" as ""�Һ���ģʽ"",""Pre021"" as ""�Һ���ѹ��"",""Temp021"" as ""�Һ����¶�"",""Battery021"" as ""�Һ��ֵ��"",""AcSpeed021"" as ""�Һ��ּ��ٶ�"" ,""ID022"" as ""��ǰ��ID"",""Mdl022"" as ""��ǰ��ģʽ"",""Pre022"" as ""��ǰ��ѹ��"",""Temp022"" as ""��ǰ���¶�"",""Battery022"" as ""��ǰ�ֵ��"",""AcSpeed022"" as ""��ǰ�ּ��ٶ�"" ,""ID023"" as ""�����ID"" ,""Mdl023"" as ""�����ģʽ"",""Pre023"" as ""�����ѹ��"",""Temp023"" as ""������¶�"",""Battery023"" as ""����ֵ��"",""AcSpeed023"" as ""����ּ��ٶ�"" ,""TestTime"" as ""����ʱ��"",""WriteInTime"" as ""д��ʱ��"" from " _
    & " ""T_Result"" where   ""TestTime"">='" & lowDate & "' and ""TestTime""<='" & highDate & "'"
    
    '��ϵ�����ѯ��䣬���õ�������
    
    exportExcel sqlText
    

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
    '�ؼ����XP��
    WindowsXPC1.InitSubClassing
    
    '����ؼ�����
    dtpLow.value = DateAdd("d", -7, Date)
    dtpHigh.value = Date
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub






