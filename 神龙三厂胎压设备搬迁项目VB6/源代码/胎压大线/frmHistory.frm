VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmHistory 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ʾ��ʷ��¼"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   Icon            =   "frmHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����  "
      Height          =   5745
      Left            =   8880
      TabIndex        =   2
      Top             =   2580
      Width           =   2895
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1920
         Width           =   885
      End
      Begin VB.CommandButton Command3 
         Caption         =   "ȡ    ��"
         Height          =   435
         Left            =   720
         TabIndex        =   11
         Top             =   3600
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��    ��"
         Height          =   435
         Left            =   720
         TabIndex        =   9
         Top             =   2760
         Width           =   1545
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "��һҳ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   750
         TabIndex        =   13
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "��һҳ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   750
         TabIndex        =   12
         Top             =   1260
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�����ʾ    "
      Height          =   5745
      Left            =   90
      TabIndex        =   1
      Top             =   2580
      Width           =   8685
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   5235
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   9234
         _Version        =   393216
         Rows            =   10
         Cols            =   10
         BackColor       =   16777215
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         FocusRect       =   0
         AllowUserResizing=   3
         Appearance      =   0
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "����ѡ��    "
      Height          =   2355
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   11715
      Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
         Left            =   8880
         Top             =   600
         _ExtentX        =   6588
         _ExtentY        =   1085
         ColorScheme     =   2
         Common_Dialog   =   0   'False
      End
      Begin VB.CommandButton Command4 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   9660
         TabIndex        =   15
         Top             =   1290
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpHigh 
         Height          =   375
         Left            =   6000
         TabIndex        =   8
         Top             =   1290
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   62914563
         CurrentDate     =   39872
         MaxDate         =   55153
      End
      Begin MSComCtl2.DTPicker dtpLow 
         Height          =   375
         Left            =   2370
         TabIndex        =   7
         Top             =   1290
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         Format          =   62914563
         CurrentDate     =   39872
         MaxDate         =   55153
      End
      Begin VB.TextBox txtVIN 
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2370
         MaxLength       =   17
         TabIndex        =   4
         Top             =   450
         Width           =   5535
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ʼ����"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   720
         TabIndex        =   6
         Top             =   1260
         Width           =   1635
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ֹ����"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   4350
         TabIndex        =   5
         Top             =   1290
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VIN"
         BeginProperty Font 
            Name            =   "����_GB2312"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   1590
         TabIndex        =   3
         Top             =   480
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ofy As CFY
Dim SelectMember As String
Dim nowPage As Long
Dim rs As ADODB.Recordset



Private Sub Combo1_Click()
    Dim SqlStr As String
    nowPage = Me.Combo1.text
    ofy.PageNum = nowPage
    ofy.getRecordSet rs
    SqlStr = ofy.SelectSqlStr
    showDataInMSFlexGrid Me.MSFlexGrid1, DBCnnStr, """T_Result""", SqlStr
End Sub

Private Sub Command2_Click()
    Dim WhereMenber As String
    Dim SqlStr As String
    If txtVIN.text <> "" Then
        WhereMenber = " and ""VIN"" like '%" & txtVIN.text & "%' "
    End If
    WhereMenber = WhereMenber & " and  ""TestTime"">='" & Me.dtpLow.value & "' and ""TestTime""<='" & Me.dtpHigh.value & "'"
    SelectMember = " ""ID"", ""VIN"",""VIS"",""ID020"" as ""��ǰ��ID"",""Mdl020"" as ""��ǰ��ģʽ"",""Pre020"" as ""��ǰ��ѹ��"",""Temp020"" as ""��ǰ���¶�"",""Battery020"" as ""��ǰ�ֵ��"",""AcSpeed020"" as ""��ǰ�ּ��ٶ�"" ,""ID021"" as ""�Һ���ID"",""Mdl021"" as ""�Һ���ģʽ"",""Pre021"" as ""�Һ���ѹ��"",""Temp021"" as ""�Һ����¶�"",""Battery021"" as ""�Һ��ֵ��"",""AcSpeed021"" as ""�Һ��ּ��ٶ�"" ,""ID022"" as ""��ǰ��ID"",""Mdl022"" as ""��ǰ��ģʽ"",""Pre022"" as ""��ǰ��ѹ��"",""Temp022"" as ""��ǰ���¶�"",""Battery022"" as ""��ǰ�ֵ��"",""AcSpeed022"" as ""��ǰ�ּ��ٶ�"" ,""ID023"" as ""�����ID"" ,""Mdl023"" as ""�����ģʽ"",""Pre023"" as ""�����ѹ��"",""Temp023"" as ""������¶�"",""Battery023"" as ""����ֵ��"",""AcSpeed023"" as ""����ּ��ٶ�"" ,""TestTime"" as ""����ʱ��"",""WriteInTime"" as ""д��ʱ��"" "

    SqlStr = "select " & SelectMember & " from ""T_Result"" where 1=1 " & WhereMenber
    
    exportExcel SqlStr
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
On Error GoTo select_ERR
    Dim WhereMenber As String
    Dim SqlStr As String
    If txtVIN.text <> "" Then
        WhereMenber = " and ""VIN"" like '%" & txtVIN.text & "%' "
    End If
    WhereMenber = WhereMenber & " and  ""TestTime"">='" & Me.dtpLow.value & "' and ""TestTime""<='" & Me.dtpHigh.value & "'"
    


    Set ofy = New CFY
    ofy.ConnectionString = DBCnnStr
    
    SelectMember = " ""ID"", ""VIN"",""VIS"",""ID020"" as ""��ǰ��ID"",""Mdl020"" as ""��ǰ��ģʽ"",""Pre020"" as ""��ǰ��ѹ��"",""Temp020"" as ""��ǰ���¶�"",""Battery020"" as ""��ǰ�ֵ��"",""AcSpeed020"" as ""��ǰ�ּ��ٶ�"" ,""ID021"" as ""�Һ���ID"",""Mdl021"" as ""�Һ���ģʽ"",""Pre021"" as ""�Һ���ѹ��"",""Temp021"" as ""�Һ����¶�"",""Battery021"" as ""�Һ��ֵ��"",""AcSpeed021"" as ""�Һ��ּ��ٶ�"" ,""ID022"" as ""��ǰ��ID"",""Mdl022"" as ""��ǰ��ģʽ"",""Pre022"" as ""��ǰ��ѹ��"",""Temp022"" as ""��ǰ���¶�"",""Battery022"" as ""��ǰ�ֵ��"",""AcSpeed022"" as ""��ǰ�ּ��ٶ�"" ,""ID023"" as ""�����ID"" ,""Mdl023"" as ""�����ģʽ"",""Pre023"" as ""�����ѹ��"",""Temp023"" as ""������¶�"",""Battery023"" as ""����ֵ��"",""AcSpeed023"" as ""����ּ��ٶ�"" ,""TestTime"" as ""����ʱ��"",""WriteInTime"" as ""д��ʱ��"" "
    ofy.tableName = " ""T_Result"" "
    
    nowPage = 1
    ofy.WhereMenber = WhereMenber
    ofy.KeyField = " ""ID"" "
    ofy.PageNum = nowPage
    ofy.SelectMember = SelectMember
    ofy.getRecordSet rs
    SqlStr = ofy.SelectSqlStr
    showDataInMSFlexGrid Me.MSFlexGrid1, DBCnnStr, """T_Result""", SqlStr
    nowPage = 1
    loadCombo
    Exit Sub
    
select_ERR:
    MsgBox Err.Description
End Sub







Private Sub Form_Load()
    Dim WhereMenber As String
    Dim SqlStr As String
    WindowsXPC1.InitSubClassing
    dtpLow.value = DateAdd("d", -7, Date)
    dtpHigh.value = DateAdd("d", 1, Date)
    


    If txtVIN.text <> "" Then
        WhereMenber = " and ""VIN"" like '" & txtVIN.text & "' "
    End If
    WhereMenber = WhereMenber & " and  ""TestTime"">='" & Me.dtpLow.value & "' and ""TestTime""<='" & Me.dtpHigh.value & "'"
    


    Set ofy = New CFY
    ofy.ConnectionString = DBCnnStr
    
    SelectMember = " ""ID"", ""VIN"",""VIS"",""ID020"" as ""��ǰ��ID"",""Mdl020"" as ""��ǰ��ģʽ"",""Pre020"" as ""��ǰ��ѹ��"",""Temp020"" as ""��ǰ���¶�"",""Battery020"" as ""��ǰ�ֵ��"",""AcSpeed020"" as ""��ǰ�ּ��ٶ�"" ,""ID021"" as ""�Һ���ID"",""Mdl021"" as ""�Һ���ģʽ"",""Pre021"" as ""�Һ���ѹ��"",""Temp021"" as ""�Һ����¶�"",""Battery021"" as ""�Һ��ֵ��"",""AcSpeed021"" as ""�Һ��ּ��ٶ�"" ,""ID022"" as ""��ǰ��ID"",""Mdl022"" as ""��ǰ��ģʽ"",""Pre022"" as ""��ǰ��ѹ��"",""Temp022"" as ""��ǰ���¶�"",""Battery022"" as ""��ǰ�ֵ��"",""AcSpeed022"" as ""��ǰ�ּ��ٶ�"" ,""ID023"" as ""�����ID"" ,""Mdl023"" as ""�����ģʽ"",""Pre023"" as ""�����ѹ��"",""Temp023"" as ""������¶�"",""Battery023"" as ""����ֵ��"",""AcSpeed023"" as ""����ּ��ٶ�"" ,""TestTime"" as ""����ʱ��"",""WriteInTime"" as ""д��ʱ��"" "
    ofy.tableName = " ""T_Result"" "
    
    nowPage = 1
    ofy.WhereMenber = WhereMenber
    ofy.KeyField = " ""ID"" "
    ofy.PageNum = nowPage
    ofy.SelectMember = SelectMember
    ofy.getRecordSet rs
    SqlStr = ofy.SelectSqlStr
    showDataInMSFlexGrid Me.MSFlexGrid1, DBCnnStr, """T_Result""", SqlStr
    nowPage = 1
    loadCombo
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Exit Sub
End Sub
'******************************************************************************
'** �� �� ����showDataInMSFlexGrid
'** ��    �룺
'** ��    ����
'** ������������ʾ���
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************

Public Sub showDataInMSFlexGrid(msFG As MSFlexGrid, CnnStr As String, tableName As String, sql As String)
'On Error GoTo Err_ShowGrid
    msFG.Clear
    If sql = "" Then
        Exit Sub
    End If
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset, rsTmp As New ADODB.Recordset
    Dim i As Integer, J As Integer
   
    cnn.Open CnnStr
    rs.Open sql, cnn, 1, 3
    
    With msFG
        .Visible = True
        .cols = rs.Fields.Count
        .Rows = 55
        .FillStyle = 1
        '.CellAlignment = flexAlignLeftCenter
        For i = 0 To rs.Fields.Count - 1
            .TextMatrix(0, i) = rs.Fields(i).Name
        Next
        J = 1
        Do While Not rs.EOF
            For i = 0 To rs.Fields.Count - 1
                If IsNull(rs(i)) Then
                    .TextMatrix(J, i) = ""
                Else
                    .TextMatrix(J, i) = rs(i)
                    
                End If
            Next
            rs.MoveNext
            J = J + 1
        Loop
    End With
    Call setColWidth(msFG, rs.Fields.Count)  '�����п�������̿��Ը����Լ���Ҫ����
    rs.Close
    Set rs = Nothing
    cnn.Close
    
    Exit Sub
Err_ShowGrid:
    MsgBox "��ʾ���ݳ���������Ϣ��" & Err.Description
End Sub
Private Sub setColWidth(msFG As MSFlexGrid, cols As Integer)
With msFG
    Dim i As Integer
    .ColWidth(0) = 0
    .ColWidth(1) = 2000
    For i = 2 To cols - 2 'Ϊÿ���е��н�������
        .ColWidth(i) = 1150 '�еĿ��,�Ժ��Լ�����
    Next
    .ColWidth(i - 1) = 1600
    .ColWidth(i) = 1600
End With
End Sub








Private Sub Label4_Click()
    If nowPage < ofy.PageCount Then
        Dim SqlStr As String
        nowPage = nowPage + 1
        ofy.PageNum = nowPage
        ofy.getRecordSet rs
        SqlStr = ofy.SelectSqlStr
        showDataInMSFlexGrid Me.MSFlexGrid1, DBCnnStr, """T_Result""", SqlStr
    Else
        MsgBox "�Ѿ���βҳ��"
    End If
    Dim i As Long
    loadCombo
    Exit Sub
End Sub

Private Sub Label5_Click()
    If nowPage > 1 Then
        Dim SqlStr As String
        nowPage = nowPage - 1
        ofy.PageNum = nowPage
        ofy.getRecordSet rs
        SqlStr = ofy.SelectSqlStr
        showDataInMSFlexGrid Me.MSFlexGrid1, DBCnnStr, """T_Result""", SqlStr
    Else
        MsgBox "�Ѿ�����ҳ��"
    End If
    loadCombo
    Exit Sub
End Sub

Public Sub loadCombo()
    Me.Combo1.Clear
    Dim i As Long
    For i = 1 To ofy.PageCount
        Me.Combo1.AddItem i, i - 1
    Next
    If Me.Combo1.ListCount > 0 Then
        Me.Combo1.ListIndex = nowPage - 1
    End If
End Sub
