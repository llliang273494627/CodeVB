VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmHistory 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   Picture         =   "FrmHistory.frx":0000
   ScaleHeight     =   7605
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComCtl2.DTPicker DpEnd 
      Height          =   375
      Left            =   6390
      TabIndex        =   11
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   72679425
      CurrentDate     =   40887
   End
   Begin MSComCtl2.DTPicker DpBegin 
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   72679425
      CurrentDate     =   40887
   End
   Begin VB.TextBox TxtVin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2040
      TabIndex        =   1
      Top             =   810
      Width           =   4815
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   8705
      _Version        =   393216
      Rows            =   20
      Cols            =   28
      Redraw          =   -1  'True
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "历史数据查询"
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
      TabIndex        =   14
      Top             =   25
      Width           =   7575
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   10680
      Picture         =   "FrmHistory.frx":14EAB
      Top             =   120
      Width           =   285
   End
   Begin VB.Label lbNow 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   13
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   12
      Top             =   7200
      Width           =   135
   End
   Begin VB.Image ImgSearch 
      Height          =   465
      Left            =   9120
      Picture         =   "FrmHistory.frx":15323
      Top             =   1560
      Width           =   1515
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "结束日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   9
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "起始日期"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label LbCount 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label LbEnd 
      BackStyle       =   0  'Transparent
      Caption         =   "尾 页"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4560
      TabIndex        =   6
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label LbFirst 
      BackStyle       =   0  'Transparent
      Caption         =   "首 页"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   7200
      Width           =   495
   End
   Begin VB.Label LbBefore 
      BackStyle       =   0  'Transparent
      Caption         =   "上一页"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3000
      TabIndex        =   4
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label LbNext 
      BackStyle       =   0  'Transparent
      Caption         =   "下一页"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   7200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "VIN:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   850
      Width           =   735
   End
End
Attribute VB_Name = "FrmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strVin As String
Dim timeBegin As String
Dim timeEnd As String
Dim PageNow As Long
Dim PageMax As Long
Dim strSql As String
Const IPage = 20 '每页显示数据

Private Sub Form_Load()
    LbFirst.Enabled = False
    LbEnd.Enabled = False
    LbNext.Enabled = False
    LbBefore.Enabled = False
    
    DpBegin.value = DateAdd("d", -7, Now)
    DpEnd.value = Now
End Sub

Private Sub Image1_Click()
    Unload Me
End Sub

Private Sub ImgSearch_Click()

    strVin = Trim(UCase(TxtVin.Text))
    timeBegin = Trim(str(DpBegin.value))
    timeEnd = Trim(str(DpEnd.value))
    
    If strVin = "" Then
        strSql = "select ""ID"", ""VIN"",""VIS"",""ID020"" as ""右前轮ID"",""Mdl020"" as ""右前轮模式"",""Pre020"" as ""右前轮压力"",""Temp020"" as ""右前轮温度"",""Battery020"" as ""右前轮电池"",""AcSpeed020"" as ""右前轮加速度"" ,""ID021"" as ""右后轮ID"",""Mdl021"" as ""右后轮模式"",""Pre021"" as ""右后轮压力"",""Temp021"" as ""右后轮温度"",""Battery021"" as ""右后轮电池"",""AcSpeed021"" as ""右后轮加速度"" ,""ID022"" as ""左前轮ID"",""Mdl022"" as ""左前轮模式"",""Pre022"" as ""左前轮压力"",""Temp022"" as ""左前轮温度"",""Battery022"" as ""左前轮电池"",""AcSpeed022"" as ""左前轮加速度"" ,""ID023"" as ""左后轮ID"" ,""Mdl023"" as ""左后轮模式"",""Pre023"" as ""左后轮压力"",""Temp023"" as ""左后轮温度"",""Battery023"" as ""左后轮电池"",""AcSpeed023"" as ""左后轮加速度"" ,""TestTime"" as ""测试时间"" from ""T_Result"" where ""TestTime"">'" & timeBegin & "' and ""TestTime""<'" & timeEnd & "'"
    Else
        strSql = "select ""ID"", ""VIN"",""VIS"",""ID020"" as ""右前轮ID"",""Mdl020"" as ""右前轮模式"",""Pre020"" as ""右前轮压力"",""Temp020"" as ""右前轮温度"",""Battery020"" as ""右前轮电池"",""AcSpeed020"" as ""右前轮加速度"" ,""ID021"" as ""右后轮ID"",""Mdl021"" as ""右后轮模式"",""Pre021"" as ""右后轮压力"",""Temp021"" as ""右后轮温度"",""Battery021"" as ""右后轮电池"",""AcSpeed021"" as ""右后轮加速度"" ,""ID022"" as ""左前轮ID"",""Mdl022"" as ""左前轮模式"",""Pre022"" as ""左前轮压力"",""Temp022"" as ""左前轮温度"",""Battery022"" as ""左前轮电池"",""AcSpeed022"" as ""左前轮加速度"" ,""ID023"" as ""左后轮ID"" ,""Mdl023"" as ""左后轮模式"",""Pre023"" as ""左后轮压力"",""Temp023"" as ""左后轮温度"",""Battery023"" as ""左后轮电池"",""AcSpeed023"" as ""左后轮加速度"" ,""TestTime"" as ""测试时间"" from ""T_Result"" where ""TestTime"">'" & timeBegin & "' and ""TestTime""<'" & timeEnd & "' and ""VIN"" like '%" & strVin & "%'"
    End If
    MSFGPullFY MSFlexGrid1, strSql, IPage, 1
    LbFirst.Enabled = True
    LbEnd.Enabled = True
    LbNext.Enabled = True
    LbBefore.Enabled = True
End Sub


'填充指定表单的内容MSFlexG表单名称，strSql查询语句，PageSize每页记录数，PageN显示指定页
Public Function MSFGPullFY(MSFlexG As MSFlexGrid, strSql As String, PageSize As Integer, PageN As Long)
    Dim i As Integer
    Dim Tmpi, Tmpj As Long
    Dim tmpRs As ADODB.Recordset
    Set tmpRs = New ADODB.Recordset
    tmpRs.Open strSql, LocalDBConnStr, adOpenKeyset, adLockReadOnly, adCmdText
    
    MSFlexG.TextMatrix(0, 0) = "编号"
    MSFlexG.ColWidth(0) = 500
    For i = 1 To tmpRs.Fields.Count - 1
        MSFlexG.TextMatrix(0, i) = tmpRs.Fields(i).Name
        MSFlexG.ColWidth(i) = 1000
    Next
    MSFlexG.ColWidth(1) = 1800
    MSFlexG.ColWidth(27) = 1800
    
    If tmpRs.BOF And tmpRs.EOF Then
        PopMsg "查询提示", "未查询到满足以下条件的信息" & vbCrLf & "时间段:" & timeBegin & " 到 " & timeEnd & vbCrLf & "VIN:" & strVin
        MSFlexG.Clear
    Else
        i = 1
        tmpRs.PageSize = PageSize
        PageMax = tmpRs.PageCount
        LbCount.Caption = PageMax
        tmpRs.MoveLast
        tmpRs.MoveFirst
        MSFlexG.Rows = tmpRs.PageSize + 1
        MSFlexG.Cols = tmpRs.Fields.Count
'        MSFlexGrid1.CellWidth = 10 '////
        tmpRs.AbsolutePage = PageN
        lbNow.Caption = Trim(str(PageN))
        For Tmpi = 1 To tmpRs.PageSize
            If tmpRs.BOF Or tmpRs.EOF Then
                Exit For
            End If
            MSFlexG.TextMatrix(Tmpi, 0) = str(Tmpi)

            For Tmpj = 1 To tmpRs.Fields.Count
                If IsNull(tmpRs.Fields(Tmpj - 1).value) Then        '控制填充内容不能为空
                    MSFlexG.TextMatrix(Tmpi, Tmpj - 1) = ""
                Else
                    MSFlexG.TextMatrix(Tmpi, Tmpj - 1) = Trim(tmpRs.Fields(Tmpj - 1).value)
                End If
            Next Tmpj
            
            MSFlexG.TextMatrix(i, 0) = i
            i = i + 1
            
            tmpRs.MoveNext
        Next Tmpi
        tmpRs.Close
        Set tmpRs = Nothing
    End If
    
    MSFlexG.Refresh
End Function



Private Sub LbBefore_Click()
    If PageNow > 1 Then
        PageNow = PageNow - 1
        lbNow.Caption = Trim(str(PageNow))
        Call MSFGPullFY(MSFlexGrid1, strSql, IPage, PageNow)
    End If
End Sub

Private Sub LbEnd_Click()
    PageNow = PageMax
    lbNow.Caption = Trim(str(PageNow))
    Call MSFGPullFY(MSFlexGrid1, strSql, IPage, PageNow)
End Sub

Private Sub LbFirst_Click()
    PageNow = 1
    lbNow.Caption = Trim(str(PageNow))
    Call MSFGPullFY(MSFlexGrid1, strSql, IPage, PageNow)
End Sub

Private Sub LbNext_Click()
    If PageNow < PageMax Then
        PageNow = PageNow + 1
        lbNow.Caption = Trim(str(PageNow))
        Call MSFGPullFY(MSFlexGrid1, strSql, IPage, PageNow)
    End If
End Sub
