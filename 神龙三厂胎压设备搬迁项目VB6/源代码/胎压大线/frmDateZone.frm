VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D1C90141-3FBE-4464-B25B-D4CA17FB66F3}#1.0#0"; "WinXPC Engine.ocx"
Begin VB.Form frmDateZone 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择导出时间"
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
   StartUpPosition =   2  '屏幕中心
   Begin WinXPC_Engine.WindowsXPC WindowsXPC1 
      Left            =   600
      Top             =   2160
      _ExtentX        =   6588
      _ExtentY        =   1085
      ColorScheme     =   2
      Common_Dialog   =   0   'False
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   315
      Left            =   1680
      TabIndex        =   6
      Top             =   1950
      Width           =   915
   End
   Begin VB.CommandButton cmdSaveAs 
      Caption         =   "导出"
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
      Caption         =   "选择导出的起止日期"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "止日期"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
      Caption         =   "起日期"
      BeginProperty Font 
         Name            =   "楷体_GB2312"
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
'** 文件名：frmDateZone.frm
'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
'** 创建人：yangshuai
'** 邮  箱：shuaigoplay@live.cn
'** 日  期：2009-2-27
'** 修改人：
'** 日  期：
'** 描  述：时间选择对话框
'** 版  本：1.0
'******************************************************************************

Option Explicit
'******************************************************************************
'** 函 数 名：cmdCancel
'** 输    入：
'** 输    出：
'** 功能描述：取消按钮事件响应
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'******************************************************************************
'** 函 数 名：cmdSaveAs_Click
'** 输    入：
'** 输    出：
'** 功能描述：导出按钮事件响应
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
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
    
    sqlText = "select ""VIN"",""VIS"",""ID020"" as ""右前轮ID"",""Mdl020"" as ""右前轮模式"",""Pre020"" as ""右前轮压力"",""Temp020"" as ""右前轮温度"",""Battery020"" as ""右前轮电池"",""AcSpeed020"" as ""右前轮加速度"" ,""ID021"" as ""右后轮ID"",""Mdl021"" as ""右后轮模式"",""Pre021"" as ""右后轮压力"",""Temp021"" as ""右后轮温度"",""Battery021"" as ""右后轮电池"",""AcSpeed021"" as ""右后轮加速度"" ,""ID022"" as ""左前轮ID"",""Mdl022"" as ""左前轮模式"",""Pre022"" as ""左前轮压力"",""Temp022"" as ""左前轮温度"",""Battery022"" as ""左前轮电池"",""AcSpeed022"" as ""左前轮加速度"" ,""ID023"" as ""左后轮ID"" ,""Mdl023"" as ""左后轮模式"",""Pre023"" as ""左后轮压力"",""Temp023"" as ""左后轮温度"",""Battery023"" as ""左后轮电池"",""AcSpeed023"" as ""左后轮加速度"" ,""TestTime"" as ""测试时间"",""WriteInTime"" as ""写入时间"" from " _
    & " ""T_Result"" where   ""TestTime"">='" & lowDate & "' and ""TestTime""<='" & highDate & "'"
    
    '组合导出查询语句，调用导出函数
    
    exportExcel sqlText
    

End Sub



'******************************************************************************
'** 函 数 名：Form_Load
'** 输    入：
'** 输    出：
'** 功能描述：窗体加载时间响应
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub Form_Load()
    '控件风格XP化
    WindowsXPC1.InitSubClassing
    
    '界面控件控制
    dtpLow.value = DateAdd("d", -7, Date)
    dtpHigh.value = Date
    
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
End Sub






