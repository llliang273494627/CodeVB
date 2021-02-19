VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form FrmShowLog 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   Picture         =   "FrmShowLog.frx":0000
   ScaleHeight     =   5055
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSACAL.Calendar CalSelect 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8895
      _Version        =   524288
      _ExtentX        =   15690
      _ExtentY        =   7646
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2011
      Month           =   12
      Day             =   10
      DayLength       =   0
      MonthLength     =   0
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   8.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image ImgClose 
      Height          =   285
      Left            =   8520
      Picture         =   "FrmShowLog.frx":127D5
      Top             =   120
      Width           =   285
   End
   Begin VB.Label LbTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "日志查询"
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
      TabIndex        =   1
      Top             =   25
      Width           =   7575
   End
End
Attribute VB_Name = "FrmShowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CalSelect_Click()
    Dim strLogPath As String
    
    strLogPath = AppPath & "Log" & "\"
    If Right(strLogPath, 1) <> "\" Then
        strLogPath = strLogPath & "\"
    End If
    strLogPath = Replace(strLogPath, "\\", "\")
    If DateDiff("d", CalSelect.value, Date) < 0 Then
        PopMsg "提示", "对不起,没有" & CalSelect.value & "日的日志!"
    Else
        Shell "notepad " & strLogPath & Format(CalSelect.value, "yyyy-mm-dd") & ".txt", vbNormalFocus
    End If
End Sub

Private Sub Form_Load()
    CalSelect.value = Now
End Sub

Private Sub ImgClose_Click()
    Unload Me
End Sub
