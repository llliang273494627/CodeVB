VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form frmShowLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "日志查询"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   Icon            =   "frmShowLog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin MSACAL.Calendar DateSelect 
      Height          =   2955
      Left            =   45
      TabIndex        =   0
      Top             =   15
      Width           =   4440
      _Version        =   524288
      _ExtentX        =   7832
      _ExtentY        =   5212
      _StockProps     =   1
      BackColor       =   -2147483639
      Year            =   2011
      Month           =   5
      Day             =   19
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
End
Attribute VB_Name = "frmShowLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DateSelect_DblClick()
'    Debug.Print GetProjectPath() & "Log\" & DateSelect.value & ".txt"
'    Exit Sub

    If DateDiff("d", DateSelect.value, Date) < 0 Then
        MsgBox "对不起没有" & DateSelect.value & "的日志！"
    Else
        Shell "notepad " & GetProjectPath() & "Log\" & DateSelect.value & ".txt", vbNormalFocus
        
    End If
    
End Sub

Private Sub Form_Load()
    DateSelect.value = Now
End Sub
