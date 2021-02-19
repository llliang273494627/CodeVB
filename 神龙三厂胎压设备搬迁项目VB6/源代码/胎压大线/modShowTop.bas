Attribute VB_Name = "modShowTop"
'窗体置于顶层
'使用方法：在form load过程中SetTopWindow Me.hWnd

Option Explicit

Public Enum VideoWindowType
    OneWindow = 0
    FourWindow = 1
    NineWindow = 2
    SixteenWindow = 3
    OneplusFiveWindow = 4
End Enum

Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOREDRAW = &H8
Public Const SWP_SHOWWINDOW = &H0

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Declare Function SetWindowPos Lib "user32 " (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, ByVal Y As Long, _
                                                    ByVal cx As Long, ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
                                                
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_SYSMENU = &H80000
'
'Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
'Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Const GWL_STYLE = (-16)
'Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const GWL_WNDPROC = (-4)

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_SYSCOMMAND = &H112
Private Const SC_CLOSE = &HF060&
Private Const WM_CLOSE = &H10
Private Const WM_DESTROY = &H2
 
'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

 
Dim mlOldproc As Long
 
Private Function WndProc(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_SYSCOMMAND
            If wParam = SC_CLOSE Then
                SendMessage hWnd, WM_CLOSE, ByVal 0&, ByVal 0&
            End If
        Case WM_DESTROY
            SetWindowLong hWnd, GWL_WNDPROC, mlOldproc
    End Select
    WndProc = CallWindowProc(mlOldproc, hWnd, Msg, wParam, lParam)
End Function

Public Sub subclass(hWnd As Long)
    Dim lStyle As Long
    lStyle = GetWindowLong(hWnd, GWL_STYLE)
    lStyle = lStyle Or WS_MINIMIZEBOX Or WS_SYSMENU
    SetWindowLong hWnd, GWL_STYLE, lStyle
    mlOldproc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf WndProc)
End Sub

'设置窗口为最顶层
'函数:SetTopWindow
'参数:Winwnd   要设置为最顶层窗口的HWND
'返回值:
'例子：
Public Function SetTopWindow(WinWnd As Long)
'    SetWindowPos WinWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    SetWindowPos WinWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
End Function

