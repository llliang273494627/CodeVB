VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CVT520"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private m_CommPort As Integer

Private m_ComSettings As String

Private m_OpenPort As Boolean

Private m_Result As String

Private m_TireIDResult As String

Private m_TireMdlResult As String
Private m_TirePreResult As String
Private m_TireTempResult As String
Private m_TireBatteryResult As String
Private m_TireAcSpeedResult As String

Private m_Status As Integer

Private WithEvents m_Comm As MSComm
Attribute m_Comm.VB_VarHelpID = -1

Dim m_Form As Form


Private Sub Class_Initialize()
    Set m_Form = New Form1
    Set m_Comm = m_Form.MSComm1
    OpenPort = False
End Sub


Public Property Get CommPort() As Integer     '串口号
    CommPort = m_CommPort
End Property


Public Property Let CommPort(ByVal value As Integer)
    m_CommPort = value
    m_Comm.CommPort = value
End Property


Public Property Get ComSettings() As String    '串口参数
    ComSettings = m_ComSettings
End Property


Public Property Let ComSettings(ByVal value As String)
    m_ComSettings = value
    m_Comm.Settings = value
    m_Comm.RThreshold = 1
'    m_Comm.InputLen = 0
    m_Comm.InputMode = comInputModeBinary
'    m_Comm.InBufferSize = 1024
'    m_Comm.OutBufferSize = 512
End Property


Public Property Get OpenPort() As Boolean   '串口开关
    OpenPort = m_OpenPort
End Property


Public Property Let OpenPort(value As Boolean)
    On Error Resume Next
    If m_OpenPort = value Then
    Else
        If value = True Then
            m_Comm.PortOpen = True
            m_OpenPort = True
        Else
            m_Comm.PortOpen = False
            m_OpenPort = False
        End If
    End If
End Property



Public Property Get status() As Integer    '对VT520操作后返回状态
    status = m_Status
End Property




Public Property Get Result() As String     '返回测量结果字符串
    Result = m_Result
End Property

Public Property Get TireIDResult() As Variant '返回测量轮胎ID字符串
    TireIDResult = m_TireIDResult
End Property

Public Property Get TireTempResult() As Variant '温度
    TireTempResult = m_TireTempResult
End Property
Public Property Get TirePreResult() As Variant '压力
    TirePreResult = m_TirePreResult
End Property
Public Property Get TireAcSpeedResult() As Variant '加速度
    TireAcSpeedResult = m_TireAcSpeedResult
End Property
Public Property Get TireMdlResult() As Variant '模式
    TireMdlResult = m_TireMdlResult
End Property
Public Property Get TireBatteryResult() As Variant '电池
    TireBatteryResult = m_TireBatteryResult
End Property


'******************************************************************************
'** 函 数 名：Start
'** 输    入：
'** 输    出：
'** 功能描述：开始检测
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com.cn
'** 日    期：2009-03-03
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub Start(TestType As String, Optional portNum As Integer, Optional state As Boolean)
On Error Resume Next
    Select Case TestType
    Case "Comm"
       'FF 05 00 01 FF 00 C8 24
       Dim byt(7) As Byte
       byt(0) = &HFF
       byt(1) = &H5
       byt(2) = &H0
       byt(3) = &H1
       byt(4) = &HFF
       byt(5) = &H0
       byt(6) = &HC8
       byt(7) = &H24
       m_Comm.Output = byt
    Case "IO"
       Call OutputController(portNum, state)
    Case Else
        MsgBox "无效的TestType参数"
    End Select
    DelayTime 100
End Sub

'Add by ZCJ 2014-05-08
Public Sub SendProNum(proNum As Integer)
On Error Resume Next

    '设置程序号，第9个字节为程序号，最后两个字节为校验
    'FF 10 02 00 00 01 02 00 00 CD F4
    
    'FF 10 02 00 00 01 02 01 00 CC 64
    
    'FF 10 02 00 00 01 02 02 00 CC 94
    'FF 10 02 00 00 01 02 03 00 CD 04
    'FF 10 02 00 00 01 02 04 00 CF 34
    
    'FF 10 02 00 00 01 02 05 00 CE A4
    'FF 10 02 00 00 01 02 06 00 CE 54
    'FF 10 02 00 00 01 02 07 00 CF C4
    'FF 10 02 00 00 01 02 08 00 CA 34
    'FF 10 02 00 00 01 02 09 00 CB A4

    Dim chByte(10) As Byte
    
    chByte(0) = &HFF
    chByte(1) = &H10
    chByte(2) = &H2
    chByte(3) = &H0
    chByte(4) = &H0
    chByte(5) = &H1
    chByte(6) = &H2
    
    'Add by ZCJ 2013-06-25 控制器上程序号从0开始计数
    proNum = proNum - 1
    
    If proNum = 0 Then
        chByte(7) = &H0
        chByte(8) = &H0
        chByte(9) = &HCD
        chByte(10) = &HF4
    ElseIf proNum = 1 Then
        chByte(7) = &H1
        chByte(8) = &H0
        chByte(9) = &HCC
        chByte(10) = &H64
    ElseIf proNum = 2 Then
        chByte(7) = &H2
        chByte(8) = &H0
        chByte(9) = &HCC
        chByte(10) = &H94
    ElseIf proNum = 3 Then
        chByte(7) = &H3
        chByte(8) = &H0
        chByte(9) = &HCD
        chByte(10) = &H4
    ElseIf proNum = 4 Then
        chByte(7) = &H4
        chByte(8) = &H0
        chByte(9) = &HCF
        chByte(10) = &H34
    ElseIf proNum = 5 Then
        chByte(7) = &H5
        chByte(8) = &H0
        chByte(9) = &HCE
        chByte(10) = &HA4
    ElseIf proNum = 6 Then
        chByte(7) = &H6
        chByte(8) = &H0
        chByte(9) = &HCE
        chByte(10) = &H54
    ElseIf proNum = 7 Then
        chByte(7) = &H7
        chByte(8) = &H0
        chByte(9) = &HCF
        chByte(10) = &HC4
    ElseIf proNum = 8 Then
        chByte(7) = &H8
        chByte(8) = &H0
        chByte(9) = &HCA
        chByte(10) = &H34
    ElseIf proNum = 9 Then
        chByte(7) = &H9
        chByte(8) = &H0
        chByte(9) = &HCB
        chByte(10) = &HA4
    End If
    m_Comm.Output = chByte
    
    DelayTime 30
    
End Sub

'******************************************************************************
'** 函 数 名：ResetResult
'** 输    入：
'** 输    出：
'** 功能描述：清空上次测量结果
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com.cn
'** 日    期：2009-03-03
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub ResetResult()
On Error Resume Next
    'FF 05 00 02 FF 00 38 24
    m_TireIDResult = ""
    m_TireMdlResult = ""
    m_TirePreResult = ""
    m_TireTempResult = ""
    m_TireBatteryResult = ""
    m_TireAcSpeedResult = ""
    Dim byt(7) As Byte
    byt(0) = &HFF
    byt(1) = &H5
    byt(2) = &H0
    byt(3) = &H2
    byt(4) = &HFF
    byt(5) = &H0
    byt(6) = &H38
    byt(7) = &H24
    m_Comm.Output = byt
    DelayTime 100
End Sub


'******************************************************************************
'** 函 数 名：ReadResult
'** 输    入：
'** 输    出：
'** 功能描述：读取测量结果
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com.cn
'** 日    期：2009-03-03
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub ReadResult()
On Error Resume Next
'FF 03 00 10 00 0D 90 14
'FF 03 00 10 00 6E D0 3D
    Dim byt(7) As Byte
    byt(0) = &HFF
    byt(1) = &H3
    byt(2) = &H0
    byt(3) = &H10
    byt(4) = &H0
'    byt(5) = &HD
'    byt(6) = &H90
'    byt(7) = &H14
    byt(5) = &H6E
    byt(6) = &HD0
    byt(7) = &H3D
    m_Comm.Output = byt
    DelayTime 100
End Sub


Private Sub m_Comm_OnComm()
On Error Resume Next
    DelayTime 100 '原始是100ms
    Dim byt() As Byte
    byt = m_Comm.Input
    If UBound(byt()) = -1 Then Exit Sub
    Dim tmp As String
    Dim i As Integer
    For i = 0 To UBound(byt)
    
        tmp = tmp & Hex(byt(i)) & " "
    Next
    Select Case tmp
        Case "FF 5 0 1 FF 0 C8 24 "
            m_Status = 1                      '测量成功，状态属性设为1
        Case "FF 5 0 2 FF 0 38 24 "           '清空结果成功，状态属性设为3
            m_Status = 3
        Case Else
            'If Left(tmp, 7) = "FF 3 1A" Then
            If Left(tmp, 7) = "FF 3 DC" Then
                m_Status = 2                      '结果读取成功，状态属性设为2
                m_Result = tmp
                
'                m_TireIDResult =  IIf(Len(Hex(byt(26))) = 2, Hex(byt(26)), "0" & Hex(byt(26))) & IIf(Len(Hex(byt(25))) = 2, Hex(byt(25)), "0" & Hex(byt(25))) & IIf(Len(Hex(byt(24))) = 2, Hex(byt(24)), "0" & Hex(byt(24))) '取轮胎ID
                Dim gbit As String, zbit As String, cbit As String, dbit As String
                gbit = IIf(Len(Hex(byt(26))) = 2, Hex(byt(26)), "0" & Hex(byt(26)))
                zbit = IIf(Len(Hex(byt(25))) = 2, Hex(byt(25)), "0" & Hex(byt(25)))
                cbit = IIf(Len(Hex(byt(24))) = 2, Hex(byt(24)), "0" & Hex(byt(24)))
                dbit = IIf(Len(Hex(byt(23))) = 2, Hex(byt(23)), "0" & Hex(byt(23)))
                m_TireIDResult = gbit & zbit & cbit & dbit
                
                '模式，没有包含该信息
                gbit = Chr("&H" & Hex(byt(81)))
                zbit = Chr("&H" & Hex(byt(82)))
                cbit = gbit & zbit
                dbit = cBin("&H" & cbit)
                dbit = Left(dbit, (Len(dbit) - 4))
                m_TireMdlResult = BIN_to_HEX(dbit)
                
                '压力值
                gbit = IIf(Len(Hex(byt(14))) = 2, Hex(byt(14)), "0" & Hex(byt(14)))
                zbit = IIf(Len(Hex(byt(13))) = 2, Hex(byt(13)), "0" & Hex(byt(13)))
                cbit = IIf(Len(Hex(byt(12))) = 2, Hex(byt(12)), "0" & Hex(byt(12)))
                dbit = IIf(Len(Hex(byt(11))) = 2, Hex(byt(11)), "0" & Hex(byt(11)))
                m_TirePreResult = CLng("&H" & gbit & zbit & cbit & dbit) / 1000
                '温度值
                gbit = IIf(Len(Hex(byt(18))) = 2, Hex(byt(18)), "0" & Hex(byt(18)))
                zbit = IIf(Len(Hex(byt(17))) = 2, Hex(byt(17)), "0" & Hex(byt(17)))
                'm_TireTempResult = CLng("&H" & gbit & zbit) / 10
                'Modiy by ZCJ 2012-08-14 温度要除以100，精确到整数即可
                m_TireTempResult = CInt(CLng("&H" & gbit & zbit) / 100)
                '加速度值
                gbit = IIf(Len(Hex(byt(22))) = 2, Hex(byt(22)), "0" & Hex(byt(22)))
                zbit = IIf(Len(Hex(byt(21))) = 2, Hex(byt(21)), "0" & Hex(byt(21)))
                m_TireAcSpeedResult = CLng("&H" & gbit & zbit)
                '电池状态
                gbit = IIf(Len(Hex(byt(31))) = 2, Hex(byt(31)), "0" & Hex(byt(31)))
                zbit = IIf(Len(Hex(byt(32))) = 2, Hex(byt(32)), "0" & Hex(byt(32)))
                If gbit & zbit <> "4F4B" Then
                    m_TireBatteryResult = "Low"
                Else
                    m_TireBatteryResult = "OK"
                End If
            Else
                m_Status = 0                      '操作失败，状态属性设为0
            End If
    End Select
'    Debug.Print "a" & tmp & "b"
    
End Sub

Function cBin(ByVal N As Long) As String
On Error Resume Next
    Do
        cBin = N Mod 2 & cBin
        N = N \ 2
    Loop While N > 0
End Function

Public Function BIN_to_HEX(ByVal Bin As String) As String
On Error Resume Next
    Dim i As Long
    Dim H As String
    If Len(Bin) Mod 4 <> 0 Then
    Bin = String(4 - Len(Bin) Mod 4, "0") & Bin
    End If
    
    For i = 1 To Len(Bin) Step 4
    Select Case Mid(Bin, i, 4)
    Case "0000": H = H & "0"
    Case "0001": H = H & "1"
    Case "0010": H = H & "2"
    Case "0011": H = H & "3"
    Case "0100": H = H & "4"
    Case "0101": H = H & "5"
    Case "0110": H = H & "6"
    Case "0111": H = H & "7"
    Case "1000": H = H & "8"
    Case "1001": H = H & "9"
    Case "1010": H = H & "A"
    Case "1011": H = H & "B"
    Case "1100": H = H & "C"
    Case "1101": H = H & "D"
    Case "1110": H = H & "E"
    Case "1111": H = H & "F"
    End Select
    Next i
    While Left(H, 1) = "0"
    H = Right(H, Len(H) - 1)
    Wend
    BIN_to_HEX = H
End Function

Private Sub DelayTime(LngTime As Long)
  Dim LngTick As Long
  LngTick = GetTickCount()
  Do
   
     DoEvents: DoEvents
  Loop Until (GetTickCount() - LngTick) >= LngTime
End Sub


'******************************************************************************
'** 函 数 名：OutputController
'** 输    入：portNum——端口号（0-15）；state——开关状态（true=开，false=关）
'** 输    出：
'** 功能描述：IO卡输出控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub OutputController(portNum As Integer, state As Boolean)
    Debug.Print "端口号:" & portNum & "状态：" & state
End Sub
