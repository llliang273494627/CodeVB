VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'** 文件名：IOCard.cls
'** 版  权：CopyRight (c) 2009-2011 武汉华信数据系统有限公司
'** 创建人：hexiaoqin
'** 邮  箱：
'** 日  期：2009-03-05
'** 修改人：
'** 日  期：2009-03-05
'** 描  述：
'**
'** 版  本：1.0
'******************************************************************************
'Option Explicit
Private DIWordState(0 To 15) As Boolean    '用于接收输入的状态
Private DIState(0 To 15) As Boolean       '用于接收输入的状态
Private TestMing As String               '用于存放IO卡地址
Private PortDOState(0 To 15) As Boolean  '用于存放IO卡输出状态
Private DiValue As Integer               '取输入状态值
Private iPreVal As Integer               '输入的中间变量
Private iPreVal1 As Integer              '输入的中间变量
Event EventTest(testPort() As Boolean)
Private m_form As Form
'Private m_timer As Timer
Private WithEvents m_timer As Timer
Attribute m_timer.VB_VarHelpID = -1


Public Function getState() As Collection
    Dim col As Collection
    Dim i As Integer
    Set col = New Collection
    For i = 0 To UBound(DIWordState)
        col.Add DIWordState(i)
    Next
    Set getState = col
End Function







'******************************************************************************
'** 函 数 名：ShowPortChang
'** 输    入：
'** 输    出：IO输入的状态
'** 功能描述：用于提供外部查询IO输入的状态
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub ShowPortChang()
    Dim raiseE As Boolean
    raiseE = True
    Dim i As Integer
    For i = 0 To 15
        If DIWordState(i) <> DIState(i) Then
            raiseE = False
        '       MsgBox "portChang"
        End If
        DIWordState(i) = DIState(i)
    Next i
    If Not raiseE Then
        RaiseEvent EventTest(DIWordState)
    End If
End Sub

'******************************************************************************
'** 函 数 名：IniStallCard
'** 输    入：
'** 输    出：
'** 功能描述：初始化IO卡
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub IniStallCard()
  Dim gnNumOfDevices As Integer
  Dim nOutEntries As Integer
  Dim i, ii As Integer
  Dim tt As Long
  Dim tempStr As String

  ' Add type of PC Laboratory Card
  tt = DRV_GetAddress(devicelist(0))                                 '扫描设备
  ErrCde = DRV_DeviceGetList(tt, MaxEntries, nOutEntries)
  If (ErrCde <> 0) Then
    DRV_GetErrorMessage ErrCde, szErrMsg
    'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
    Exit Sub
  End If

  ErrCde = DRV_DeviceGetNumOfList(gnNumOfDevices)                    '扫描设备号
  If (ErrCde <> 0) Then
    DRV_GetErrorMessage ErrCde, szErrMsg
    'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
    Exit Sub
  End If

  For i = 0 To (gnNumOfDevices - 1)                                  '扫描设备地址
    tempStr = ""
    For ii = 0 To MaxDevNameLen
      tempStr = tempStr + Chr(devicelist(i).szDeviceName(ii))
    Next ii
      TestMing = Mid(tempStr, 1, 16)
      Debug.Print TestMing
  Next i
Call GetDevice
End Sub
'******************************************************************************
'** 函 数 名：ActivateCard
'** 输    入：IO卡通道号，730共有二个通道
'** 输    出：
'** 功能描述：激活当前卡的通道
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub ActivateCard(AddressNO As Integer)
Dim value As Long
lpDioGetCurrentDoByte.port = AddressNO
lpDioGetCurrentDoByte.value = DRV_GetAddress(value)

ErrCde = DRV_DioGetCurrentDOByte(DeviceHandle, lpDioGetCurrentDoByte)
If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szszErrMsg
        'Response = MsgBox(szszErrMsg, vbOKOnly, "Error!!")
        Exit Sub
End If
End Sub
'******************************************************************************
'** 函 数 名：GetDevice
'** 输    入：
'** 输    出：
'** 功能描述：取IO卡的一些参数和设置730
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Private Sub GetDevice()
  Dim i, ii As Integer
  Dim tempNum As Integer
  Dim TestRes As Boolean
  Dim gnNumOfSubdevices As Integer
  Dim nOutEntries As Integer
  Dim lpSubDeviceList As Long
  Dim dwDeviceNum As Long


  ' Avoid to open Advantech Demo Card
  TestRes = TestStr(TestMing, "DEMO")
  If (Not TestRes) Then
    ' Check if there is any device attatched on this COM port or CAN
    gnNumOfSubdevices = devicelist(0).nNumOfSubdevices
    If (gnNumOfSubdevices > MaxDev) Then
      gnNumOfSubdevices = MaxDev
    End If

    ' retrieve the information of all installed devices

    If (gnNumOfSubdevices = 0) Then
      dwDeviceNum = devicelist(0).dwDeviceNum
      ErrCde = DRV_DeviceOpen(dwDeviceNum, DeviceHandle)
      If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szErrMsg
        'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
        Exit Sub
      Else
        bRun = True
      End If

      ptDevGetFeatures.buffer = DRV_GetAddress(lpDevFeatures)
      ErrCde = DRV_DeviceGetFeatures(DeviceHandle, ptDevGetFeatures)
      If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szErrMsg
        'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
        Exit Sub
      End If
      tempNum = (lpDevFeatures.usMaxDOChl + 7) \ 8
      For i = 0 To (tempNum - 1)
        temp$ = "Port#" + str(i)
      Next i
    End If
  End If
End Sub
'******************************************************************************
'** 函 数 名：DOBit
'** 输    入：当前数据分解后的数据
'** 输    出：
'** 功能描述：用于分解IO卡的数据
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Function DOBit(bit As Integer) As Integer
  Dim i As Integer

  DOBit = 1
  If bit >= 1 Then
    For i = 1 To bit
      DOBit = DOBit * 2
    Next i
  End If

End Function
'******************************************************************************
'** 函 数 名：DOBitPort
'** 输    入：IO卡通道号，开关状态
'** 输    出：
'** 功能描述：执行IO输出操作
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Function DOBitPort(DOPort As Integer, OFFState As Boolean)
Dim i As Integer
PortDOState(DOPort) = OFFState
  DoValue = 0
For i = 0 To 7
  If PortDOState(i) = True Then
    DoValue = DoValue + DOBit(i)
'  Else
'    DoValue = DoValue + DOBit(DOPort)
  End If
Next i
  lpDioWritePort.port = lpDioPortMode.port
  lpDioWritePort.Mask = 255
  lpDioWritePort.state = DoValue
  ErrCde = DRV_DioWritePortByte(DeviceHandle, lpDioWritePort)
    If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szErrMsg
        'Response = MsgBox(szszErrMsg, vbOKOnly, "Error!!")
        Exit Function
    End If
End Function
'******************************************************************************
'** 函 数 名：OutputController
'** 输    入：DOportNo通道号，关开
'** 输    出：
'** 功能描述：供外部调用输出模块
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub OutputController(DOportNo As Integer, OFFState As Boolean)
If DOportNo < 8 Then
    
        lpDioPortMode.port = 0
        
        lpDioPortMode.dir = OUTPORT
    
        ' not every digital I/O card could use DRV_DioSetPortMode function
        If lpDevFeatures.usDIOPort > 0 Then
            ErrCde = DRV_DioSetPortMode(DeviceHandle, lpDioPortMode)
            If (ErrCde <> 0) Then
                DRV_GetErrorMessage ErrCde, szErrMsg
                'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
                Exit Sub
            End If
        End If
        Call ActivateCard(0)
        Call DOBitPort(DOportNo, OFFState)
End If
If DOportNo > 7 Then
    lpDioPortMode.port = 1
    
    lpDioPortMode.dir = OUTPORT

    ' not every digital I/O card could use DRV_DioSetPortMode function
    If lpDevFeatures.usDIOPort > 0 Then
        ErrCde = DRV_DioSetPortMode(DeviceHandle, lpDioPortMode)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            Exit Sub
        End If
    End If
    Call ActivateCard(1)
    Call DOBitPort(DOportNo - 8, OFFState)
End If
End Sub
'******************************************************************************
'** 函 数 名：TestStr
'** 输    入：地址码，状态码
'** 输    出：
'** 功能描述：分解出IO卡的当前地址号
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Function TestStr(DStr As String, TStr As String) As Boolean
  Dim lenD, lenT As Integer
  Dim i As Integer

  TestStr = False
  lenD = Len(DStr)
  lenT = Len(TStr)

  For i = 1 To (lenD - lenT + 1)
    If (Mid(DStr, i, lenT) = TStr) Then
      TestStr = True
    End If
  Next i

  If DStr = "" Then
    TestStr = True
  End If
End Function
'******************************************************************************
'** 函 数 名：GetPortValue
'** 输    入：IO卡通道号
'** 输    出：
'** 功能描述：用于获取IO卡的当前状态值
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub GetPortValue(PortAddress As Integer)
    lpDioPortMode.port = PortAddress
    lpDioPortMode.dir = INPORT
    If lpDevFeatures.usDIOPort > 0 Then
        ErrCde = DRV_DioSetPortMode(DeviceHandle, lpDioPortMode)
        If (ErrCde <> 0) Then
            DRV_GetErrorMessage ErrCde, szErrMsg
            'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
            Exit Sub
        End If
    End If
  
  lpDioReadPort.port = PortAddress
  lpDioReadPort.value = DRV_GetAddress(DiValue)
  ErrCde = DRV_DioReadPortByte(DeviceHandle, lpDioReadPort)
  If (ErrCde <> 0) Then
        DRV_GetErrorMessage ErrCde, szErrMsg
        'Response = MsgBox(szErrMsg, vbOKOnly, "Error!!")
        Exit Sub
  End If
  Call UpdateLed(PortAddress, DiValue)
End Sub
'******************************************************************************
'** 函 数 名：ActivateCard
'** 输    入：IO卡通道号，当前IO卡的状态值
'** 输    出：
'** 功能描述：主要用于更新输入状态
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub UpdateLed(AddressNO As Integer, iValue As Integer)
Dim i, iShift As Integer
    iShift = 1
    If AddressNO = 0 Then
        For i = 0 To 7
            If (iValue And iShift) <> (iPreVal And iShift) Then
                If (iValue And iShift) = iShift Then
                            DIState(i) = True
                Else
                            DIState(i) = False
                End If
            End If
            iShift = iShift * 2
        Next
        iPreVal = iValue
    ElseIf AddressNO = 1 Then
        For i = 0 To 7
            If (iValue And iShift) <> (iPreVal1 And iShift) Then
                If (iValue And iShift) = iShift Then
                            DIState(i + 8) = True
                Else
                            DIState(i + 8) = False
                End If
            End If
            iShift = iShift * 2
        Next
        iPreVal1 = iValue
    End If
End Sub
'******************************************************************************
'** 函 数 名：Class_Initialize
'** 输    入：
'** 输    出：
'** 功能描述：初始化类
'** 全局变量：
'** 调用模块：
'** 作    者：hexiaoqin
'** 邮    箱：
'** 日    期：2009-03-05
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub Class_Initialize()
    Set m_form = FrmIOTimer
    Set m_timer = FrmIOTimer.IOTimer
    Call IniStallCard
    m_timer.Enabled = True
    m_timer.Interval = 100
End Sub

Private Sub m_timer_Timer()
    Call GetPortValue(0)
    Call GetPortValue(1)
    ShowPortChang
End Sub
