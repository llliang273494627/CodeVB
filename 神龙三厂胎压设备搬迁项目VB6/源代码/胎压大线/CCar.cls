VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_VINCode As String

Private m_TireRFID As String

Private m_TireRFMdl As String
Private m_TireRFPre As String
Private m_TireRFTemp As String
Private m_TireRFBattery As String
Private m_TireRFAcSpeed As String

Private m_TireRRID As String

Private m_TireRRMdl As String
Private m_TireRRPre As String
Private m_TireRRTemp As String
Private m_TireRRBattery As String
Private m_TireRRAcSpeed As String

Private m_TireLFID As String

Private m_TireLFMdl As String
Private m_TireLFPre As String
Private m_TireLFTemp As String
Private m_TireLFBattery As String
Private m_TireLFAcSpeed As String

Private m_TireLRID As String

Private m_TireLRMdl As String
Private m_TireLRPre As String
Private m_TireLRTemp As String
Private m_TireLRBattery As String
Private m_TireLRAcSpeed As String

Private testState As Integer
Private overStandard As Boolean

'Public LastResulr As Boolean
Private LastVin As String
Public printFlag As Boolean
Public LastCar As CCar

Public Sub SetCarInfByVIN(vin As String)
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = New ADODB.Recordset
    rs.Open "select * from ""T_Result"" where ""VIN""='" & vin & "' ", DBCnnStr, 1, 3

    If rs.EOF Then
        Exit Sub
    End If
    m_VINCode = IIf(IsNull(rs("VIN").value), "", rs("VIN").value)
    m_TireRFID = IIf(IsNull(rs("ID020").value), "", rs("ID020").value)
    m_TireLFID = IIf(IsNull(rs("ID022").value), "", rs("ID022").value)
    m_TireRRID = IIf(IsNull(rs("ID021").value), "", rs("ID021").value)
    m_TireLRID = IIf(IsNull(rs("ID023").value), "", rs("ID023").value)
    
    m_TireRFMdl = IIf(IsNull(rs("Mdl020").value), "", rs("Mdl020").value)

    m_TireLFMdl = IIf(IsNull(rs("Mdl022").value), "", rs("Mdl022").value)

    m_TireRRMdl = IIf(IsNull(rs("Mdl021").value), "", rs("Mdl021").value)
    m_TireLRMdl = IIf(IsNull(rs("Mdl023").value), "", rs("Mdl023").value)
    
    m_TireRFPre = IIf(IsNull(rs("Pre020").value), "", rs("Pre020").value)
    m_TireLFPre = IIf(IsNull(rs("Pre022").value), "", rs("Pre022").value)
    m_TireRRPre = IIf(IsNull(rs("Pre021").value), "", rs("Pre021").value)
    m_TireLRPre = IIf(IsNull(rs("Pre023").value), "", rs("Pre023").value)
    
    m_TireRFTemp = IIf(IsNull(rs("Temp020").value), "", rs("Temp020").value)
    m_TireLFTemp = IIf(IsNull(rs("Temp022").value), "", rs("Temp022").value)
    m_TireRRTemp = IIf(IsNull(rs("Temp021").value), "", rs("Temp021").value)
    m_TireLRTemp = IIf(IsNull(rs("Temp023").value), "", rs("Temp023").value)
    
    m_TireRFBattery = IIf(IsNull(rs("Battery020").value), "", rs("Battery020").value)
    m_TireLFBattery = IIf(IsNull(rs("Battery022").value), "", rs("Battery022").value)
    m_TireRRBattery = IIf(IsNull(rs("Battery021").value), "", rs("Battery021").value)
    m_TireLRBattery = IIf(IsNull(rs("Battery023").value), "", rs("Battery023").value)
    
    m_TireRFAcSpeed = IIf(IsNull(rs("AcSpeed020").value), "", rs("AcSpeed020").value)
    m_TireLFAcSpeed = IIf(IsNull(rs("AcSpeed022").value), "", rs("AcSpeed022").value)
    m_TireRRAcSpeed = IIf(IsNull(rs("AcSpeed021").value), "", rs("AcSpeed021").value)
    m_TireLRAcSpeed = IIf(IsNull(rs("AcSpeed023").value), "", rs("AcSpeed023").value)
    
    testState = IIf(IsNull(rs("TestState").value), 0, rs("TestState").value)
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub CheckResultIsOverStandard()
On Error Resume Next
    Dim Result As Boolean
    overStandard = False
    Dim mdlArr() As String
    Dim i As Integer
    
    mdlArr = Split(mdlValue, ",")
    
    '判断右前轮
    '模式
    Result = judgeMdlIsOK(m_TireRFMdl, mdlArr)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    '压力
    Result = judgeResultIsOK(m_TireRFPre, preMinValue, preMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '温度
    Result = judgeResultIsOK(m_TireRFTemp, tempMinValue, tempMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '电池
    If m_TireRFBattery <> "OK" Then
        overStandard = True
        Exit Sub
    End If
    '加速度
    Result = judgeResultIsOK(m_TireRFAcSpeed, acSpeedMinValue, acSpeedMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    
    '判断左前轮
    '模式
    Result = judgeMdlIsOK(m_TireLFMdl, mdlArr)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    '压力
    Result = judgeResultIsOK(m_TireLFPre, preMinValue, preMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '温度
    Result = judgeResultIsOK(m_TireLFTemp, tempMinValue, tempMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '电池
    If m_TireLFBattery <> "OK" Then
        overStandard = True
        Exit Sub
    End If
    '加速度
    Result = judgeResultIsOK(m_TireLFAcSpeed, acSpeedMinValue, acSpeedMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    
    '判断右后轮
    '模式
    Result = judgeMdlIsOK(m_TireRRMdl, mdlArr)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    '压力
    Result = judgeResultIsOK(m_TireRRPre, preMinValue, preMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '温度
    Result = judgeResultIsOK(m_TireRRTemp, tempMinValue, tempMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '电池
    If m_TireRRBattery <> "OK" Then
        overStandard = True
        Exit Sub
    End If
    '加速度
    Result = judgeResultIsOK(m_TireRRAcSpeed, acSpeedMinValue, acSpeedMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    '判断左后轮
    '模式
    Result = judgeMdlIsOK(m_TireLRMdl, mdlArr)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    
    '压力
    Result = judgeResultIsOK(m_TireLRPre, preMinValue, preMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '温度
    Result = judgeResultIsOK(m_TireLRTemp, tempMinValue, tempMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
    '电池
    If m_TireLRBattery <> "OK" Then
        overStandard = True
        Exit Sub
    End If
    '加速度
    Result = judgeResultIsOK(m_TireLRAcSpeed, acSpeedMinValue, acSpeedMaxValue)
    If Not Result Then
        overStandard = True
        Exit Sub
    End If
End Sub

Public Property Get GetTestState() As String
    GetTestState = testState
End Property

Public Property Get IsOverStandard() As String
    IsOverStandard = overStandard
End Property

Public Property Get VINCode() As String
    VINCode = m_VINCode
End Property


Public Property Let VINCode(value As String)
    m_VINCode = value
End Property


Public Property Get TireRFID() As String
    TireRFID = m_TireRFID
End Property

Public Property Get TireRFMdl() As String
    TireRFMdl = m_TireRFMdl
End Property

Public Property Get TireRFPre() As String
    TireRFPre = m_TireRFPre
End Property

Public Property Get TireRFTemp() As String
    TireRFTemp = m_TireRFTemp
End Property

Public Property Get TireRFBattery() As String
    TireRFBattery = m_TireRFBattery
End Property

Public Property Get TireRFAcSpeed() As String
    TireRFAcSpeed = m_TireRFAcSpeed
End Property

Public Property Let TireRFID(value As String)
        m_TireRFID = value
End Property

Public Property Let TireRFMdl(value As String)
        m_TireRFMdl = value
End Property

Public Property Let TireRFPre(value As String)
        m_TireRFPre = value
End Property

Public Property Let TireRFTemp(value As String)
        m_TireRFTemp = value
End Property

Public Property Let TireRFBattery(value As String)
        m_TireRFBattery = value
End Property

Public Property Let TireRFAcSpeed(value As String)
        m_TireRFAcSpeed = value
End Property
'右后
Public Property Get TireRRID() As String
    TireRRID = m_TireRRID
End Property

Public Property Get TireRRMdl() As String
    TireRRMdl = m_TireRRMdl
End Property

Public Property Get TireRRPre() As String
    TireRRPre = m_TireRRPre
End Property

Public Property Get TireRRTemp() As String
    TireRRTemp = m_TireRRTemp
End Property

Public Property Get TireRRBattery() As String
    TireRRBattery = m_TireRRBattery
End Property

Public Property Get TireRRAcSpeed() As String
    TireRRAcSpeed = m_TireRRAcSpeed
End Property

Public Property Let TireRRID(ByVal value As String)
    m_TireRRID = value
End Property

Public Property Let TireRRMdl(value As String)
        m_TireRRMdl = value
End Property

Public Property Let TireRRPre(value As String)
        m_TireRRPre = value
End Property

Public Property Let TireRRTemp(value As String)
        m_TireRRTemp = value
End Property

Public Property Let TireRRBattery(value As String)
        m_TireRRBattery = value
End Property

Public Property Let TireRRAcSpeed(value As String)
        m_TireRRAcSpeed = value
End Property
'左前
Public Property Get TireLFID() As String
    TireLFID = m_TireLFID
End Property

Public Property Get TireLFMdl() As String
    TireLFMdl = m_TireLFMdl
End Property

Public Property Get TireLFPre() As String
    TireLFPre = m_TireLFPre
End Property

Public Property Get TireLFTemp() As String
    TireLFTemp = m_TireLFTemp
End Property

Public Property Get TireLFBattery() As String
    TireLFBattery = m_TireLFBattery
End Property

Public Property Get TireLFAcSpeed() As String
    TireLFAcSpeed = m_TireLFAcSpeed
End Property

Public Property Let TireLFID(value As String)
    m_TireLFID = value
End Property

Public Property Let TireLFMdl(value As String)
        m_TireLFMdl = value
End Property

Public Property Let TireLFPre(value As String)
        m_TireLFPre = value
End Property

Public Property Let TireLFTemp(value As String)
        m_TireLFTemp = value
End Property

Public Property Let TireLFBattery(value As String)
        m_TireLFBattery = value
End Property

Public Property Let TireLFAcSpeed(value As String)
        m_TireLFAcSpeed = value
End Property
'右后
Public Property Get TireLRID() As String
    TireLRID = m_TireLRID
End Property

Public Property Get TireLRMdl() As String
    TireLRMdl = m_TireLRMdl
End Property

Public Property Get TireLRPre() As String
    TireLRPre = m_TireLRPre
End Property

Public Property Get TireLRTemp() As String
    TireLRTemp = m_TireLRTemp
End Property

Public Property Get TireLRBattery() As String
    TireLRBattery = m_TireLRBattery
End Property

Public Property Get TireLRAcSpeed() As String
    TireLRAcSpeed = m_TireLRAcSpeed
End Property

Public Property Let TireLRID(value As String)
    m_TireLRID = value
End Property

Public Property Let TireLRMdl(value As String)
        m_TireLRMdl = value
End Property

Public Property Let TireLRPre(value As String)
        m_TireLRPre = value
End Property

Public Property Let TireLRTemp(value As String)
        m_TireLRTemp = value
End Property

Public Property Let TireLRBattery(value As String)
        m_TireLRBattery = value
End Property

Public Property Let TireLRAcSpeed(value As String)
        m_TireLRAcSpeed = value
End Property


Public Sub Save()
    On Error GoTo CCAR_SAVE_ERR
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim lastrs As ADODB.Recordset
    Dim sql As String
    Dim mtoc As String
    cnn.CommandTimeout = TimeOutNum
    cnn.Open DBCnnStr
    
    Set lastrs = cnn.Execute("select ""VIN"" from ""T_Result"" order by ""ID"" desc limit 1")
    printFlag = False
    If Not lastrs.EOF Then
        LastVin = lastrs(0).value
    End If
    Set LastCar = New CCar
    If LastVin <> m_VINCode And LastVin <> "" Then
                
        LastCar.SetCarInfByVIN LastVin
        If m_TireRFID <> "" And m_TireRFID <> "00000000" Then
            Select Case m_TireRFID
            Case LastCar.TireLFID
                printFlag = True
                LastCar.TireLFID = "00000000"
                m_TireRFID = "00000000"
            Case LastCar.TireLRID
                printFlag = True
                LastCar.TireLRID = "00000000"
                m_TireRFID = "00000000"
            Case LastCar.TireRFID
                printFlag = True
                LastCar.TireRFID = "00000000"
                m_TireRFID = "00000000"
            Case LastCar.TireRRID
                printFlag = True
                LastCar.TireRRID = "00000000"
                m_TireRFID = "00000000"
            End Select
        End If
        
        If m_TireLFID <> "" And m_TireLFID <> "00000000" Then
            Select Case m_TireLFID
            Case LastCar.TireLFID
                printFlag = True
                LastCar.TireLFID = "00000000"
                m_TireLFID = "00000000"
            Case LastCar.TireLRID
                printFlag = True
                LastCar.TireLRID = "00000000"
                m_TireLFID = "00000000"
            Case LastCar.TireRFID
                printFlag = True
                LastCar.TireRFID = "00000000"
                m_TireLFID = "00000000"
            Case LastCar.TireRRID
                printFlag = True
                LastCar.TireRRID = "00000000"
                m_TireLFID = "00000000"
            End Select
        End If
        
        If m_TireRRID <> "" And m_TireRRID <> "00000000" Then
            Select Case m_TireRRID
            Case LastCar.TireLFID
                printFlag = True
                LastCar.TireLFID = "00000000"
                m_TireRRID = "00000000"
            Case LastCar.TireLRID
                printFlag = True
                LastCar.TireLRID = "00000000"
                m_TireRRID = "00000000"
            Case LastCar.TireRFID
                printFlag = True
                LastCar.TireRFID = "00000000"
                m_TireRRID = "00000000"
            Case LastCar.TireRRID
                printFlag = True
                LastCar.TireRRID = "00000000"
                m_TireRRID = "00000000"
            End Select
        End If
        If m_TireLRID <> "" And m_TireLRID <> "00000000" Then
            Select Case m_TireLRID
            Case LastCar.TireLFID
                printFlag = True
                LastCar.TireLFID = "00000000"
                m_TireLRID = "00000000"
            Case LastCar.TireLRID
                printFlag = True
                LastCar.TireLRID = "00000000"
                m_TireLRID = "00000000"
            Case LastCar.TireRFID
                printFlag = True
                LastCar.TireRFID = "00000000"
                m_TireLRID = "00000000"
            Case LastCar.TireRRID
                printFlag = True
                LastCar.TireRRID = "00000000"
                m_TireLRID = "00000000"
            End Select
        End If
        'LastCar.Save
    End If
    

    testState = 0
    If Len(m_TireRFID) = 8 And m_TireRFID <> "00000000" And Trim(m_TireRFID) <> "" And m_TireRFID <> m_TireLFID And m_TireRFID <> m_TireLRID And m_TireRFID <> m_TireRRID Then
        testState = testState + 8
    End If
    If Len(m_TireLFID) = 8 And m_TireLFID <> "00000000" And Trim(m_TireLFID) <> "" And m_TireLFID <> m_TireRFID And m_TireLFID <> m_TireLRID And m_TireLFID <> m_TireRRID Then
        testState = testState + 4
    End If
    If Len(m_TireRRID) = 8 And m_TireRRID <> "00000000" And Trim(m_TireRRID) <> "" And m_TireRRID <> m_TireLFID And m_TireRRID <> m_TireLRID And m_TireRRID <> m_TireRFID Then
        testState = testState + 2
    End If
    If Len(m_TireLRID) = 8 And m_TireLRID <> "00000000" And Trim(m_TireLRID) <> "" And m_TireLRID <> m_TireLFID And m_TireLRID <> m_TireRFID And m_TireLRID <> m_TireRRID Then
        testState = testState + 1
    End If
                
    rs.Open "select * from ""T_Result"" where ""VIN""='" & m_VINCode & "' ", DBCnnStr, 1, 3
    If rs.EOF Then
        rs.AddNew
    End If
    
    
    rs("VIN").value = m_VINCode
    rs("VIS").value = Right(m_VINCode, 8)
    rs("ID020").value = m_TireRFID
    rs("ID022").value = m_TireLFID
    rs("ID021").value = m_TireRRID
    rs("ID023").value = m_TireLRID
    
    rs("Mdl020").value = m_TireRFMdl
    rs("Mdl022").value = m_TireLFMdl
    rs("Mdl021").value = m_TireRRMdl
    rs("Mdl023").value = m_TireLRMdl
    
    rs("Pre020").value = m_TireRFPre
    rs("Pre022").value = m_TireLFPre
    rs("Pre021").value = m_TireRRPre
    rs("Pre023").value = m_TireLRPre
    
    rs("Temp020").value = m_TireRFTemp
    rs("Temp022").value = m_TireLFTemp
    rs("Temp021").value = m_TireRRTemp
    rs("Temp023").value = m_TireLRTemp
    
    rs("Battery020").value = m_TireRFBattery
    rs("Battery022").value = m_TireLFBattery
    rs("Battery021").value = m_TireRRBattery
    rs("Battery023").value = m_TireLRBattery
    
    rs("AcSpeed020").value = m_TireRFAcSpeed
    rs("AcSpeed022").value = m_TireLFAcSpeed
    rs("AcSpeed021").value = m_TireRRAcSpeed
    rs("AcSpeed023").value = m_TireLRAcSpeed
    
    rs("TestTime").value = Now
    rs("TestState").value = testState
    rs("Dev").value = "101"
    rs("UploadSign").value = False
    rs("DownloadSign").value = False
    
    '获取MTOC码
    mtoc = GetMtocFromVinColl(m_VINCode)
    If mtoc <> "" Then
        rs("MTOC").value = mtoc
    End If
    
    rs.Update
    rs.Close
    Set rs = Nothing
    cnn.Close

    '存入远程数据库
    cnn.ConnectionTimeout = TimeOutNum 'Add by ZCJ 2012/02/21
    cnn.CommandTimeout = TimeOutNum
    cnn.Open RDBCnnStr

    rs.Open "select * from ""T_Result"" where ""VIN""='" & m_VINCode & "' ", RDBCnnStr, 1, 3

    If rs.EOF Then
        rs.AddNew
    End If


    rs("VIN").value = m_VINCode
    rs("VIS").value = Right(m_VINCode, 8)
    rs("ID020").value = m_TireRFID
    rs("ID022").value = m_TireLFID
    rs("ID021").value = m_TireRRID
    rs("ID023").value = m_TireLRID
    
    rs("Mdl020").value = m_TireRFMdl
    rs("Mdl022").value = m_TireLFMdl
    rs("Mdl021").value = m_TireRRMdl
    rs("Mdl023").value = m_TireLRMdl
    
    rs("Pre020").value = m_TireRFPre
    rs("Pre022").value = m_TireLFPre
    rs("Pre021").value = m_TireRRPre
    rs("Pre023").value = m_TireLRPre
    
    rs("Temp020").value = m_TireRFTemp
    rs("Temp022").value = m_TireLFTemp
    rs("Temp021").value = m_TireRRTemp
    rs("Temp023").value = m_TireLRTemp
    
    rs("Battery020").value = m_TireRFBattery
    rs("Battery022").value = m_TireLFBattery
    rs("Battery021").value = m_TireRRBattery
    rs("Battery023").value = m_TireLRBattery
    
    rs("AcSpeed020").value = m_TireRFAcSpeed
    rs("AcSpeed022").value = m_TireLFAcSpeed
    rs("AcSpeed021").value = m_TireRRAcSpeed
    rs("AcSpeed023").value = m_TireLRAcSpeed
    
    rs("Dev").value = "101"
    rs("TestTime").value = Now
    rs("TestState").value = testState
    rs("UploadSign").value = False
    rs("DownloadSign").value = False
    If mtoc <> "" Then
        rs("MTOC").value = mtoc
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    cnn.Close
    
    
    


    Set cnn = Nothing
    Exit Sub
CCAR_SAVE_ERR:
    LogWritter "CCAR_SAVE_ERR 错误信息：" & Err.Description
End Sub

Public Function GetMtocFromVinColl(vin As String) As String
On Error GoTo Err
    Dim cnn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    cnn.CommandTimeout = TimeOutNum
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select mtoc from vincoll where vin='" & vin & "'")
    If Not rs.EOF Then
        GetMtocFromVinColl = rs("mtoc").value & ""
    Else
        GetMtocFromVinColl = ""
    End If
    rs.Close
    Set rs = Nothing
    cnn.Close
    Exit Function
Err:
    GetMtocFromVinColl = ""
End Function

