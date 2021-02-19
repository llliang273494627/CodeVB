Attribute VB_Name = "modPublic"
'******************************************************************************
'** �ļ�����modPublic.bas
'** ��  Ȩ��CopyRight (c) 2008-2010 �人��������ϵͳ���޹�˾
'** �����ˣ�yangshuai
'** ��  �䣺shuaigoplay@live.cn
'** ��  �ڣ�2009-2-27
'** �޸��ˣ�
'** ��  �ڣ�
'** ��  ��������ģ��
'** ��  ����1.0
'******************************************************************************

Option Explicit

'�ر�ָ������
Private Type PROCESSENTRY32
dwSize As Long
cntUsage As Long
th32ProcessID As Long
th32DefaultHeapID As Long
th32ModuleID As Long
cntThreads As Long
th32ParentProcessID As Long
pcPriClassBase As Long
dwFlags As Long
szExeFile As String * 260
End Type
Private Declare Function CreateToolhelp32Snapshot Lib "KERNEL32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "KERNEL32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "KERNEL32" (ByVal hSnapShot As Long, lppe As PROCESSENTRY32) As Long
Private Declare Function OpenProcess Lib "KERNEL32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function TerminateProcess Lib "KERNEL32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Sub CloseHandle Lib "KERNEL32" (ByVal hPass As Long)
Private Const TH32CS_SNAPPROCESS = &H2&


Private Declare Function GetTickCount Lib "KERNEL32" () As Long
Public ProgramTitle As String       '����Title��������Ҫ��ʾ�ĵط�ȫ���øñ���������msgbox������Title����
Public DBCnnStr As String           '���ݿ������ַ���ȫ����Ҫ�������ݿ�ĵط�ȫ�����øñ���
Public RDBCnnStr As String

Public MESCnnStr As String      'MES���ݿ�������ַ���
Public MES_IP As String    'MES������IP��ַ

Public oIOCard As IOControl.IOCard  'IO���ƶ���

'VT520������ز���
Public oLVT520 As VT520DLL.CVT520    'VT520���ƶ���
Public LVT520_PortNum As Integer
Public LVT520_Settings As String
Public oRVT520 As VT520DLL.CVT520    'VT520���ƶ���
Public RVT520_PortNum As Integer
Public RVT520_Settings As String

'�źŵ���ؿ��Ʋ�����io�ź�����˿ڣ�
Public Lamp_GreenFlash_IOPort As Integer
Public Lamp_GreenLight_IOPort As Integer
Public Lamp_YellowLight_IOPort As Integer
Public Lamp_YellowFlash_IOPort As Integer
Public Lamp_RedLight_IOPort As Integer
Public Lamp_RedFlash_IOPort As Integer
Public Lamp_Buzzer_IOPort As Integer
Public Line_IOPort As Integer

'����ǹ����
Public WirledCodeGun_PortNum As String
Public WirledCodeGun_Settings As String
Public WirlessCodeGun_PortNum As String
Public WirlessCodeGun_Settings As String

'���ȿ��Ʋ�����io�ź�����˿ڣ�

'��ͬ���͵���̥����������Ӧ�Ŀ����������
Public ProNum_OldSensor As Integer '��ͨX7����(�ɴ�����)
Public ProNum_NewSensor As Integer 'X7 DSG&MRN ����(�´�����)


Public rdOutput As Integer
Public rdResetCommand As Integer

'��翪�ؿ������Լ����Ʋ���
Public sensor0 As CSensor
Public sensor1 As CSensor
Public sensor2 As CSensor
Public sensor3 As CSensor
Public sensor4 As CSensor
Public sensor5 As CSensor

Public sensorCommand As CSensor
Public sensorLine As CSensor
Public rdResetCommandS As CSensor

Public sensor0Port As Integer
Public sensor1Port As Integer
Public sensor2Port As Integer
Public sensor3Port As Integer
Public sensor4Port As Integer
Public sensor5Port As Integer

Public sensorCommandPort As Integer
Public sensorLinePort As Integer

'��������������
Public mdlValue As String
Public preMinValue As String
Public preMaxValue As String
Public tempMinValue As String
Public tempMaxValue As String
Public acSpeedMinValue As String
Public acSpeedMaxValue As String
Public mTOCStartIndex As String
Public tPMSCodeLen As String

'ϵͳɨ�������ģʽ
Public isCheckAllQueue As Boolean '�Ƿ�У���Ų�����
Public isOnlyScanVINCode As Boolean '�Ƿ�ֻɨ��VIN�룬MTOC�뽫���MESϵͳ�л��
Public isOnlyPrintNGWriteResult As Boolean '�Ƿ�ֻ��ӡ��Ͻ��ΪNG����ϵ���
Public isOnlyPrintNGFlow As Boolean '�Ƿ�ֻ��ӡNG��������̣��ϸ�����̲���ӡ

Public TimeOutNum As Integer
Public lineCommandFlag As Boolean

'******************************************************************************
'** �� �� ����main
'** ��    �룺
'** ��    ����
'** ����������������������ʼ��ȫ����������������������
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Sub Main()
On Error GoTo Main_Err

    DBCnnStr = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=DPCAWH1_DSG101" 'DSG101ODBC
    RDBCnnStr = getConfigValue("T_RunParam", "DB", "RDBCnnStr")
    TimeOutNum = getConfigValue("T_RunParam", "DB", "TimeOutNum")
    Dim X As Form
    For Each X In Forms
        Unload X
    Next

    '�õ���������getConfigValue
    '��̬��ȡ��������

    ProgramTitle = "DSG��ʼ��ϵͳ"

    MESCnnStr = getConfigValue("T_RunParam", "DB", "MESCnnStr")     'MESϵͳOracle���ݿ������ַ���
    MES_IP = getConfigValue("T_RunParam", "MES", "MESIP")   'MESϵͳ���ݿ����ڷ�����IP��ַ

    '��ʼ�����ƶ���

    '��ʼ��VT520����
    LVT520_PortNum = getConfigValue("T_CtrlParam", "LVT520", "LVT520_PortNum")
    LVT520_Settings = getConfigValue("T_CtrlParam", "LVT520", "LVT520_Settings")

    Set oLVT520 = New VT520DLL.CVT520
    oLVT520.CommPort = LVT520_PortNum
    oLVT520.ComSettings = LVT520_Settings
    oLVT520.OpenPort = True

    RVT520_PortNum = getConfigValue("T_CtrlParam", "RVT520", "RVT520_PortNum")
    RVT520_Settings = getConfigValue("T_CtrlParam", "RVT520", "RVT520_Settings")

    Set oRVT520 = New VT520DLL.CVT520
    oRVT520.CommPort = RVT520_PortNum
    oRVT520.ComSettings = RVT520_Settings
    oRVT520.OpenPort = True

    Set oIOCard = New IOControl.IOCard

    '��ȡ����ʼ�������źŵƿ��Ʋ���
    Lamp_GreenFlash_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_GreenFlash_IOPort")
    Lamp_GreenLight_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_GreenLight_IOPort")
    Lamp_YellowLight_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_YellowLight_IOPort")
    Lamp_RedLight_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_RedLight_IOPort")
    Lamp_RedFlash_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_RedFlash_IOPort")
    Lamp_Buzzer_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_Buzzer_IOPort")
    Lamp_YellowFlash_IOPort = getConfigValue("T_CtrlParam", "Lamp", "Lamp_YellowFlash_IOPort")

    Line_IOPort = getConfigValue("T_CtrlParam", "Line", "Line_IOPort")
    rdOutput = getConfigValue("T_CtrlParam", "Lamp", "rdOutput")
    rdResetCommand = getConfigValue("T_CtrlParam", "Lamp", "rdResetCommand")
    sensorCommandPort = getConfigValue("T_CtrlParam", "Line", "sensorCommandPort")
    sensorLinePort = getConfigValue("T_CtrlParam", "Line", "sensorLinePort")
    '��ʼ����翪��
    sensor0Port = getConfigValue("T_CtrlParam", "sensor", "sensor0Port")
    sensor1Port = getConfigValue("T_CtrlParam", "sensor", "sensor1Port")
    sensor2Port = getConfigValue("T_CtrlParam", "sensor", "sensor2Port")
    sensor3Port = getConfigValue("T_CtrlParam", "sensor", "sensor3Port")
    sensor4Port = getConfigValue("T_CtrlParam", "sensor", "sensor4Port")
    sensor5Port = getConfigValue("T_CtrlParam", "sensor", "sensor5Port")

    '�����������趨
    mdlValue = getConfigValue("T_RunParam", "StandardValue", "MdlValue")
    preMinValue = getConfigValue("T_RunParam", "StandardValue", "PreMinValue")
    preMaxValue = getConfigValue("T_RunParam", "StandardValue", "PreMaxValue")
    tempMinValue = getConfigValue("T_RunParam", "StandardValue", "TempMinValue")
    tempMaxValue = getConfigValue("T_RunParam", "StandardValue", "TempMaxValue")
    acSpeedMinValue = getConfigValue("T_RunParam", "StandardValue", "AcSpeedMinValue")
    acSpeedMaxValue = getConfigValue("T_RunParam", "StandardValue", "AcSpeedMaxValue")
    mTOCStartIndex = getConfigValue("T_RunParam", "TPMSCode", "MTOCStartIndex")
    tPMSCodeLen = getConfigValue("T_RunParam", "TPMSCode", "TPMSCodeLen")

    WirledCodeGun_PortNum = getConfigValue("T_CtrlParam", "BarCodeGun", "WirledCodeGun_PortNum")
    WirledCodeGun_Settings = getConfigValue("T_CtrlParam", "BarCodeGun", "WirledCodeGun_Settings")
    
    WirlessCodeGun_PortNum = getConfigValue("T_CtrlParam", "BarCodeGun", "WirlessCodeGun_PortNum")
    WirlessCodeGun_Settings = getConfigValue("T_CtrlParam", "BarCodeGun", "WirlessCodeGun_Settings")
    
    '��ͬ���͵���̥����������Ӧ�Ŀ����������
    ProNum_OldSensor = getConfigValue("T_CtrlParam", "ProgramNum", "ProNum_OldSensor")
    ProNum_NewSensor = getConfigValue("T_CtrlParam", "ProgramNum", "ProNum_NewSensor")

    lineCommandFlag = CBool(getConfigValue("T_CtrlParam", "sensor", "lineCommandFlag"))
    
    isCheckAllQueue = CBool(getConfigValue("T_RunParam", "Queue", "CheckAllQueue"))
    isOnlyScanVINCode = CBool(getConfigValue("T_RunParam", "Queue", "OnlyScanVINCode"))
    isOnlyPrintNGWriteResult = CBool(getConfigValue("T_RunParam", "Print", "OnlyPrintNGWriteResult"))
    isOnlyPrintNGFlow = CBool(getConfigValue("T_RunParam", "Print", "OnlyPrintNGFlow"))

    Set sensor0 = New CSensor
    Set sensor1 = New CSensor
    Set sensor2 = New CSensor
    Set sensor3 = New CSensor
    Set sensor4 = New CSensor
    Set sensor5 = New CSensor
    Set rdResetCommandS = New CSensor
    Set sensorCommand = New CSensor
    Set sensorLine = New CSensor

    sensor0.IOPort = sensor0Port
    sensor1.IOPort = sensor1Port
    sensor2.IOPort = sensor2Port
    sensor3.IOPort = sensor3Port
    sensor4.IOPort = sensor4Port
    sensor5.IOPort = sensor5Port

    rdResetCommandS.IOPort = rdResetCommand
    sensorCommand.IOPort = sensorCommandPort
    sensorLine.IOPort = sensorLinePort

    FrmMain.Show

    Exit Sub
Main_Err:
    
    MsgBox "��ʼ������ʧ�ܣ�������Ϣ��" & Err.Description & "������������Ϣ��"
    
End Sub



'******************************************************************************
'** �� �� ����exportExcel
'** ��    �룺sqlText����sql���
'** ��    ����
'** ��������������sql��ѯ���Ĳ�ѯ���
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-28
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************

Public Sub exportExcel(sqlText As String)
    Dim excelzfc As String
    Dim fileName As String
    Dim FSO As FileSystemObject
    Dim txtfile As TextStream
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim NowOutputDir As String
On Error GoTo exportExcel_ERR
    fileName = getExcelFileName '�õ�����ļ���
    cnn.Open DBCnnStr
    Set rs = cnn.Execute(sqlText)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    NowOutputDir = App.Path & "\Export"
    If Trim(Dir(NowOutputDir, vbDirectory)) = "" Then
        FSO.CreateFolder NowOutputDir
    End If
    
    Set txtfile = FSO.CreateTextFile(fileName, True)
    
'    For I = 0 To Me.MSFlexGrid1.Rows - 1
'        For J = 1 To Me.MSFlexGrid1.Cols - 1
'            excelzfc = excelzfc & MSFlexGrid1.TextMatrix(I, J) & Chr(9)
'        Next
'        txtfile.WriteLine
'    Next
    
    
    '�����ͷ
    Dim i As Integer
    Dim tmp As String
    For i = 0 To rs.Fields.Count - 1
         tmp = tmp & rs.Fields(i).Name & Chr(9)
    Next
    txtfile.WriteLine tmp
    
    '�������ڲ�
    Do While Not rs.EOF
        tmp = ""
        For i = 0 To rs.Fields.Count - 1
            tmp = tmp & rs(rs.Fields(i).Name).value & Chr(9)
        Next
        txtfile.WriteLine tmp
        rs.MoveNext
    Loop
    
    Set txtfile = Nothing
    Set FSO = Nothing
    
    '��excel
    Dim xlApp, xlbook, db1, xlsheet
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False '����ʾ����
    xlApp.Application.Visible = True '����ʾ����
    Set xlbook = xlApp.Workbooks.Open(fileName)
    Exit Sub
exportExcel_ERR:
    MsgBox "���ݵ���Excel����������Ϣ��" & Err.Description
End Sub

'******************************************************************************
'** �� �� ����getExcelFileName
'** ��    �룺
'** ��    �������ɵ��µ�excel�ļ���
'** �������������ɵ��µ�excel�ļ��� ��+��+��+ʱ+��+��+1000�������.xls
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-28
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Function getExcelFileName() As String
    Dim MyValue As String
    Randomize   ' �����������������ʼ���Ķ�����
    MyValue = Format(Int((1000 * Rnd) + 1), "0000")  ' ���� 1 �� 1000 ֮��������ֵ��
    getExcelFileName = GetProjectPath & "export\"
    getExcelFileName = getExcelFileName & Format(Year(Now), "0000")
    getExcelFileName = getExcelFileName & Format(Month(Now), "00")
    getExcelFileName = getExcelFileName & Format(Day(Now), "00")
    getExcelFileName = getExcelFileName & Format(Hour(Now), "00")
    getExcelFileName = getExcelFileName & Format(Minute(Now), "00")
    getExcelFileName = getExcelFileName & Format(Second(Now), "00")
    getExcelFileName = getExcelFileName & MyValue
    getExcelFileName = getExcelFileName & ".xls"
End Function

'******************************************************************************
'** �� �� ����GetProjectPath
'** ��    �룺
'** ��    ����
'** ����������
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************

Public Function GetProjectPath() As String
   If Right(App.Path, 1) <> "\" Then
      GetProjectPath = App.Path + "\"
   Else
      GetProjectPath = App.Path
   End If
End Function

'******************************************************************************
'** �� �� ����hasDSG
'** ��    �룺
'** ��    ����
'** �����������Ƿ�װ��DSG
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Function hasDSG(CarCode As String) As Boolean
On Error GoTo hasDSG_Err
    Dim tmpV As String
    tmpV = Mid(CarCode, 24, 1) 'ȡ��24λֵ
    'Modiy by ZCJ 20130625 ������һ�ִ�����
    If tmpV = "D" Or tmpV = "A" Then
        hasDSG = True
        
        'Add by ZCJ 20130625
        'FrmMain.CarTypeCode = tmpV
        '���ó����
        If tmpV = "D" Then
            SetProNum (ProNum_OldSensor) '�ɴ�����
        ElseIf tmpV = "A" Then
            SetProNum (ProNum_NewSensor) '�´�����
        End If
    Else
        hasDSG = False
    End If
    Exit Function
    
hasDSG_Err:
    LogWritter "hasDSG�����ڷ��ִ��󣬴�����Ϣ��" & Err.Description
    hasDSG = False
End Function
'Add by ZCJ 2012-10-20 �����������߿������ĳ����
Public Function SetProNum(ProNum As String)
On Error GoTo SetProNum_Err
    oRVT520.SendProNum CInt(ProNum)
    oLVT520.SendProNum CInt(ProNum)
    LogWritter "���������ĳ��������Ϊ" & ProNum
    
    Exit Function
SetProNum_Err:
    LogWritter "�����ÿ����������Ϊ" & ProNum & "ʱ����������Ϣ��" & Err.Description
End Function

'******************************************************************************
'** �� �� ����getConfigValue
'** ��    �룺
'** ��    ����
'** �����������õ�����ֵ
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Function getConfigValue(tableName As String, group As String, key As String) As String
    On Error GoTo getConfigValue_err
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""Value"" from """ & tableName & """ where ""Group""='" & group & "' and ""Key""='" & key & "' ")
    If Not rs.EOF Then
        getConfigValue = rs(0).value
    Else
        getConfigValue = ""
    End If
    cnn.Close
    Set cnn = Nothing
    Exit Function
getConfigValue_err:
    LogWritter "���ݿ�������󣡴�����Ϣ��" & Err.Description
    If cnn.state = 1 Then
        cnn.Close
    End If
    Set cnn = Nothing
End Function
''******************************************************************************
''** �� �� ����setConfigValue
''** ��    �룺
''** ��    ����
''** ������������������ֵ
''** ȫ�ֱ�����
''** ��    �ߣ�yangshuai
''** ��    �䣺shuaigoplay@live.cn
''** ��    �ڣ�2009-2-27
''** �� �� �ߣ�
''** ��    �ڣ�
''** ��    ����1.0
''******************************************************************************
'Public Function setConfigValue(tableName As String, Group As String, key As String, value As String)
'    Dim cnn As New ADODB.Connection
'    Dim rs As ADODB.Recordset
'    cnn.Open DBCnnStr
'    Set rs = cnn.Execute("select ")
'    cnn.Close
'End Function

Public Sub printErrResult(car As CCar)

    Dim tmpStr As String
    Dim rs As New ADODB.Recordset
    Dim mdlArr() As String
    
    rs.Fields.Append "name", adBSTR
    rs.Open
    rs.AddNew
    rs("name").value = "name"

    Set DataReport1.DataSource = rs

    DataReport1.Sections("Section1").Controls("lblVIN").Caption = DataReport1.Sections("Section1").Controls("lblVIN").Caption & car.VINCode


    DataReport1.Sections("Section1").Controls("lbldate").Caption = DataReport1.Sections("Section1").Controls("lbldate").Caption & Date
    DataReport1.Sections("Section1").Controls("lbltime").Caption = DataReport1.Sections("Section1").Controls("lbltime").Caption & Time
    If car.GetTestState = 15 Then
'        If car.IsOverStandard Then 'Modiy by ZCJ 2012-07-09
'            DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
'            DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
'        Else
            DataReport1.Sections("Section1").Controls("labResult").Caption = "OK"
'        End If
    Else
        DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
        DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
    End If
    Dim resultState As String
    resultState = DToB(car.GetTestState)

    mdlArr = Split(mdlValue, ",")

    If Mid(resultState, 1, 1) = "1" Then
        DataReport1.Sections("Section1").Controls("lbl1").Caption = DataReport1.Sections("Section1").Controls("lbl1").Caption & car.TireRFID
        '�ж�ģʽ
        If judgeMdlIsOK(car.TireRFMdl, mdlArr) = False Then
            tmpStr = ";ģʽ" & car.TireRFMdl & "(���ϸ�)"
        End If
        
        '�ж�ѹ��ֵ�Ƿ�ϸ�
        If CCur(car.TireRFPre) < CCur(preMinValue) Then
            tmpStr = ";ѹ��" & car.TireRFPre & "kPa(ƫ��)"
        ElseIf CCur(car.TireRFPre) > CCur(preMaxValue) Then
            tmpStr = ";ѹ��" & car.TireRFPre & "kPa(ƫ��)"
        End If
        '�ж��¶�ֵ�Ƿ�ϸ�
        If CCur(car.TireRFTemp) < CCur(tempMinValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireRFTemp & "��(ƫ��)"
        ElseIf CCur(car.TireRFTemp) > CCur(tempMaxValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireRFTemp & "��(ƫ��)"
        End If
        '�жϼ��ٶ��Ƿ�ϸ�
        If CCur(car.TireRFAcSpeed) < CCur(acSpeedMinValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireRFAcSpeed & "g(ƫ��)"
        ElseIf CCur(car.TireRFAcSpeed) > CCur(acSpeedMaxValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireRFAcSpeed & "g(ƫ��)"
        End If
        '�жϵ�ص���
        If car.TireRFBattery <> "OK" Then
            tmpStr = tmpStr & ";��ص�����"
        End If
    Else
        DataReport1.Sections("Section1").Controls("lbl1").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl1").Caption = DataReport1.Sections("Section1").Controls("lbl1").Caption & "���ʧ��"
    End If
    If tmpStr <> "" Then
        DataReport1.Sections("Section1").Controls("lbl1").Caption = DataReport1.Sections("Section1").Controls("lbl1").Caption & tmpStr
        tmpStr = ""
        DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
        DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl1").ForeColor = &HFF&
    End If
    
        
        

    If Mid(resultState, 2, 1) = "1" Then
        DataReport1.Sections("Section1").Controls("lbl2").Caption = DataReport1.Sections("Section1").Controls("lbl2").Caption & car.TireLFID
        '�ж�ģʽ
        If judgeMdlIsOK(car.TireLFMdl, mdlArr) = False Then
            tmpStr = ";ģʽ" & car.TireLFMdl & "(���ϸ�)"
        End If
        
        '�ж�ѹ��ֵ�Ƿ�ϸ�
        If CCur(car.TireLFPre) < CCur(preMinValue) Then
            tmpStr = ";ѹ��" & car.TireLFPre & "kPa(ƫ��)"
        ElseIf CCur(car.TireLFPre) > CCur(preMaxValue) Then
            tmpStr = ";ѹ��" & car.TireLFPre & "kPa(ƫ��)"
        End If
        '�ж��¶�ֵ�Ƿ�ϸ�
        If CCur(car.TireLFTemp) < CCur(tempMinValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireLFTemp & "��(ƫ��)"
        ElseIf CCur(car.TireLFTemp) > CCur(tempMaxValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireLFTemp & "��(ƫ��)"
        End If
        '�жϼ��ٶ��Ƿ�ϸ�
        If CCur(car.TireLFAcSpeed) < CCur(acSpeedMinValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireLFAcSpeed & "g(ƫ��)"
        ElseIf CCur(car.TireLFAcSpeed) > CCur(acSpeedMaxValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireLFAcSpeed & "g(ƫ��)"
        End If
        '�жϵ�ص���
        If car.TireLFBattery <> "OK" Then
            tmpStr = tmpStr & ";��ص�����"
        End If
    Else
        DataReport1.Sections("Section1").Controls("lbl2").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl2").Caption = DataReport1.Sections("Section1").Controls("lbl2").Caption & "���ʧ��"
    End If
    If tmpStr <> "" Then
        DataReport1.Sections("Section1").Controls("lbl2").Caption = DataReport1.Sections("Section1").Controls("lbl2").Caption & tmpStr
        tmpStr = ""
        DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
        DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl2").ForeColor = &HFF&
    End If
    
    
    If Mid(resultState, 3, 1) = "1" Then
        DataReport1.Sections("Section1").Controls("lbl4").Caption = DataReport1.Sections("Section1").Controls("lbl4").Caption & car.TireRRID
        '�ж�ģʽ
        If judgeMdlIsOK(car.TireRRMdl, mdlArr) = False Then
            tmpStr = ";ģʽ" & car.TireRRMdl & "(���ϸ�)"
        End If
        
        '�ж�ѹ��ֵ�Ƿ�ϸ�
        If CCur(car.TireRRPre) < CCur(preMinValue) Then
            tmpStr = ";ѹ��" & car.TireRRPre & "kPa(ƫ��)"
        ElseIf CCur(car.TireRRPre) > CCur(preMaxValue) Then
            tmpStr = ";ѹ��" & car.TireRRPre & "kPa(ƫ��)"
        End If
        '�ж��¶�ֵ�Ƿ�ϸ�
        If CCur(car.TireRRTemp) < CCur(tempMinValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireRRTemp & "��(ƫ��)"
        ElseIf CCur(car.TireRRTemp) > CCur(tempMaxValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireRRTemp & "��(ƫ��)"
        End If
        '�жϼ��ٶ��Ƿ�ϸ�
        If CCur(car.TireRRAcSpeed) < CCur(acSpeedMinValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireRRAcSpeed & "g(ƫ��)"
        ElseIf CCur(car.TireRRAcSpeed) > CCur(acSpeedMaxValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireRRAcSpeed & "g(ƫ��)"
        End If
        '�жϵ�ص���
        If car.TireRRBattery <> "OK" Then
            tmpStr = tmpStr & ";��ص�����"
        End If
    Else
        DataReport1.Sections("Section1").Controls("lbl4").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl4").Caption = DataReport1.Sections("Section1").Controls("lbl4").Caption & "���ʧ��"
    End If
    If tmpStr <> "" Then
        DataReport1.Sections("Section1").Controls("lbl4").Caption = DataReport1.Sections("Section1").Controls("lbl4").Caption & tmpStr
        tmpStr = ""
        DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
        DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl4").ForeColor = &HFF&
    End If
    
    

    If Mid(resultState, 4, 1) = "1" Then
        DataReport1.Sections("Section1").Controls("lbl3").Caption = DataReport1.Sections("Section1").Controls("lbl3").Caption & car.TireLRID
        '�ж�ģʽ
        If judgeMdlIsOK(car.TireLRMdl, mdlArr) = False Then
            tmpStr = ";ģʽ" & car.TireLRMdl & "(���ϸ�)"
        End If
        
        '�ж�ѹ��ֵ�Ƿ�ϸ�
        If CCur(car.TireLRPre) < CCur(preMinValue) Then
            tmpStr = ";ѹ��" & car.TireLRPre & "kPa(ƫ��)"
        ElseIf CCur(car.TireLRPre) > CCur(preMaxValue) Then
            tmpStr = ";ѹ��" & car.TireLRPre & "kPa(ƫ��)"
        End If
        '�ж��¶�ֵ�Ƿ�ϸ�
        If CCur(car.TireLRTemp) < CCur(tempMinValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireLRTemp & "��(ƫ��)"
        ElseIf CCur(car.TireLRTemp) > CCur(tempMaxValue) Then
            tmpStr = tmpStr & ";�¶�" & car.TireLRTemp & "��(ƫ��)"
        End If
        '�жϼ��ٶ��Ƿ�ϸ�
        If CCur(car.TireLRAcSpeed) < CCur(acSpeedMinValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireLRAcSpeed & "g(ƫ��)"
        ElseIf CCur(car.TireLRAcSpeed) > CCur(acSpeedMaxValue) Then
            tmpStr = tmpStr & ";���ٶ�" & car.TireLRAcSpeed & "g(ƫ��)"
        End If
        '�жϵ�ص���
        If car.TireLRBattery <> "OK" Then
            tmpStr = tmpStr & ";��ص�����"
        End If
    Else
        DataReport1.Sections("Section1").Controls("lbl3").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl3").Caption = DataReport1.Sections("Section1").Controls("lbl3").Caption & "���ʧ��"
    End If
    If tmpStr <> "" Then
        DataReport1.Sections("Section1").Controls("lbl3").Caption = DataReport1.Sections("Section1").Controls("lbl3").Caption & tmpStr
        tmpStr = ""
        DataReport1.Sections("Section1").Controls("labResult").Caption = "NG"
        DataReport1.Sections("Section1").Controls("labResult").ForeColor = &HFF&
        DataReport1.Sections("Section1").Controls("lbl3").ForeColor = &HFF&
    End If


    DataReport1.PrintReport
    Unload DataReport1
End Sub

Public Sub printErrCode()
    On Error Resume Next
    
    'DoEvents
    
    Dim tmpStr As String
    Dim rsDB As New ADODB.Recordset
    rsDB.Fields.Append "name", adBSTR
    rsDB.Open
    rsDB.AddNew
    rsDB("name").value = "name"
    Set WriteInErrorCode.DataSource = rsDB
    
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim isWriteIn As Boolean
    Dim writeInResult As Boolean
    Dim isPrint As Boolean
    Dim errorCodeList() As String
    Dim i As Integer
    
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""VIN"",""ID020"",""ID022"",""ID021"",""ID023"",""WriteInTime"",""IsWriteIn"",""WriteInResult"",""ErrorCode"",""IsPrint"" from ""T_Result"" where ""IsWriteIn""=true and ""IsPrint""=false order by ""ID"" asc limit 1")
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        cnn.Close
        Set cnn = Nothing
        Exit Sub
    End If
    
    isWriteIn = IIf(IsNull(rs("IsWriteIn")), False, CBool(rs("IsWriteIn")))
    writeInResult = IIf(IsNull(rs("WriteInResult")), False, CBool(rs("WriteInResult")))
    isPrint = IIf(IsNull(rs("IsPrint")), True, CBool(rs("IsPrint")))
    
    If isWriteIn And Not isPrint Then
        
        WriteInErrorCode.Sections("Section1").Controls("lbVIN").Caption = "VIN�룺" & rs("VIN")
        WriteInErrorCode.Sections("Section1").Controls("lbDateTime").Caption = "���ڣ�" & Format(rs("WriteInTime"), "yyyy-MM-dd HH:mm:ss")
        WriteInErrorCode.Sections("Section1").Controls("lbResult").Caption = "���                            " & IIf(writeInResult, "�ϸ�", "���ϸ�")
        
        WriteInErrorCode.Sections("Section1").Controls("lbLF").Caption = "��ǰ�֣�" & rs("ID022")
        If CStr(rs("ID022")) = "00000000" Or CStr(rs("ID022")) = "" Then
            WriteInErrorCode.Sections("Section1").Controls("lbLF").ForeColor = &HFF&
        End If
        
        WriteInErrorCode.Sections("Section1").Controls("lbRF").Caption = "��ǰ�֣�" & rs("ID020")
        If CStr(rs("ID020")) = "00000000" Or CStr(rs("ID020")) = "" Then
            WriteInErrorCode.Sections("Section1").Controls("lbRF").ForeColor = &HFF&
        End If
        
        WriteInErrorCode.Sections("Section1").Controls("lbLR").Caption = "����֣�" & rs("ID023")
        If CStr(rs("ID023")) = "00000000" Or CStr(rs("ID023")) = "" Then
            WriteInErrorCode.Sections("Section1").Controls("lbLR").ForeColor = &HFF&
        End If
        
        WriteInErrorCode.Sections("Section1").Controls("lbRR").Caption = "�Һ��֣�" & rs("ID021")
        If CStr(rs("ID021")) = "00000000" Or CStr(rs("ID021")) = "" Then
            WriteInErrorCode.Sections("Section1").Controls("lbRR").ForeColor = &HFF&
        End If
        
        If Not writeInResult Then
            WriteInErrorCode.Sections("Section1").Controls("lbResult").ForeColor = &HFF&
        End If
        
        errorCodeList = Split(CStr(rs("ErrorCode")), ";")
        For i = 0 To UBound(errorCodeList)
            
            If i <> UBound(errorCodeList) Then
                WriteInErrorCode.Sections("Section1").Controls("lbError" & (i + 1)).Caption = errorCodeList(i)
                If Right(errorCodeList(i), 2) = "ʧ��" Or Right(errorCodeList(i), 3) = "���ϸ�" Then
                    WriteInErrorCode.Sections("Section1").Controls("lbError" & (i + 1)).ForeColor = &HFF&
                End If
            End If
        Next
        
        cnn.Execute "update ""T_Result"" set ""IsPrint""=true where ""VIN""='" & rs("VIN") & "'"
        
        WriteInErrorCode.PrintReport
        Unload WriteInErrorCode
    Else
        
    End If
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub printErrCodeAuto()
    On Error Resume Next
    
    'DoEvents
    
    Dim tmpStr As String
    Dim rsDB As New ADODB.Recordset
    rsDB.Fields.Append "name", adBSTR
    rsDB.Open
    rsDB.AddNew
    rsDB("name").value = "name"
    Set WriteInErrorCodeAuto.DataSource = rsDB
    
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim isWriteIn As Boolean
    Dim writeInResult As Boolean
    Dim isPrint As Boolean
    Dim errorCodeList() As String
    Dim rowArr() As String
    Dim i As Integer
    Dim tmpIndex As Integer
    Dim maxID As Integer
    Dim tmp As Integer
    
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select max(""ID"") as ""ID"" from ""T_Result"" where ""IsPrint""=true")
    If Not rs.EOF Then
        maxID = CInt(rs("ID"))
    Else
        maxID = 0
    End If
    
    If isOnlyPrintNGWriteResult Then
        Set rs = cnn.Execute("select ""VIN"",""ID020"",""ID022"",""ID021"",""ID023"",""WriteInTime"",""IsWriteIn"",""WriteInResult"",""ErrorCode"",""IsPrint"",""MTOC"" from ""T_Result"" where ""IsWriteIn""=true and ""WriteInResult""=false and ""IsPrint""=false and ""ID"">" & maxID & " order by ""ID"" asc limit 1")
    Else
        Set rs = cnn.Execute("select ""VIN"",""ID020"",""ID022"",""ID021"",""ID023"",""WriteInTime"",""IsWriteIn"",""WriteInResult"",""ErrorCode"",""IsPrint"",""MTOC"" from ""T_Result"" where ""IsWriteIn""=true and ""IsPrint""=false and ""ID"">" & maxID & " order by ""ID"" asc limit 1")
    End If
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        cnn.Close
        Set cnn = Nothing
        Exit Sub
    End If
    
    isWriteIn = IIf(IsNull(rs("IsWriteIn")), False, CBool(rs("IsWriteIn")))
    writeInResult = IIf(IsNull(rs("WriteInResult")), False, CBool(rs("WriteInResult")))
    isPrint = IIf(IsNull(rs("IsPrint")), True, CBool(rs("IsPrint")))
    
    If isWriteIn And Not isPrint Then
        
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbVIN").Caption = "�������룺" & rs("VIN")
        WriteInErrorCodeAuto.Sections("Section1").Controls("lblMTOC").Caption = "MTOC�룺" & rs("MTOC")
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbDateTime").Caption = "���ڣ�" & Format(rs("WriteInTime"), "yyyy-MM-dd HH:mm:ss")
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult").Caption = "���                            " & IIf(writeInResult, "�ϸ�", "���ϸ�")
        
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbLF").Caption = "��ǰ�֣�" & rs("ID022")
        If CStr(rs("ID022")) = "00000000" Or CStr(rs("ID022")) = "" Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbLF").ForeColor = &HFF&
        End If
        
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbRF").Caption = "��ǰ�֣�" & rs("ID020")
        If CStr(rs("ID020")) = "00000000" Or CStr(rs("ID020")) = "" Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbRF").ForeColor = &HFF&
        End If
        
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbLR").Caption = "����֣�" & rs("ID023")
        If CStr(rs("ID023")) = "00000000" Or CStr(rs("ID023")) = "" Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbLR").ForeColor = &HFF&
        End If
        
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbRR").Caption = "�Һ��֣�" & rs("ID021")
        If CStr(rs("ID021")) = "00000000" Or CStr(rs("ID021")) = "" Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbRR").ForeColor = &HFF&
        End If
        
        If Not writeInResult Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult").ForeColor = &HFF&
        End If
        
        If CStr(rs("ErrorCode") & "") = "" Then
            errorCodeList = Split(CStr(rs("ErrorCode") & "&S"), "&S")
        Else
            errorCodeList = Split(CStr(rs("ErrorCode")), "&S")
        End If
        
        'WriteInErrorCodeAuto.Sections("Section1").Visible = False
        'WriteInErrorCodeAuto.Sections("Section1").Height = 3000
        'DataReport1.Sections("Section1").Controls("Text1").CanGrow = True '�Զ�����

        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (1)).Caption = ""
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (1)).Caption = ""
        
        i = 0
        If UBound(errorCodeList) > -1 Then
            For i = 0 To UBound(errorCodeList)
                
                If i <> UBound(errorCodeList) Then
                    If Left(errorCodeList(i), 2) = "&P" Then
                        rowArr = Split(CStr(errorCodeList(i)), "&C")
                        rowArr(0) = Replace(rowArr(0), "&P", (i + 1) & " ") '���
                        If rowArr(1) = "ʧ��" Or rowArr(1) = "���ϸ�" Then
                            tmpIndex = tmpIndex + 1
                            If isOnlyPrintNGFlow Then
                                tmp = tmpIndex
                            Else
                                tmp = i + 1
                            End If
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).ForeColor = &HFF&
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).ForeColor = &HFF&
                        End If
                        If Len(rowArr(0)) > 32 Then
                            rowArr(0) = Mid(rowArr(0), 1, 32)
                        End If
                        If isOnlyPrintNGFlow Then
                            If rowArr(1) = "�ɹ�" Then
                                
                            Else
                                WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmpIndex)).Caption = rowArr(0)
                                WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmpIndex)).Caption = rowArr(1)
                            End If
                        Else
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Caption = rowArr(0)
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Caption = rowArr(1)
                        End If
                    Else
                        tmpIndex = tmpIndex + 1
                        errorCodeList(i) = "  " & errorCodeList(i)
                        If isOnlyPrintNGFlow Then
                            tmp = tmpIndex
                        Else
                            tmp = i + 1
                        End If
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).ForeColor = &HFF&
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).Top = 15
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).Visible = False
                        If Len(errorCodeList(i)) > 32 Then
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).Width = 4050
                        End If
                        If Len(errorCodeList(i)) > 36 Then
                            errorCodeList(i) = Mid(errorCodeList(i), 1, 36)
                        End If
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).Caption = errorCodeList(i)
                    End If
                End If
            Next
        
            If isOnlyPrintNGFlow Then
                i = tmpIndex
            Else
                i = UBound(errorCodeList)
            End If
        End If
        
        For i = i To 31
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Top = 15
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Visible = False
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Top = 15
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Visible = False
        Next i
        
        If isOnlyPrintNGFlow Then
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbErrorEnd").Top = 3000 + tmpIndex * 330
            WriteInErrorCodeAuto.Sections("Section1").Height = 3300 + tmpIndex * 330
        Else
            WriteInErrorCodeAuto.Sections("Section1").Controls("lbErrorEnd").Top = 3000 + UBound(errorCodeList) * 330
            WriteInErrorCodeAuto.Sections("Section1").Height = 3300 + UBound(errorCodeList) * 330
        End If
        
        cnn.Execute "update ""T_Result"" set ""IsPrint""=true where ""VIN""='" & rs("VIN") & "'"
        
        WriteInErrorCodeAuto.PrintReport
        Unload WriteInErrorCodeAuto
    Else
        
    End If
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub
'����VIN�����ӡ����
Public Sub printErrCodeByVIN(vin As String)
    On Error Resume Next
    
    'DoEvents
    
    Dim tmpStr As String
    Dim rsDB As New ADODB.Recordset
    rsDB.Fields.Append "name", adBSTR
    rsDB.Open
    rsDB.AddNew
    rsDB("name").value = "name"
    Set WriteInErrorCodeAuto.DataSource = rsDB
    
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim isWriteIn As Boolean
    Dim writeInResult As Boolean
    Dim isPrint As Boolean
    Dim errorCodeList() As String
    Dim rowArr() As String
    Dim i As Integer
    Dim tmpIndex As Integer
    Dim maxID As Integer
    Dim tmp As Integer
    
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select ""VIN"",""ID020"",""ID022"",""ID021"",""ID023"",""WriteInTime"",""ErrorCode"",""MTOC"",""WriteInResult"" from ""T_Result"" where ""VIN""='" & vin & "'")
    
    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        cnn.Close
        Set cnn = Nothing
        Exit Sub
    End If
        
    writeInResult = IIf(IsNull(rs("WriteInResult")), False, CBool(rs("WriteInResult")))
        
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbVIN").Caption = "VIN�룺" & rs("VIN")
    WriteInErrorCodeAuto.Sections("Section1").Controls("lblMTOC").Caption = "MTOC�룺" & rs("MTOC")
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbDateTime").Caption = "���ڣ�" & Format(rs("WriteInTime"), "yyyy-MM-dd HH:mm:ss")
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult").Caption = "���                            " & IIf(writeInResult, "�ϸ�", "���ϸ�")
    
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbLF").Caption = "��ǰ�֣�" & rs("ID022")
    If CStr(rs("ID022")) = "00000000" Or CStr(rs("ID022")) = "" Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbLF").ForeColor = &HFF&
    End If
    
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbRF").Caption = "��ǰ�֣�" & rs("ID020")
    If CStr(rs("ID020")) = "00000000" Or CStr(rs("ID020")) = "" Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbRF").ForeColor = &HFF&
    End If
    
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbLR").Caption = "����֣�" & rs("ID023")
    If CStr(rs("ID023")) = "00000000" Or CStr(rs("ID023")) = "" Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbLR").ForeColor = &HFF&
    End If
    
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbRR").Caption = "�Һ��֣�" & rs("ID021")
    If CStr(rs("ID021")) = "00000000" Or CStr(rs("ID021")) = "" Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbRR").ForeColor = &HFF&
    End If
    
    If Not writeInResult Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult").ForeColor = &HFF&
    End If
    
    If CStr(rs("ErrorCode") & "") = "" Then
        errorCodeList = Split(CStr(rs("ErrorCode") & "&S"), "&S")
    Else
        errorCodeList = Split(CStr(rs("ErrorCode")), "&S")
    End If

    
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (1)).Caption = ""
    WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (1)).Caption = ""
    
    i = 0
    If UBound(errorCodeList) > -1 Then
        For i = 0 To UBound(errorCodeList)
            
            If i <> UBound(errorCodeList) Then
                If Left(errorCodeList(i), 2) = "&P" Then
                    rowArr = Split(CStr(errorCodeList(i)), "&C")
                    rowArr(0) = Replace(rowArr(0), "&P", (i + 1) & " ") '���
                    If rowArr(1) = "ʧ��" Or rowArr(1) = "���ϸ�" Then
                        tmpIndex = tmpIndex + 1
                        If isOnlyPrintNGFlow Then
                            tmp = tmpIndex
                        Else
                            tmp = i + 1
                        End If
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).ForeColor = &HFF&
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).ForeColor = &HFF&
                    End If
                    If Len(rowArr(0)) > 32 Then
                        rowArr(0) = Mid(rowArr(0), 1, 32)
                    End If
                    If isOnlyPrintNGFlow Then
                        If rowArr(1) = "�ɹ�" Then
                            
                        Else
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmpIndex)).Caption = rowArr(0)
                            WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmpIndex)).Caption = rowArr(1)
                        End If
                    Else
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Caption = rowArr(0)
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Caption = rowArr(1)
                    End If
                Else
                    tmpIndex = tmpIndex + 1
                    errorCodeList(i) = "  " & errorCodeList(i)
                    If isOnlyPrintNGFlow Then
                        tmp = tmpIndex
                    Else
                        tmp = i + 1
                    End If
                    WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).ForeColor = &HFF&
                    WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).Top = 15
                    WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (tmp)).Visible = False
                    If Len(errorCodeList(i)) > 32 Then
                        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).Width = 4050
                    End If
                    If Len(errorCodeList(i)) > 36 Then
                        errorCodeList(i) = Mid(errorCodeList(i), 1, 36)
                    End If
                    WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (tmp)).Caption = errorCodeList(i)
                End If
            End If
        Next
        
        If isOnlyPrintNGFlow Then
            i = tmpIndex
        Else
            i = UBound(errorCodeList)
        End If
    End If
    For i = i To 31
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Top = 15
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbError" & (i + 1)).Visible = False
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Top = 15
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbResult" & (i + 1)).Visible = False
    Next i
    
    If isOnlyPrintNGFlow Then
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbErrorEnd").Top = 3000 + tmpIndex * 330
        WriteInErrorCodeAuto.Sections("Section1").Height = 3300 + tmpIndex * 330
    Else
        WriteInErrorCodeAuto.Sections("Section1").Controls("lbErrorEnd").Top = 3000 + UBound(errorCodeList) * 330
        WriteInErrorCodeAuto.Sections("Section1").Height = 3300 + UBound(errorCodeList) * 330
    End If
    
    WriteInErrorCodeAuto.PrintReport
    Unload WriteInErrorCodeAuto
        
    LogWritter "�ֶ���ӡ" & vin & "����Ͻ����Ϣ�ɹ���"
        
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Sub

'******************************************************************************
'** �� �� ����closeAll
'** ��    �룺
'** ��    ����
'** �����������رյ������������ߣ��κε�����������Ҫ�ȵ��ø÷���
'** ȫ�ֱ�����
'** ��    �ߣ�yangshuai
'** ��    �䣺shuaigoplay@live.cn
'** ��    �ڣ�2009-2-27
'** �� �� �ߣ�
'** ��    �ڣ�
'** ��    ����1.0
'******************************************************************************
Public Sub closeAll()
    'oIOCard.OutputController Lamp_Buzzer_IOPort, False '�رշ���
    oIOCard.OutputController Lamp_GreenLight_IOPort, False '�ر���ɫ
    oIOCard.OutputController Lamp_GreenFlash_IOPort, False '�ر���ɫ��˸
    oIOCard.OutputController Lamp_YellowLight_IOPort, False '�رջ�ɫ
    oIOCard.OutputController Lamp_YellowFlash_IOPort, False '�رջ�ɫ��˸
    oIOCard.OutputController Lamp_RedLight_IOPort, False '�رպ�ɫ
    oIOCard.OutputController Lamp_RedFlash_IOPort, False '�رպ�ɫ��˸
End Sub

Public Sub openLamp(IOPort As Integer)
    Call closeAll
    oIOCard.OutputController IOPort, True
End Sub
Public Sub flashLamp(IOPort As Integer)
    Call closeAll
    oIOCard.OutputController IOPort, True
End Sub

Public Sub flashBuzzerLamp(IOPort As Integer)
    Call closeAll
    oIOCard.OutputController Lamp_Buzzer_IOPort, True
    oIOCard.OutputController IOPort, True
End Sub

Public Sub DelayTime(LngTime As Long)
  On Error Resume Next
  Dim LngTick As Long
  LngTick = GetTickCount()
  Do
     DoEvents: DoEvents
  Loop Until (GetTickCount() - LngTick) >= LngTime
End Sub


Function DToB(v As Integer) As String
    If v > 15 Then
        DToB = ""
        Exit Function
    End If
    Select Case v
    Case 0
        DToB = "0000"
    Case 1
        DToB = "0001"
    Case 2
        DToB = "0010"
    Case 3
        DToB = "0011"
    Case 4
        DToB = "0100"
    Case 5
        DToB = "0101"
    Case 6
        DToB = "0110"
    Case 7
        DToB = "0111"
    Case 8
        DToB = "1000"
    Case 9
        DToB = "1001"
    Case 10
        DToB = "1010"
    Case 11
        DToB = "1011"
    Case 12
        DToB = "1100"
    Case 13
        DToB = "1101"
    Case 14
        DToB = "1110"
    Case 15
        DToB = "1111"
    End Select
End Function


Public Sub updateState(key As String, value As String)
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute "update runstate set " & key & "='" & value & "'"
    cnn.Close
    Set cnn = Nothing
End Sub

Public Function readState(key As String) As String
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select * from runstate")
    readState = rs(key)
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Function

Public Sub resetState()
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute "UPDATE runstate SET  test='False', dsgrf=null, dsglf=null, dsgrr=null, dsglr=null,mdlrf=null, mdllf=null, mdlrr=null, mdllr=null,prerf=null, prelf=null, prerr=null, prelr=null,temprf=null, templf=null, temprr=null, templr=null,batteryrf=null, batterylf=null, batteryrr=null, batterylr=null,acspeedrf=null, acspeedlf=null, acspeedrr=null, acspeedlr=null, state=9999"
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub insertColl(code As String)
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute "insert into vincoll(vin) values ('" & code & "')"
    cnn.Close
    Set cnn = Nothing
End Sub

Public Sub delColl(vin As String)
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute "delete from vincoll where vin like '%" & vin & "%'"
    cnn.Close
    Set cnn = Nothing
End Sub
Public Sub delallColl()
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    cnn.Open DBCnnStr
    cnn.Execute "delete from vincoll"
    cnn.Close
    Set cnn = Nothing
End Sub
Public Function getRunStateCar() As CCar
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Set getRunStateCar = New CCar
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select * from runstate")
    
    getRunStateCar.VINCode = IIf(IsNull(rs("vin")), "", rs("vin"))
    
    getRunStateCar.TireRFID = IIf(IsNull(rs("dsgrf")), "", rs("dsgrf"))
    getRunStateCar.TireRFMdl = IIf(IsNull(rs("mdlrf")), "", rs("mdlrf"))
    getRunStateCar.TireRFPre = IIf(IsNull(rs("prerf")), "", rs("prerf"))
    getRunStateCar.TireRFTemp = IIf(IsNull(rs("temprf")), "", rs("temprf"))
    getRunStateCar.TireRFBattery = IIf(IsNull(rs("batteryrf")), "", rs("batteryrf"))
    getRunStateCar.TireRFAcSpeed = IIf(IsNull(rs("acspeedrf")), "", rs("acspeedrf"))
    
    getRunStateCar.TireLFID = IIf(IsNull(rs("dsglf")), "", rs("dsglf"))
    getRunStateCar.TireLFMdl = IIf(IsNull(rs("mdllf")), "", rs("mdllf"))
    getRunStateCar.TireLFPre = IIf(IsNull(rs("prelf")), "", rs("prelf"))
    getRunStateCar.TireLFTemp = IIf(IsNull(rs("templf")), "", rs("templf"))
    getRunStateCar.TireLFBattery = IIf(IsNull(rs("batterylf")), "", rs("batterylf"))
    getRunStateCar.TireLFAcSpeed = IIf(IsNull(rs("acspeedlf")), "", rs("acspeedlf"))
    
    getRunStateCar.TireRRID = IIf(IsNull(rs("dsgrr")), "", rs("dsgrr"))
    getRunStateCar.TireRRMdl = IIf(IsNull(rs("mdlrr")), "", rs("mdlrr"))
    getRunStateCar.TireRRPre = IIf(IsNull(rs("preRR")), "", rs("preRR"))
    getRunStateCar.TireRRTemp = IIf(IsNull(rs("temprr")), "", rs("temprr"))
    getRunStateCar.TireRRBattery = IIf(IsNull(rs("batteryrr")), "", rs("batteryrr"))
    getRunStateCar.TireRRAcSpeed = IIf(IsNull(rs("acspeedrr")), "", rs("acspeedrr"))
    
    getRunStateCar.TireRFID = IIf(IsNull(rs("dsgrf")), "", rs("dsgrf"))
    getRunStateCar.TireRFMdl = IIf(IsNull(rs("mdlrf")), "", rs("mdlrf"))
    getRunStateCar.TireRFPre = IIf(IsNull(rs("preRF")), "", rs("preRF"))
    getRunStateCar.TireRFTemp = IIf(IsNull(rs("temprf")), "", rs("temprf"))
    getRunStateCar.TireRFBattery = IIf(IsNull(rs("batteryrf")), "", rs("batteryrf"))
    getRunStateCar.TireRFAcSpeed = IIf(IsNull(rs("acspeedrf")), "", rs("acspeedrf"))
    
    rs.Close
    Set rs = Nothing
    cnn.Close
    Set cnn = Nothing
End Function

'����VIN����Ų��������ݿ��л�ȡMTOC��
Public Function GetMTOCByVIN(vin As String)
    On Error Resume Next
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    cnn.Open DBCnnStr
    Set rs = cnn.Execute("select mtoc from vinlist where vin='" & vin & "'")
    If Not rs.EOF Then
        GetMTOCByVIN = rs("mtoc")
    Else
        GetMTOCByVIN = ""
    End If
    cnn.Close
    Set rs = Nothing
End Function

'�жϴ�������ѹ�����¶ȡ����ٶ�ֵ�Ƿ���ϱ�׼����ص���״̬
Public Function judgeResultIsOK(value As String, min As String, max As String) As Boolean
On Error Resume Next
    judgeResultIsOK = False
    If CCur(min) <= CCur(value) And CCur(max) >= CCur(value) Then
        judgeResultIsOK = True
    End If
End Function
'�жϴ�����ģʽ�Ƿ�ϸ�
Public Function judgeMdlIsOK(mdl As String, mdlValueArr() As String) As Boolean
    Dim index As Integer
    judgeMdlIsOK = False
    For index = 0 To UBound(mdlValueArr)
        If mdl = mdlValueArr(index) Then
            judgeMdlIsOK = True
            Exit Function
        End If
    Next index
End Function

'�ر�ָ�����ƵĽ���
Public Sub KillProcess(sProcess As String)
    Dim lSnapShot As Long
    Dim lNextProcess As Long
    Dim tPE As PROCESSENTRY32
    lSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If lSnapShot <> -1 Then
    tPE.dwSize = Len(tPE)
    lNextProcess = Process32First(lSnapShot, tPE)
    Do While lNextProcess
    If LCase$(sProcess) = LCase$(Left(tPE.szExeFile, InStr(1, tPE.szExeFile, Chr(0)) - 1)) Then
    Dim lProcess As Long
    Dim lExitCode As Long
    lProcess = OpenProcess(1, False, tPE.th32ProcessID)
    TerminateProcess lProcess, lExitCode
    CloseHandle lProcess
    End If
    lNextProcess = Process32Next(lSnapShot, tPE)
    Loop
    CloseHandle (lSnapShot)
    End If
End Sub


