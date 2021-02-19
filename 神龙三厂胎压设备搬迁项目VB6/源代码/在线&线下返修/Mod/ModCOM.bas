Attribute VB_Name = "ModCOM"
Option Explicit
Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public VIN As String
Public RedoDev1 As Boolean
Public RedoDev2 As Boolean
Public RedoDev3 As Boolean
Public Dev1VIN As String   '�豸1VIN
Public Dev2VIN As String
Public Dev3VIN As String
Public gComResp As Long
Public PostgresStr As String
Public RMTPostgresStr As String
Public AdmkState As Integer
Public Const COM_GOTNAK = 0
Public Const COM_OTHER = 1
'Public gDataStay As Long
Public Const gDataStay = 2000
Public blnOpenstat As Boolean
Public Derult As String
Public strin As String
Public Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type
Public Declare Function SetCommTimeouts Lib "KERNEL32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public Declare Function GetCommTimeouts Lib "KERNEL32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Public timeouts As COMMTIMEOUTS
Public Sub DelayTime(LngTime As Long)
  Dim LngTick As Long
  LngTick = GetTickCount()
  Do
   
     DoEvents: DoEvents
  Loop Until (GetTickCount() - LngTick) >= LngTime
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
On Error GoTo exportExcel_ERR
    fileName = getExcelFileName '�õ�����ļ���
    cnn.Open PostgresStr
    Set rs = cnn.Execute(sqlText)
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set txtfile = FSO.CreateTextFile(fileName, True)
    
'    For I = 0 To Me.MSFlexGrid1.Rows - 1
'        For J = 1 To Me.MSFlexGrid1.Cols - 1
'            excelzfc = excelzfc & MSFlexGrid1.TextMatrix(I, J) & Chr(9)
'        Next
'        txtfile.WriteLine
'    Next
    
    
    '�����ͷ
    Dim I As Integer
    Dim tmp As String
    For I = 0 To rs.Fields.Count - 1
         tmp = tmp & rs.Fields(I).Name & Chr(9)
    Next
    txtfile.WriteLine tmp
    
    '�������ڲ�
    Do While Not rs.EOF
        tmp = ""
        For I = 0 To rs.Fields.Count - 1
            tmp = tmp & rs(rs.Fields(I).Name).value & Chr(9)
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
    MsgBox "���ݵ���Excel����������Ϣ��" & err.Description
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
