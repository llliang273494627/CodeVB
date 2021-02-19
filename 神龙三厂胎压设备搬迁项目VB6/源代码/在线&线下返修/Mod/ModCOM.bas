Attribute VB_Name = "ModCOM"
Option Explicit
Public Declare Function GetTickCount Lib "KERNEL32" () As Long
Public VIN As String
Public RedoDev1 As Boolean
Public RedoDev2 As Boolean
Public RedoDev3 As Boolean
Public Dev1VIN As String   '设备1VIN
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
'** 函 数 名：exportExcel
'** 输    入：sqlText——sql语句
'** 输    出：
'** 功能描述：导出sql查询语句的查询结果
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-28
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Public Sub exportExcel(sqlText As String)
    Dim excelzfc As String
    Dim fileName As String
    Dim FSO As FileSystemObject
    Dim txtfile As TextStream
    Dim cnn As New ADODB.Connection
    Dim rs As ADODB.Recordset
On Error GoTo exportExcel_ERR
    fileName = getExcelFileName '得到随机文件名
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
    
    
    '构造表头
    Dim I As Integer
    Dim tmp As String
    For I = 0 To rs.Fields.Count - 1
         tmp = tmp & rs.Fields(I).Name & Chr(9)
    Next
    txtfile.WriteLine tmp
    
    '构造表格内部
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
    
    '打开excel
    Dim xlApp, xlbook, db1, xlsheet
    Set xlApp = CreateObject("Excel.Application")
    xlApp.DisplayAlerts = False '不显示警告
    xlApp.Application.Visible = True '不显示界面
    Set xlbook = xlApp.Workbooks.Open(fileName)
    Exit Sub
exportExcel_ERR:
    MsgBox "数据导出Excel出错，错误信息：" & err.Description
End Sub




'******************************************************************************
'** 函 数 名：getExcelFileName
'** 输    入：
'** 输    出：生成的新的excel文件名
'** 功能描述：生成的新的excel文件名 年+月+日+时+分+秒+1000内随机数.xls
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-28
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Function getExcelFileName() As String
    Dim MyValue As String
    Randomize   ' 对随机数生成器做初始化的动作。
    MyValue = Format(Int((1000 * Rnd) + 1), "0000")  ' 生成 1 到 1000 之间的随机数值。
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
