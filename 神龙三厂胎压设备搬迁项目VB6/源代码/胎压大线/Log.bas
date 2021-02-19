Attribute VB_Name = "log"
'///////////////////////////////日志////////////////////////////////////

'////////////////////////START///////////////////////////////
'可以写入数据，追加模式
'默认写到当前目录下Log目录内，以当前日期命名的txt文件
Public Sub LogWritter(txt As String) '写日志,追加模式,
    Dim FSO As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    Dim typeid As Integer
    Dim NowOutput As String, NowOutputDir As String
    
On Error Resume Next
    
    NowOutputDir = App.Path & "\Log"
    NowOutput = App.Path & "\Log\" & Trim(Date) & ".txt"
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    If Trim(Dir(NowOutputDir, vbDirectory)) = "" Then
        FSO.CreateFolder NowOutputDir
    End If
    If Trim(Dir(NowOutput)) = "" Then
        FSO.CreateTextFile NowOutput, False
    End If
    
    Set fil = FSO.GetFile(NowOutput)
    Set ts = fil.OpenAsTextStream(ForAppending)
    
    ts.Write "[" & Now() & "]" & txt & vbCrLf
    ts.Close
End Sub
'////////////////////////END/////////////////////////////////
Public Sub SensorLogWritter(txt As String) '写日志,追加模式,
    Dim FSO As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    Dim typeid As Integer
    Dim NowOutput As String, NowOutputDir As String
    
On Error Resume Next
    
    NowOutputDir = App.Path & "\SensorLog"
    NowOutput = App.Path & "\SensorLog\" & Trim(Date) & ".txt"
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    If Trim(Dir(NowOutputDir, vbDirectory)) = "" Then
        FSO.CreateFolder NowOutputDir
    End If
    If Trim(Dir(NowOutput)) = "" Then
        FSO.CreateTextFile NowOutput, False
    End If
    
    Set fil = FSO.GetFile(NowOutput)
    Set ts = fil.OpenAsTextStream(ForAppending)
    
    ts.Write "[" & Now() & "]" & txt & vbCrLf
    ts.Close
End Sub
'////////////////////////START///////////////////////////////
'可以写入数据，覆盖模式,可以指定路径，但必需保证该目录存在
'如果不指定路径，则默认写到当前目录下Output.txt文件

Public Sub DataWritter(txt As String, OutputPath As String)
    Dim FSO As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    Dim typeid As Integer
    Dim NowOutput As String
    
On Error Resume Next
    
    If OutputPath <> "" Then
        NowOutput = OutputPath
    Else
        NowOutput = App.Path & "\Output.txt"
    End If
    Set FSO = CreateObject("Scripting.Filesystemobject")

    If Trim(Dir(NowOutput)) = "" Then
        FSO.CreateTextFile NowOutput, False
    End If
    
    Set fil = FSO.GetFile(NowOutput)
    Set ts = fil.OpenAsTextStream(ForWriting)
    
    ts.Write ""
    ts.Write txt
    ts.Close
End Sub
'////////////////////////END/////////////////////////////////

