Attribute VB_Name = "log"
'///////////////////////////////��־////////////////////////////////////

'////////////////////////START///////////////////////////////
'����д�����ݣ�׷��ģʽ
'Ĭ��д����ǰĿ¼��LogĿ¼�ڣ��Ե�ǰ����������txt�ļ�
Public Sub LogWritter(txt As String) 'д��־,׷��ģʽ,
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
Public Sub SensorLogWritter(txt As String) 'д��־,׷��ģʽ,
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
'����д�����ݣ�����ģʽ,����ָ��·���������豣֤��Ŀ¼����
'�����ָ��·������Ĭ��д����ǰĿ¼��Output.txt�ļ�

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

