VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'** 文件名：Clog.cls
'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
'** 创建人：李操
'** 邮  箱：tonylicao@163.com
'** 日  期：2008-08-18
'** 修改人：
'** 日  期：
'** 描  述：日志模块记录程序运行中的异常信息和操作记录。这些信息都被以.txt的文本形式保存在本地路径上。
'**         需要引用Microsoft Scripting Runtime
'** 版  本：1.0
'******************************************************************************


Option Explicit

Private m_UseTime As Boolean
Private m_LogPath As String





'******************************************************************************
'** 函 数 名：UseTime
'** 输    入：Boolean型
'** 输    出：Boolean型
'** 功能描述：Clog类的属性，用于控制所写日志是否带上当前时间。为真时带时间；为假时不带时间。
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com
'** 日    期：2008-08-18
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Public Property Get UseTime() As Boolean
    UseTime = m_UseTime
End Property

Public Property Let UseTime(ByVal vUseTime As Boolean)
    m_UseTime = vUseTime
End Property


'******************************************************************************
'** 函 数 名：LogPath
'** 输    入：String型
'** 输    出：String型
'** 功能描述：Clog类的属性，日志文件的保存位置。设置后如果没有发现则自动创建目录。
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com
'** 日    期：2008-08-18
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Public Property Get LogPath() As String
    LogPath = m_LogPath
End Property

Public Property Let LogPath(ByVal vLogPath As String)
    m_LogPath = vLogPath
End Property


Private Sub Class_Initialize()
    m_UseTime = True
    m_LogPath = ""
End Sub


'******************************************************************************
'** 函 数 名：WriteErrInfo
'** 输    入：参数为String型：ModuleName—异常所在的模块名；SubName—异常所在的方法名；ErrInfo—异常信息。
'** 输    出：返回值为Boolean型：为真时写日志成功；为假时失败。
'** 功能描述：Clog类的方法，用于写异常日志。
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com
'** 日    期：2008-08-18
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Public Function WriteErrInfo(ByVal ModuleName As String, ByVal SubName As String, ByVal ErrInfo As String) As Boolean
    Dim FSO As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    Dim typeid As Integer
    Dim NowOutputDir As String
    Dim ArrNowOutput() As String
    Dim NumArray As Integer, i As Integer
    
    m_LogPath = App.Path & "\Log\" & Format(Trim(Date), "yyyy-mm-dd") & ".txt"
    NowOutputDir = m_LogPath
    ArrNowOutput = Split(NowOutputDir, "\")
    NumArray = UBound(ArrNowOutput)
    If FSO.DriveExists(ArrNowOutput(0)) = False Then           '判断盘符是否存在，不存在向上抛出错误
        Err.Raise 10021, "路径", "创建文件夹或文件失败"
        Exit Function
    End If
    
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    NowOutputDir = ""
    For i = 0 To NumArray
        If i = NumArray Then
            NowOutputDir = NowOutputDir & ArrNowOutput(i)
            If Trim(Dir(NowOutputDir)) = "" Then
                FSO.CreateTextFile NowOutputDir, False
            End If
        Else
            NowOutputDir = NowOutputDir & ArrNowOutput(i) & "\"
            If Trim(Dir(NowOutputDir, vbDirectory)) = "" Then
                FSO.CreateFolder NowOutputDir
            End If
        End If
    Next
    
    On Error Resume Next
    Set fil = FSO.GetFile(NowOutputDir)
    Set ts = fil.OpenAsTextStream(ForAppending)
    
    Err.Clear
    If m_UseTime = True Then
        ts.Write "[" & Now() & "]" & "在" & ModuleName & "模块，" & SubName & "方法中，发生了" & ErrInfo & "异常！" & vbCrLf
    Else
        ts.Write "在" & ModuleName & "模块，" & SubName & "方法中，发生了" & ErrInfo & "异常！" & vbCrLf
    End If
    If Err.Number = 0 Then
        WriteErrInfo = True
    Else
        WriteErrInfo = False
    End If

    Set FSO = Nothing
    ts.Close
End Function

'******************************************************************************
'** 函 数 名：WriteOprInfo
'** 输    入：参数为String型：OprInfo—操作信息
'** 输    出：返回值为Boolean型：为真时写日志成功；为假时失败。
'** 功能描述：Clog类的方法，用于写操作日志。
'** 全局变量：
'** 调用模块：
'** 作    者：李操
'** 邮    箱：tonylicao@163.com
'** 日    期：2008-08-18
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************

Public Function WriteOprInfo(ByVal OprInfo As String) As Boolean
    Dim FSO As New FileSystemObject
    Dim fil As File
    Dim ts As TextStream
    Dim typeid As Integer
    Dim NowOutputDir As String
    Dim ArrNowOutput() As String
    Dim NumArray As Integer, i As Integer
    
    m_LogPath = App.Path & "\Log\" & Format(Trim(Date), "yyyy-mm-dd") & ".txt"
    NowOutputDir = m_LogPath
    ArrNowOutput = Split(NowOutputDir, "\")
    NumArray = UBound(ArrNowOutput)
    If FSO.DriveExists(ArrNowOutput(0)) = False Then          '判断盘符是否存在，不存在向上抛出错误
        Err.Raise 10022, "路径", "创建文件夹或文件失败"
        Exit Function
    End If
    
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    NowOutputDir = ""
    For i = 0 To NumArray
        If i = NumArray Then
            NowOutputDir = NowOutputDir & ArrNowOutput(i)
            If Trim(Dir(NowOutputDir)) = "" Then
                FSO.CreateTextFile NowOutputDir, False
            End If
        Else
            NowOutputDir = NowOutputDir & ArrNowOutput(i) & "\"
            If Trim(Dir(NowOutputDir, vbDirectory)) = "" Then
                FSO.CreateFolder NowOutputDir
            End If
        End If
    Next
    
    On Error Resume Next
    Set fil = FSO.GetFile(NowOutputDir)
    Set ts = fil.OpenAsTextStream(ForAppending)
    
    Err.Clear
    If m_UseTime = True Then
        ts.Write "[" & Now() & "]" & OprInfo & vbCrLf
    Else
        ts.Write OprInfo & vbCrLf
    End If
    If Err.Number = 0 Then
        WriteOprInfo = True
    Else
        WriteOprInfo = False
    End If
    
    Set FSO = Nothing
    ts.Close
End Function
