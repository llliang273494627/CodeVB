VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CFY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim stableName As String
Dim skeyField As String
Dim CnnStr As String
Dim sselectMember As String
Dim swhereMenber As String
Dim spageCount As Long
Dim spaginalRecordCount As Long
Dim spageNum As Long
Dim SqlStr As String
Dim cnn As ADODB.Connection
Dim srecordCount As Long
Dim rs As ADODB.Recordset

Public Property Let ConnectionString(ByVal vNewValue As String)
    On Error Resume Next
    CnnStr = vNewValue
    cnn.Open CnnStr

    If Err.Number <> 0 Then
        Err.Raise 400000, , "连接错误！错误信息：" & Err.Description
    End If
    
End Property

Public Property Get tableName() As String
    tableName = stableName
End Property

Public Property Let tableName(ByVal vNewValue As String)
    stableName = vNewValue
End Property
Public Property Get PageNum() As Long
    PageNum = spageNum
End Property

Public Property Let PageNum(ByVal vNewValue As Long)
    spageNum = vNewValue
End Property

Public Property Get KeyField() As String
    KeyField = skeyField
End Property

Public Property Let KeyField(ByVal vNewValue As String)
    skeyField = vNewValue
End Property


Public Property Let SelectMember(ByVal vNewValue As String)
    sselectMember = vNewValue
End Property

Public Property Let WhereMenber(ByVal vNewValue As String)
    swhereMenber = vNewValue
End Property


Public Property Get SelectSqlStr() As String
     SelectSqlStr = SqlStr
End Property
Public Property Get PageCount() As Long
     PageCount = spageCount
End Property
Public Property Let PaginalRecordCount(ByVal vNewValue As Long)
    spaginalRecordCount = vNewValue
End Property
Public Property Get PaginalRecordCount() As Long
    PaginalRecordCount = spaginalRecordCount
End Property
Public Property Get RecordCount() As Long
    RecordCount = srecordCount
End Property

Public Sub getRecordSet(ByRef record As ADODB.Recordset)
    On Error GoTo getRecordSet_Err
    If stableName = "" Then
        'Err.Raise 270002, "", "属性TableName没有配置！"
        Exit Sub
    End If
    If skeyField = "" Then
        'Err.Raise 270003, "", "属性KeyField没有配置！"
        Exit Sub
    End If
    
    Dim tmpSqlStr As String
    Dim tmpArr()
    Set record = Nothing
    tmpSqlStr = "select " & skeyField & " from " & stableName & " where 1=1 " & swhereMenber

   
    Set rs = cnn.Execute(tmpSqlStr)
    
    
    If Not rs.EOF Then
        tmpArr = rs.GetRows
        srecordCount = UBound(tmpArr, 2) + 1
    Else
        Set record = rs
        SqlStr = ""
        Exit Sub
    End If
    
    Dim remainder As Long
    remainder = srecordCount Mod spaginalRecordCount
    If remainder = 0 Then
        spageCount = srecordCount \ spaginalRecordCount
    Else
        spageCount = (srecordCount \ spaginalRecordCount) + 1
    End If
        
    Dim fromInt As Long
    Dim toInt As Long
    
    If spageNum < spageCount Then
        fromInt = spaginalRecordCount * (spageNum - 1)
        toInt = spaginalRecordCount * spageNum - 1
    ElseIf spageNum = spageCount Then
        fromInt = spaginalRecordCount * (spageNum - 1)
        toInt = UBound(tmpArr(), 2)
    ElseIf spageNum > spageCount Then
        'Err.Raise 270002, , "属性spageNum(当前页数)大于总页数！"
        Exit Sub
    End If
    
    Dim i As Long
    Dim inINStr  As String
    inINStr = ""
    For i = fromInt To toInt
        inINStr = inINStr & tmpArr(0, i) & ","
    Next
    inINStr = Left(inINStr, Len(inINStr) - 1)
    SqlStr = "select " & sselectMember & " from " & stableName & " where " & skeyField & " in (" & inINStr & ")"
    Set rs = cnn.Execute(SqlStr)
    
    Set record = rs
 
    Exit Sub
getRecordSet_Err:
    MsgBox Err.Description
End Sub

Private Sub Class_Initialize()
    Set cnn = New ADODB.Connection
    spaginalRecordCount = 50
    sselectMember = "*"
    srecordCount = 0
    spageNum = 1
End Sub

Private Sub Class_Terminate()
    If cnn.state = 1 Then
        cnn.Close
    End If
    Set cnn = Nothing
    If rs Is Nothing Then
        Exit Sub
    End If
    If rs.state = 1 Then
        rs.Close
    End If
    
    Set rs = Nothing
End Sub

