Attribute VB_Name = "ModGlobel"
Public AppPath As String 'Ӧ�ó���·��
Public AppName As String 'Ӧ�ó�������
Public PswMode As String '���ڱ�ʶ������ȷ����󵯳����������

Public WriteFlag As Integer '��ʶд��״̬��0δд�� 1����д�� -1д�뷢������

Public LocalDBConnStr As String '�������ݿ������ַ���
Public RemoteDBConnStr As String 'Զ�����ݿ������ַ���

Public oIOCard As IOControl.IOCard  'IO���ƶ���

'***************************************************************************
' ��ʾ������ʾ��Ϣ��
'***************************************************************************
Public Sub PopMsg(strTitle As String, strMsg As String)
    If Trim(strtile) = "" Then
        strTitle = "��ʾ��Ϣ"
    End If
    FrmMsg.LbTitle = strTitle
    FrmMsg.LbMsg = strMsg
    FrmMsg.Show 1
End Sub


'**************************************************************************
' ���VIN�Ϸ���
' ����17λ�������ܿ���չ��
'**************************************************************************
Public Function CheckVin(ByVal strVin As String) As Boolean
On Error GoTo CheckVinErr
    Dim Result As Boolean
    
    Result = True
    
    '��鳤���Ƿ�Ϊ17λ�����򷵻�ֵresult=false
    If Len(strVin) <> 17 Then
        Result = False
    End If
    
    CheckVin = Result
    Exit Function
CheckVinErr:
    CheckVin = False
    Exit Function
End Function

'***************************************************************************
' ��鳵��ID�Ϸ���
'***************************************************************************
Public Function CheckWheelID(ByVal strlf As String, ByVal strlb As String, _
                    ByVal strrf As String, ByVal strrb As String) As Boolean
On Error GoTo CheckWheelErr
    Dim Result As Boolean
    
    Result = True
    If strlf = "" Or strlb = "" Or strrf = "" Or strrb = "" Then
        Result = False
    End If
    
    CheckWheelID = Result
    Exit Function
CheckWheelErr:
    CheckWheelID = False
    Exit Function
End Function


Public Function GenerateKey(ByRef seed As String, seednumber As String, _
        ByRef key As String, keynumber As Integer, pin As Integer)
        
        pin = 13151720
        Dim i As Integer
        Dim j As Integer
        
        i = 0: j = 0
        
        For i = 0 To seednumber
            pin = pin ^ (LeftMove(pin, 5) + Asc(Mid(seed, i + 1, 1)) + RightMove(pin, 4))
        Next i
        
        keynumber = seednumber
        
        For j = 0 To keynumber
            key = Replace(key, Mid(key, j + 1, 1), Chr((RightMove(pin, j * 8) And &HFF)))
        Next j
End Function


Function LeftMove(ByVal iValue As Integer, iStep As Integer) As Integer
    Dim i As Integer
    For i = 1 To iStep
        iValue = iValue * 2
    Next i
    LeftMove = iValue
End Function

Function RightMove(ByVal iValue As Integer, iStep As Integer) As Integer
    Dim i As Integer
    For i = 1 To iStep
        iValue = iValue / 2
    Next i
    RightMove = iValue
End Function

Public Sub Main()
On Error GoTo Main_Err
    Set oIOCard = New IOControl.IOCard
    frmMain.Show
    Exit Sub
Main_Err:
    MsgBox "��ʼ������ʧ�ܣ�������Ϣ��" & Err.Description & "������������Ϣ��"
End Sub
