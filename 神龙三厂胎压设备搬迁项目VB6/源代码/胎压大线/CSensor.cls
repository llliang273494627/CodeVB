VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSensor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents IOC As IOControl.IOCard
Attribute IOC.VB_VarHelpID = -1
Event onChange(state As Boolean)
Public m_IOPort As Integer
Private m_State As Boolean

Private Sub Class_Initialize()
    Set IOC = oIOCard
End Sub



Private Sub IOC_EventTest(testPort() As Boolean)
    If m_State <> testPort(m_IOPort) Then
        m_State = testPort(m_IOPort)
        RaiseEvent onChange(m_State)
    End If
End Sub



Public Property Let IOPort(ByVal Num As Integer)
    Dim col As Collection
    m_IOPort = Num
    Set col = IOC.getState
    m_State = col(m_IOPort + 1)
End Property


'得到当前状态
Public Property Get state() As Variant
    state = m_State
End Property


