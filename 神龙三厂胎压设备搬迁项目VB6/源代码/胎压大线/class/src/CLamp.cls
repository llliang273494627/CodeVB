VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

Private m_IOGPort As Integer

Private m_IOYPort As Integer

Private m_IORPort As Integer

Private m_IOBPort As Integer

Private m_IOFlashPort As Integer

Private m_IOBuzzerPort As Integer

'state参数表示将把灯设置成的状态
'0=关
'1=开
'2=闪烁

Public Sub GreenLight(state As Integer)
End Sub


Public Sub RedLight(state As Intege)
End Sub


Public Sub BlueLight(state As Intege)
End Sub


Public Sub YellowLight(state As Intege)
End Sub
Public Sub Buzzer(state As Intege)
End Sub

Public Property Get IOGPort() As Integer
    IOPort = m_IOPort
End Property


Public Property Let IOGPort(ByVal value As Integer)
    m_IOPort = value
End Property


Public Property Get IOYPort() As Integer
    Property1 = m_Property1
End Property


Public Property Let IOYPort(ByVal value As Integer)
    m_Property1 = value
End Property


Public Property Get IORPort() As Integer
    Property1 = m_Property1
End Property


Public Property Let IORPort(ByVal value As Integer)
    m_Property1 = value
End Property


Public Property Get IOBPort() As Integer
    IOBPort = m_IOBPort
End Property


Public Property Let IOBPort(value As Integer)
    m_IOBPort = value
End Property


Public Property Get IOFlashPort() As Integer
    Property1 = m_Property1
End Property


Public Property Let IOFlashPort(ByVal value As Integer)
    m_Property1 = value
End Property


Public Property Get IOBuzzerPort() As Integer
    Property1 = m_Property1
End Property


Public Property Let IOBuzzerPort(ByVal value As Integer)
    m_Property1 = value
End Property


