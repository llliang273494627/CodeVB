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
'******************************************************************************
'** 文件名：CLamp.cls
'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
'** 创建人：yangshuai
'** 邮  箱：shuaigoplay@live.cn
'** 日  期：2009-2-27
'** 修改人：
'** 日  期：
'** 描  述：信号灯控制
'** 版  本：1.0
'******************************************************************************

Option Explicit

Private m_IOGPort As Integer '绿灯IO输出口

Private m_IOYPort As Integer '黄灯IO输出口

Private m_IORPort As Integer '红灯IO输出口

Private m_IOBPort As Integer '蓝灯IO输出口

Private m_IOFlashPort As Integer '闪烁IO输出口

Private m_IOBuzzerPort As Integer '蜂鸣IO输出口



'state参数表示将把灯设置成的状态
'0=关
'1=开
'2=闪烁

'******************************************************************************
'** 函 数 名：GreenLight
'** 输    入：
'** 输    出：
'** 功能描述：绿色信号灯控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub GreenLight(state As Integer)
    Select Case state
    Case 0
        
    Case 1
    
    Case 2
    
    End Select
End Sub

'******************************************************************************
'** 函 数 名：Form_Load
'** 输    入：
'** 输    出：
'** 功能描述：黄色信号灯控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub RedLight(state As Intege)

End Sub

'******************************************************************************
'** 函 数 名：Form_Load
'** 输    入：
'** 输    出：
'** 功能描述：红色信号灯控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub BlueLight(state As Intege)
End Sub

'******************************************************************************
'** 函 数 名：Form_Load
'** 输    入：
'** 输    出：
'** 功能描述：蓝色信号灯控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Public Sub YellowLight(state As Intege)
End Sub

'******************************************************************************
'** 函 数 名：Form_Load
'** 输    入：
'** 输    出：
'** 功能描述：蜂鸣器控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
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



