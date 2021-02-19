Attribute VB_Name = "modIO"
'******************************************************************************
'** 文件名：modIO.bas
'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
'** 创建人：yangshuai
'** 邮  箱：shuaigoplay@live.cn
'** 日  期：2009-2-27
'** 修改人：
'** 日  期：
'** 描  述：IO卡控制模块模块
'** 版  本：1.0
'******************************************************************************


Option Explicit
'******************************************************************************
'** 函 数 名：OutputController
'** 输    入：portNum——端口号（0-15）；state——开关状态（true=开，false=关）
'** 输    出：
'** 功能描述：IO卡输出控制
'** 全局变量：
'** 作    者：yangshuai
'** 邮    箱：shuaigoplay@live.cn
'** 日    期：2009-2-27
'** 修 改 者：
'** 日    期：
'** 版    本：1.0
'******************************************************************************
Private Sub OutputController(portNum As Integer, state As Boolean)
    Dim str As String
    If state Then
        str = "打开"
    Else
        str = "关闭"
    End If
    Debug.Print str & portNum + 1 & "号端口"
End Sub
