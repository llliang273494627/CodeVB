Option Strict Off
Option Explicit On
Friend Class frmDateZone
	Inherits System.Windows.Forms.Form
	'******************************************************************************
	'** 文件名：frmDateZone.frm
	'** 版  权：CopyRight (c) 2008-2010 武汉华信数据系统有限公司
	'** 创建人：yangshuai
	'** 邮  箱：shuaigoplay@live.cn
	'** 日  期：2009-2-27
	'** 修改人：
	'** 日  期：
	'** 描  述：时间选择对话框
	'** 版  本：1.0
	'******************************************************************************
	
	'******************************************************************************
	'** 函 数 名：cmdCancel
	'** 输    入：
	'** 输    出：
	'** 功能描述：取消按钮事件响应
	'** 全局变量：
	'** 作    者：yangshuai
	'** 邮    箱：shuaigoplay@live.cn
	'** 日    期：2009-2-27
	'** 修 改 者：
	'** 日    期：
	'** 版    本：1.0
	'******************************************************************************
	Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
		Me.Close()
	End Sub
	
	'******************************************************************************
	'** 函 数 名：cmdSaveAs_Click
	'** 输    入：
	'** 输    出：
	'** 功能描述：导出按钮事件响应
	'** 全局变量：
	'** 作    者：yangshuai
	'** 邮    箱：shuaigoplay@live.cn
	'** 日    期：2009-2-27
	'** 修 改 者：
	'** 日    期：
	'** 版    本：1.0
	'******************************************************************************
	Private Sub cmdSaveAs_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSaveAs.Click
		Dim sqlText As String
		Dim lowDate As Object
		Dim highDate As Object
		
		'UPGRADE_WARNING: 未能解析对象 lowDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		lowDate = CDate(Me.dtpLow.value)
		'UPGRADE_WARNING: 未能解析对象 highDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		highDate = CDate(Me.dtpHigh.value)
		
		'UPGRADE_WARNING: 未能解析对象 highDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		'UPGRADE_WARNING: 未能解析对象 lowDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		If lowDate > highDate Then
			MsgBox(" ")
			Exit Sub
		End If
		
		'UPGRADE_WARNING: 未能解析对象 highDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		'UPGRADE_WARNING: 未能解析对象 lowDate 的默认属性。 单击以获得更多信息:“ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"”
		sqlText = "select ""VIN"",""VIS"",""ID020"" as ""右前轮ID"",""Mdl020"" as ""右前轮模式"",""Pre020"" as ""右前轮压力"",""Temp020"" as ""右前轮温度"",""Battery020"" as ""右前轮电池"",""AcSpeed020"" as ""右前轮加速度"" ,""ID021"" as ""右后轮ID"",""Mdl021"" as ""右后轮模式"",""Pre021"" as ""右后轮压力"",""Temp021"" as ""右后轮温度"",""Battery021"" as ""右后轮电池"",""AcSpeed021"" as ""右后轮加速度"" ,""ID022"" as ""左前轮ID"",""Mdl022"" as ""左前轮模式"",""Pre022"" as ""左前轮压力"",""Temp022"" as ""左前轮温度"",""Battery022"" as ""左前轮电池"",""AcSpeed022"" as ""左前轮加速度"" ,""ID023"" as ""左后轮ID"" ,""Mdl023"" as ""左后轮模式"",""Pre023"" as ""左后轮压力"",""Temp023"" as ""左后轮温度"",""Battery023"" as ""左后轮电池"",""AcSpeed023"" as ""左后轮加速度"" ,""TestTime"" as ""测试时间"",""WriteInTime"" as ""写入时间"" from " & " ""T_Result"" where   ""TestTime"">='" & lowDate & "' and ""TestTime""<='" & highDate & "'"
		
		'组合导出查询语句，调用导出函数
		
		exportExcel(sqlText)
		
		
	End Sub
	
	
	
	'******************************************************************************
	'** 函 数 名：Form_Load
	'** 输    入：
	'** 输    出：
	'** 功能描述：窗体加载时间响应
	'** 全局变量：
	'** 作    者：yangshuai
	'** 邮    箱：shuaigoplay@live.cn
	'** 日    期：2009-2-27
	'** 修 改 者：
	'** 日    期：
	'** 版    本：1.0
	'******************************************************************************
	Private Sub frmDateZone_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
		'控件风格XP化
		WindowsXPC1.InitSubClassing()
		
		'界面控件控制
		dtpLow.value = DateAdd(Microsoft.VisualBasic.DateInterval.Day, -7, Today)
		dtpHigh.value = Today
		
		Me.Left = VB6.TwipsToPixelsX((VB6.PixelsToTwipsX(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width) - VB6.PixelsToTwipsX(Me.Width)) / 2)
		Me.Top = VB6.TwipsToPixelsY((VB6.PixelsToTwipsY(System.Windows.Forms.Screen.PrimaryScreen.Bounds.Height) - VB6.PixelsToTwipsY(Me.Height)) / 2)
	End Sub
End Class