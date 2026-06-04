systemutil.Run "C:\Users\yll24\Desktop\结课大作业\2023592003 杨龙龙 软件测试大作业\测试程序.exe"
datatable.ImportSheet "C:\Users\yll24\Desktop\结课大作业\2023592003 杨龙龙 软件测试大作业\测试数据.xlsx",1,"Action1" @@ hightlight id_;_70660_;_script infofile_;_ZIP::ssf24.xml_;_
wait(1)
Dim Dormitory,electricity,water,expected,actual
For Iterator = 1 To 12 Step 1
	Dormitory=datatable("楼栋","Action1")
    electricity=datatable("用电量","Action1")
    water=datatable("用水量","Action1")
    expected=datatable("预期结果","Action1")
 @@ hightlight id_;_1183032_;_script infofile_;_ZIP::ssf40.xml_;_
    If Dormitory = "1号宿舍楼" Then
		Window("校园宿舍水电费计算小程序").WinObject("1号宿舍楼").Click
	ElseIf Dormitory = "2号宿舍楼" Then
		Window("校园宿舍水电费计算小程序").WinObject("2号宿舍楼").Click
	ElseIf Dormitory = "3号宿舍楼" Then
		Window("校园宿舍水电费计算小程序").WinObject("3号宿舍楼").Click
	ElseIf Dormitory = "研究生公寓" Then
		Window("校园宿舍水电费计算小程序").WinObject("研究生公寓").Click
	End If
    
    Window("校园宿舍水电费计算小程序").WinObject("本月用电量").Click
	Window("校园宿舍水电费计算小程序").WinObject("TkChild").Type electricity
	Window("校园宿舍水电费计算小程序").WinObject("本月用水量").Click
	Window("校园宿舍水电费计算小程序").WinObject("TkChild").Type water
	Window("校园宿舍水电费计算小程序").WinObject("计算费用").Click
	wait(1)
	
	If Window("校园宿舍水电费计算小程序").Dialog("计算结果").Exist(1) Then
		actual=Window("校园宿舍水电费计算小程序").Dialog("计算结果").Static("计算费用").GetROProperty("text")
		Window("校园宿舍水电费计算小程序").Dialog("计算结果").WinButton("确定").Click
	ElseIf Window("校园宿舍水电费计算小程序").Dialog("错误").Exist(1) Then
		actual=Window("校园宿舍水电费计算小程序").Dialog("错误").Static("错误信息").GetROProperty("text")
		Window("校园宿舍水电费计算小程序").Dialog("错误").WinButton("确定").Click
	Else 
		actual=Window("校园宿舍水电费计算小程序").Dialog("提示").Static("提示信息").GetROProperty("text")
		Window("校园宿舍水电费计算小程序").Dialog("提示").WinButton("确定").Click
	End If
	Window("校园宿舍水电费计算小程序").WinObject("重置清空").Click
	wait(1)
	
	datatable("实际结果","Action1") = actual
	If expected = actual Then
		datatable("是否通过","Action1") = "通过"
	else
		datatable("是否通过","Action1") = "未通过"
	End If
	datatable.GetSheet("Action1").SetNextRow
	wait(1)
	
Next
systemutil.CloseProcessByName "CostLiving.exe" @@ hightlight id_;_461152_;_script infofile_;_ZIP::ssf78.xml_;_
