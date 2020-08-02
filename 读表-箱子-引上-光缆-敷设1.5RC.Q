[General]
SyntaxVersion=2
BeginHotkey=120
BeginHotkeyMod=2
PauseHotkey=121
PauseHotkeyMod=2
StopHotkey=122
StopHotkeyMod=2
RunOnce=1
EnableWindow=
MacroID=3af7ec5b-bd99-409e-8f3e-7363b6a6e67c
Description=读表-箱子-引上-光缆-敷设1.5RC
Enable=0
AutoRun=0
[Repeat]
Type=0
Number=1
[SetupUI]
Type=2
QUI=
[Relative]
SetupOCXFile=
[Comment]

[Script]
//bug:新建引上系统不输入区县
//Todo:优化坐标
Dim Draw_Vector(100, 10, 1) //划线向量,角标从1开始,x条,y段,z0为x,z1为y
Dim FSB_CC() //分纤箱循环次数,角标从1开始,FSB=Fiber_Separate_Box,CC=Circulation_Count
Dim OCS_CC() //光缆段循环次数,角标从1开始,Ubound-1为光缆条数
Dim FSB_Info(100,100,10) //分纤箱信息,角标从1开始
Dim OC_Info(100,10,10) //光缆信息,角标从1开始,x条,y段,z属性
Dim Other_Info(14) //其他信息,角标从1开始

Delay_Unit=150 //延迟常量
C_End=100 //循环终止常量
Fiber_Count_Constant    = 1    //OC_Info 芯数常量
Lenth_Constant          = 2    //OC_Info 长度常量
Type_Constant           = 3    //OC_Info 类型常量
Vector_Constant         = 4    //OC_Info 向量常量
Start_Point_Constant    = 5    //OC_Info 起点常量

Other_Info_Data_Column      =  2 //Other_Info 数据列
RD_Name_Constant            =  1 //Other_Info 小区名称常量
SS_Name_Constant            =  2 //Other_Info 承载系统名称常量
Project_Num_Constant        =  3 //Other_Info 项目编号常量
Task_Name_Constant          =  4 //Other_Info 任务名称常量
DQS_Porject_Constant        =  5 //Other_Info 数据质量责任人_工程常量
DQS_City_Constant           =  6 //Other_Info 数据质量责任人_地市常量
DQS_District_Constant       =  7 //Other_Info 数据质量责任人_区县常量
DQS_Maintenance_Constant    =  8 //Other_Info 数据质量责任人_维护常量
District_Constant           =  9 //Other_Info 区县常量
Horizontal_Density_Constant = 10 //Other_Info 水平密度常量
Vertical_Density_Constant   = 11 //Other_Info 垂直密度常量
City_Constant               = 12 //Other_Info 地市常量
OC_Sys_Name_Constant        = 13 //Other_Info 光缆系统名称常量
OC_Sys_Type_Constant        = 14 //Other_Info 光缆规格常量

FSB_Name_Constant = 1 //FSB_Info 分纤箱名字
FSB_Longitude_Constant = 2 //FSB_Info 分纤箱经度
FSB_Latitude_Constant = 3 //FSB_Info 分纤箱纬度


Source_File = "C:\Users\Administrator\Desktop\test.xls"
FSB_Topology_Sheet = "FSB_Topology"
OC_Topology_Sheet = "OC_Topology"
Info_Sheet = "Info"
Longitude_Sheet = "longitude"
Latitude_Sheet = "latitude"
OCS_Fiber_Count_Sheet = "OCS_Fiber_Count"

Control_Step0_Auto         = 0 //无人值守0,人工确认1
Control_Step1_Prepare_Data = 1 //读表
Control_Step2_Push_Box     = 0 //箱子
Control_Step3_Draw_SS      = 0 //引上
Control_Step4_Draw_OC      = 1 //光缆
Control_Step5_OC_Laying    = 1 //敷设

Call MainProgress()
Function MainProgress()
    Delay Delay_Unit * 20
    If Control_Step1_Prepare_Data = 1 Then
        Call Prepare_Data()
        If Control_Step0_Auto      = 1 Then
            MessageBox "请检查数据表,按确定后双手离开键盘鼠标"
        End If
    End If
    If Control_Step2_Push_Box     = 1 Then
        Call Push_Box()
        If Control_Step0_Auto      = 1 Then
            MessageBox "【推箱子】完成,请检查结果,按确定后双手离开键盘鼠标"
        End If
    End If
    If Control_Step3_Draw_SS      = 1 Then
        For Main_i = 1 To (Ubound(OCS_CC) - 1)
            Call Draw_SS(Main_i)
            If Main_i = 1 Then 
                Call SS_Set_New_Branch()
            End If
            If Main_i <> 1 Then 
                Call SS_Use_Old_Branch()
            End If
        Next
        If Control_Step0_Auto      = 1 Then
            MessageBox "【引上段】完成,请检查结果,按确定后双手离开键盘鼠标"
        End If
    End If
    If Control_Step4_Draw_OC      = 1 Then
        For Main_k = 1 To (Ubound(OCS_CC) - 1)
            Call Draw_OC(Main_k)
            If Main_k = 1 Then 
                Call OC_Set_New_Branch()
            End If
            If Main_k <> 1 Then 
                Call OC_Use_Old_Branch()
            End If
        Next
        If Control_Step0_Auto      = 1 Then
            MessageBox "【光缆段】完成,请检查结果,按确定后【双手离开键盘鼠标】"
        End If
    End If
    If Control_Step5_OC_Laying    = 1 Then
        Call OC_Laying()
        If Control_Step0_Auto      = 1 Then
            MessageBox "【敷设】完成,请检查结果,按确定后【双手离开键盘鼠标】"
        End If
    End If
End Function

Function SS_Set_New_Branch() //新建引上段输入信息
    Delay Delay_Unit * 5
    MoveTo 1070, 570 //新增选项
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1250, 570 //新增按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 800, 370
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    SayString Other_Info(SS_Name_Constant)
    Delay Delay_Unit * 5
    MoveTo 890, 465
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 890, 595
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 860, 595
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1070, 470
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1070, 490
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 920, 775
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1250, 635
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Do While True //下面判断新增管线是否完成,并弹出管线编辑界面//
        FindColor 330, 200, 1000, 700, "CC9966", MouseX2, MouseY2
        If MouseX2 > 0 And MouseY2 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    Delay Delay_Unit * 5
    MoveTo 480, 290
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 480, 325
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 358
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1180, 470
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 470
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 383
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 430
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 408
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 455
    Delay Delay_Unit * 5
    LeftClick 1 //产权单位结束
    Delay Delay_Unit * 5
    MoveTo 1150, 435
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(Project_Num_Constant))
    Delay Delay_Unit * 10
    Call After_SWI()
    Call SingleWordInput(Other_Info(Task_Name_Constant))
    Delay Delay_Unit * 10
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Porject_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_City_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_District_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Maintenance_Constant))
    Call After_SWI()
    Call After_SWI()
    Call SingleWordInput(Other_Info(District_Constant))
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveTo 1150, 660
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1150, 700
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    For Main_j = 1 To OCS_CC(Main_i)
        MoveTo 777, 367
        MoveR 0, 25 * (Main_j-1)
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveTo 1060, 610
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        SayString OC_Info(Main_i, Main_j, Lenth_Constant)
        Delay Delay_Unit * 5
    Next
    MoveTo 855, 285 //保存按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0) //等待保存,编辑窗口关闭
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    Do While True//判断弹出保存成功
        FindColor 400, 300, 1000, 700, "333333", MouseX1, MouseY1
        If MouseX1 > 0 And MouseY1 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    MoveTo MouseX1 + 80, MouseY1 + 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5   
End Function
Function SS_Use_Old_Branch() //利旧引上段输入信息
    Delay Delay_Unit * 5
    MoveTo 1130, 540
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call Plugin.Sys.SetCLB(Other_Info(SS_Name_Constant))
    KeyDown "Ctrl", 1
    Delay Delay_Unit * 5
    KeyPress "V", 1
    Delay Delay_Unit * 5
    KeyUp "Ctrl", 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0)
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    MoveTo 1200, 565
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1240, 635
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 480, 290
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 480, 325
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 358
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1180, 470
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 470
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 383
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 430
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1182, 408
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1160, 455
    Delay Delay_Unit * 5
    LeftClick 1 //产权单位结束
    Delay Delay_Unit * 5
    MoveTo 1150, 435
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(Project_Num_Constant))
    Delay Delay_Unit * 10
    Call After_SWI()
    Call SingleWordInput(Other_Info(Task_Name_Constant))
    Delay Delay_Unit * 10
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Porject_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_City_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_District_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Maintenance_Constant))
    Call After_SWI()
    Call After_SWI()
    Call SingleWordInput(Other_Info(District_Constant))
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveTo 1150, 660
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1150, 700
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    For Main_j = 1 To OCS_CC(Main_i)
        MoveTo 777, 367
        MoveR 0, 25 * (Main_j-1)
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveTo 1060, 610
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        SayString OC_Info(Main_i, Main_j, Lenth_Constant)
        Delay Delay_Unit * 5
    Next
    MoveTo 855, 285 //保存按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0) //等待保存,编辑窗口关闭
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    Do While True//判断弹出保存成功
        FindColor 400, 300, 1000, 700, "333333", MouseX1, MouseY1
        If MouseX1 > 0 And MouseY1 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    MoveTo MouseX1 + 80, MouseY1 + 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5   
End Function
Function Draw_OC(SPY_in_Func) //画光缆段
    Call Determine_Box(OC_Info(SPY_in_Func, 1, Start_Point_Constant))
    MoveTo 250, 232
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    KeyDown 17, 1
    KeyPress 65, 1
    KeyUp 17, 1
    Delay Delay_Unit * 5
    KeyPress "Delete", 1
    Delay Delay_Unit * 5
    SayString OC_Info(SPY_in_Func, 1, Start_Point_Constant)
    Delay Delay_Unit * 5
    MoveR 0, 35
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0)
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    MoveR 0, 60
    Delay Delay_Unit * 5
    LeftDoubleClick 1
    Delay Delay_Unit * 10
    MoveTo 345, 175 //清除按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 345, 380 //光缆段新增按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 345, 135 //开始录入按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1116, 548 //录入起始点
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    For i = 1 To OCS_CC(SPY_in_Func)
        R_Move_X = Draw_Vector(SPY_in_Func, i, 0)
        R_Move_Y = Draw_Vector(SPY_in_Func, i, 1)
        MoveR R_Move_X * 75, R_Move_Y * -75
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
    Next
    MoveTo 520, 135 //确认编辑按钮
    Delay Delay_Unit * 5
    LeftClick 1
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
End Function
Function OC_Set_New_Branch() //新建光缆段输入信息
    Delay Delay_Unit * 5
    MoveTo 1070, 570 //新增选项
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1250, 570 //新增按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 800, 370 //光缆系统中文名称
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    SayString Other_Info(OC_Sys_Name_Constant)
    Delay Delay_Unit * 5
    MoveTo 890, 470 //光缆系统光缆规格
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    SayString Other_Info(OC_Sys_Type_Constant)
    Delay Delay_Unit * 5
    MoveTo 1180, 420 //系统级别
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1180, 520
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1150, 545 //驻地网
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1180, 490
    Delay Delay_Unit * 5
    LeftClick 1 //地市
    Delay Delay_Unit * 5
    KeyPress "Down", 1
    Delay Delay_Unit * 5
    KeyPress "Enter", 1
    Delay Delay_Unit * 5
    MoveTo 895, 515 //区县
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(District_Constant))
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveTo 920, 775
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1245, 600
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Do While True //下面判断新增光缆是否完成,并弹出光缆编辑界面//
        FindColor 330, 200, 1000, 700, "CC9966", MouseX3, MouseY3
        If MouseX3 > 0 And MouseY3 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    Mouse_searching = GetCursorShape(0) //二重判断
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    MoveTo MouseX3, MouseY3 //基准点698,251
    Delay Delay_Unit * 5
    MoveR 105, 35 //选择
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 0, 40 //全选
    Delay Delay_Unit * 5
    LeftClick 1 
    Delay Delay_Unit * 5
    MoveR 695, 35 //光缆段级别
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 40 
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, 30 //纤芯规格
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, -15 //产权性质
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 50
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, -20 //产权单位
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 40
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 0, -15
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(Project_Num_Constant))
    Delay Delay_Unit * 20
    Call After_SWI()
    Call SingleWordInput(Other_Info(Task_Name_Constant))
    Delay Delay_Unit * 20
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Porject_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_City_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_District_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Maintenance_Constant))
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveR 70, 145
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -45, -10 //地市
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call After_SWI()
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(District_Constant)) //区县
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveTo MouseX3, MouseY3 //基准点698,251,准备输入长度,芯数,类型
    Delay Delay_Unit * 5
    MoveR 444, 120 //第一行
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 400, 230 //滚动条
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    For Main_l = 1 To OCS_CC(Main_k)
        MoveTo MouseX3, MouseY3 //基准点698,251,准备输入长度,芯数,类型
        Delay Delay_Unit * 5
        MoveR 444, 120 //第一行
        Delay Delay_Unit * 5
        MoveR 0, 25 * (Main_l-1) //第N行偏移
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveR 236, (-12 + (Main_l -1) * -25) //维护长度
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5        
        SayString OC_Info(Main_k, Main_l, Lenth_Constant)
        Delay Delay_Unit * 5
        MoveR 0, 24 //纤芯数
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5        
        Select Case OC_Info(Main_k, Main_l, Fiber_Count_Constant)
        Case 6
            KeyPress "6", 1
            Delay Delay_Unit * 5
            KeyPress "Enter", 1     
            Delay Delay_Unit * 5
        Case 12
            KeyPress "1", 1
            Delay Delay_Unit * 5
            KeyPress "Enter", 1
            Delay Delay_Unit * 5
        Case 24
            KeyPress "2", 1
            Delay Delay_Unit * 3
            KeyPress "2", 1
            Delay Delay_Unit * 3
            KeyPress "2", 1
            Delay Delay_Unit * 5
            KeyPress "Enter", 1     
            Delay Delay_Unit * 5
        End Select
        Delay Delay_Unit * 5
        MoveR 0, 50//光缆段型号
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        SayString OC_Info(Main_k, Main_l, Type_Constant)
        Delay Delay_Unit * 5
    Next
    MoveTo MouseX3, MouseY3
    Delay Delay_Unit * 5
    MoveR 370, 37 //保存按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0) //等待保存,编辑窗口关闭
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    Do While True//判断弹出保存成功
        FindColor 400, 300, 1200, 700, "333333", MouseX3, MouseY3
        If MouseX3 > 0 And MouseY3 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    MoveTo MouseX3 + 80, MouseY3 + 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5   
End Function
Function OC_Use_Old_Branch() //利旧光缆段输入信息
    Delay Delay_Unit * 5
    MoveTo 1130, 545 //利旧选项
    Delay Delay_Unit * 5
    LeftClick 1
    Call Plugin.Sys.SetCLB(Other_Info(OC_Sys_Name_Constant))
    KeyDown "Ctrl", 1
    Delay Delay_Unit * 5
    KeyPress "V", 1
    Delay Delay_Unit * 5
    KeyUp "Ctrl", 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0)
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    MoveTo 1200, 565
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 1245, 600 //确定按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Do While True //下面判断新增光缆是否完成,并弹出光缆编辑界面//
        FindColor 330, 200, 1000, 700, "CC9966", MouseX3, MouseY3
        If MouseX3 > 0 And MouseY3 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    Mouse_searching = GetCursorShape(0) //二重判断
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    MoveTo MouseX3, MouseY3 //基准点698,251
    Delay Delay_Unit * 5
    MoveR 105, 35 //选择
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 0, 40 //全选
    Delay Delay_Unit * 5
    LeftClick 1 
    Delay Delay_Unit * 5
    MoveR 695, 35 //光缆段级别
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 40 
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, 30 //纤芯规格
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, -15 //产权性质
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 50
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 25, -20 //产权单位
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -25, 40
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Delay Delay_Unit * 5
    MoveR 0, -15
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(Project_Num_Constant))
    Delay Delay_Unit * 20
    Call After_SWI()
    Call SingleWordInput(Other_Info(Task_Name_Constant))
    Delay Delay_Unit * 20
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Porject_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_City_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_District_Constant))
    Call After_SWI()
    Call SingleWordInput(Other_Info(DQS_Maintenance_Constant))
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveR 70, 145
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR -45, -10 //地市
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Call After_SWI()
    Delay Delay_Unit * 5
    Call SingleWordInput(Other_Info(District_Constant)) //区县
    Call After_SWI()
    Delay Delay_Unit * 5
    MoveTo MouseX3, MouseY3 //基准点698,251,准备输入长度,芯数,类型
    Delay Delay_Unit * 5
    MoveR 444, 120 //第一行
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveR 400, 230 //滚动条
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    For Main_l = 1 To OCS_CC(Main_k)
        MoveTo MouseX3, MouseY3 //基准点698,251,准备输入长度,芯数,类型
        Delay Delay_Unit * 5
        MoveR 444, 120 //第一行
        Delay Delay_Unit * 5
        MoveR 0, 25 * (Main_l-1) //第N行偏移
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveR 236, (-12 + (Main_l -1) * -25) //维护长度
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5        
        SayString OC_Info(Main_k, Main_l, Lenth_Constant)
        Delay Delay_Unit * 5
        MoveR 0, 24 //纤芯数
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5        
        Select Case OC_Info(Main_k, Main_l, Fiber_Count_Constant)
        Case 6
            KeyPress "6", 1
            Delay Delay_Unit * 5
            KeyPress "Enter", 1     
            Delay Delay_Unit * 5                   
        Case 12
            KeyPress "1", 1
            Delay Delay_Unit * 5
            KeyPress "Enter", 1     
            Delay Delay_Unit * 5    
        Case 24
            KeyPress "2", 1
            Delay Delay_Unit * 3
            KeyPress "2", 1
            Delay Delay_Unit * 3
            KeyPress "2", 1                     			
            Delay Delay_Unit * 5
            KeyPress "Enter", 1     
            Delay Delay_Unit * 5
        End Select
        Delay Delay_Unit * 5
        MoveR 0, 50//光缆段型号
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        SayString OC_Info(Main_k, Main_l, Type_Constant)
        Delay Delay_Unit * 5
    Next
    MoveTo MouseX3, MouseY3
    Delay Delay_Unit * 5
    MoveR 370, 37 //保存按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0) //等待保存,编辑窗口关闭
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    Do While True//判断弹出保存成功
        FindColor 400, 300, 1200, 700, "333333", MouseX3, MouseY3
        If MouseX3 > 0 And MouseY3 > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    MoveTo MouseX3 + 80, MouseY3 + 65
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5   
End Function
Function OC_Laying() //敷设
    Delay Delay_Unit * 5
    MoveTo 345, 175 //清除按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 500, 105 //线设施按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 500, 135 //光缆系统管理按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MouseX = 0 //判断光缆系统管理窗口//
    MouseY = 0
    Do While True
        FindColor 400, 150, 1600, 900, "CC9966", MouseX, MouseY
        If MouseX > 0 And MouseY > 0 Then 
            Exit Do
        End If		
        Delay Delay_Unit * 5
    Loop
    Mouse_searching = GetCursorShape(0) //二重判断
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    Delay Delay_Unit * 5
    MoveTo 500, 220 //中文名称(光缆系统)
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    SayString Other_Info(OC_Sys_Name_Constant)
    Delay Delay_Unit * 5
    KeyPress "Enter", 1 //搜索
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0)
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    MoveTo 500, 330 //搜索结果(应该唯一)
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    MoveTo 430, 850 //路由管理按钮
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
    Mouse_searching = GetCursorShape(0)
    While Mouse_searching=0
        Delay Delay_Unit * 5
        Mouse_searching = GetCursorShape(0)
    Wend
    For CL1_in_PD = 1 to Ubound(OCS_CC) //计算敷设段落总数
        OC_Laying_CC = OC_Laying_CC + OCS_CC(CL1_in_PD)
    Next
    For CL1_in_PD = 1 to OC_Laying_CC
        MoveTo 500, 215 //第一行
        MoveR 0, (CL1_in_PD - 1) * 25
        Delay Delay_Unit * 5
        LeftClick 1 //复制光缆段名称
        Delay Delay_Unit * 10
        LeftDoubleClick 1
        Delay Delay_Unit * 10
        KeyPress "Enter", 1
        Delay Delay_Unit * 10
        //处理编码
        SaveMousePos
        Delay Delay_Unit * 5
		MoveTo 500, 45
		Delay Delay_Unit * 5
		LeftClick 1
        Delay Delay_Unit * 5
        KeyDown "Ctrl", 1
		Delay Delay_Unit * 5
		KeyDown "V", 1
		Delay Delay_Unit * 5
		KeyUp "Ctrl", 1
		Delay Delay_Unit * 5
		KeyDown "Ctrl", 1
		Delay Delay_Unit * 5
		KeyPress "A", 1
		Delay Delay_Unit * 5
		KeyUp "Ctrl", 1
		Delay Delay_Unit * 5
        KeyDown "Ctrl", 1
		Delay Delay_Unit * 5
		KeyDown "C", 1
		Delay Delay_Unit * 5
		KeyUp "Ctrl", 1		
		Delay Delay_Unit * 5
		RestoreMousePos
		//处理编码结束
        LeftClick 1
        Delay Delay_Unit * 5
        MoveTo 500, 850 //光缆段敷设按钮
        Delay Delay_Unit * 5
        LeftClick 1
        MouseX = 0 //前置判断光缆段未敷设确认//
        MouseY = 0
        Do While True
            FindColor 370, 250, 1600, 900, "A1886E", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MouseX = 0 //判断光缆段未敷设确认//
        MouseY = 0
        Do While True
            FindColor 370, 200, 1600, 900, "333333", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MoveTo MouseX, MouseY
        Delay Delay_Unit * 5
        MoveR 110, 65
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MouseX = 0 //判断敷设窗口//
        MouseY = 0
        Do While True
            FindColor 370, 250, 1600, 900, "CC9966", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MoveTo MouseX, MouseY
        Delay Delay_Unit * 5
        //判断光缆段名称在敷设窗口占用几行，每行偏移14px
        OCS_Name = Plugin.Sys.GetCLB()
        Chinese_Char_Num = 0
        Other_Char_Num = 0
        For CL2_in_OC_L = 1 To len(OCS_Name)
            Single_String=mid(OCS_Name,CL2_in_OC_L,1)
            Check_Result=Plugin.lxj_Plug.lxj_CheckChar(Single_String)
            If Check_Result = 1 Then
                Chinese_Char_Num = Chinese_Char_Num + 1
                TracePrint Chinese_Char_Num
            ElseIf Check_Result <> 1 Then 
                Other_Char_Num = Other_Char_Num + 1
                TracePrint Other_Char_Num
            End If
        Next
        OCS_Name_Lenth = Chinese_Char_Num * 2 + Other_Char_Num
        TracePrint OCS_Name_Lenth
        If OCS_Name_Lenth <= 50 Then                                 //一行
            Y_Offset = 0
        ElseIf OCS_Name_Lenth >50  And OCS_Name_Lenth <= 100 Then    //两行
            Y_Offset = 1
        ElseIf OCS_Name_Lenth >100 And OCS_Name_Lenth <= 150 Then    //三行
            Y_Offset = 2
        ElseIf OCS_Name_Lenth >150 And OCS_Name_Lenth <= 200 Then    //四行
            Y_Offset = 3
        End If
        TracePrint Y_Offset
        MoveR 60, (67 + Y_Offset * 14)//下拉列表
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveR 0, 40 //快捷模式
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveR 300, -43 //框选按钮
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveTo 360, 140 //左上起点
        Delay Delay_Unit * 5
        leftDown 1
        Delay Delay_Unit * 5
        MoveTo 1700, 950 //右下终点
        Delay Delay_Unit * 5
        leftUp 1
        Delay Delay_Unit * 5
        While Mouse_searching=0
            Delay Delay_Unit * 5
            Mouse_searching = GetCursorShape(0)
        Wend
        Delay Delay_Unit * 5
        MoveTo MouseX, MouseY
        Delay Delay_Unit * 5
        MoveR 100, 70 //点击通路
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MoveR 100, 315 //确认
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        LeftClick 1 //确认敷设
        Delay Delay_Unit * 5
        While Mouse_searching=0
            Delay Delay_Unit * 5
            Mouse_searching = GetCursorShape(0)
        Wend
        Delay Delay_Unit * 5
        MouseX = 0 //前置判断保存敷设成功//
        MouseY = 0
        Do While True
            FindColor 370, 250, 1600, 900, "A1886E", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MouseX = 0 //判断保存敷设成功//
        MouseY = 0
        Do While True
            FindColor 875, 510, 1040, 600, "333333", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MoveTo MouseX, MouseY
        Delay Delay_Unit * 5
        MoveR 75, 65
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
        MouseX = 0 //重新判断敷设窗口//
        MouseY = 0
        Do While True
            FindColor 370, 250, 1600, 900, "CC9966", MouseX, MouseY
            If MouseX > 0 And MouseY > 0 Then 
                Exit Do
            End If		
            Delay Delay_Unit * 5
        Loop
        MoveTo MouseX, MouseY
        Delay Delay_Unit * 5
        MoveR 407, 5
        Delay Delay_Unit * 5
        LeftClick 1
        Delay Delay_Unit * 5
    Next
    Delay Delay_Unit * 5
    MoveTo 1905, 155
    Delay Delay_Unit * 5
    LeftClick 1
    Delay Delay_Unit * 5
End Function
Function Prepare_Data() //数据准备函数
    Call Plugin.Office.OpenXls(Source_File)

    Redim OCS_CC(100) //光缆段循环计数
    For CL1_in_PD = 1 To C_End
        For CL2_in_PD =1 To C_End
            Temp_Data_in_PD = Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, CL2_in_PD + 1)
            If Temp_Data_in_PD = ""
                Exit For
            End If
        Next
        Temp_Data_in_PD = Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD,2)
        If Temp_Data_in_PD = ""
            Exit For
        End If
        OCS_CC(CL1_in_PD) = CL2_in_PD - 1
    Next
    Redim Preserve OCS_CC(CL1_in_PD)//Redim上限,此后可用
    For CL1_in_PD = 1 To (Ubound(OCS_CC) - 1)//加载芯数
        For CL2_in_PD = 1 To OCS_CC(CL1_in_PD)
            Temp_Data_in_PD = Plugin.Office.ReadXls(OCS_Fiber_Count_Sheet, CL1_in_PD, CL2_in_PD + 1)
            If Temp_Data_in_PD = "" Then 
                Exit For
            End If
            OC_Info(CL1_in_PD,CL2_in_PD,Fiber_Count_Constant) = CInt(Temp_Data_in_PD)
        Next
        Temp_Data_in_PD = Plugin.Office.ReadXls(OCS_Fiber_Count_Sheet, CL1_in_PD, 2)
        If Temp_Data_in_PD = "" Then 
            Exit For
        End If
    Next
    Dim OCS_Lenth_X_Y //加载长度
    Horizontal_Density = Plugin.Office.ReadXls(Info_Sheet, Horizontal_Density_Constant, Other_Info_Data_Column)
    Vertical_Density = Plugin.Office.ReadXls(Info_Sheet, Vertical_Density_Constant, Other_Info_Data_Column)    
    Base_Latitude = CDbl(Plugin.Office.ReadXls(Latitude_Sheet, 1, 1))    
    Horizontal_Metre = 111.11 * 1000 * cos(Base_Latitude * 3.1415926535 / 180) * Vertical_Density
    Vertical_Metre = 111.11 * 1000 * Horizontal_Density   
    For CL1_in_PD = 1 To (Ubound(OCS_CC) - 1)
        For CL2_in_PD = 1 To OCS_CC(CL1_in_PD)
            OCS_Lenth_temp = Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, (CL2_in_PD + 1))
            OCS_Lenth_X_Y = split(OCS_Lenth_temp, "&")
            If OCS_Lenth_X_Y(0) = 0 Then //竖线
                OC_Info(CL1_in_PD, CL2_in_PD, Lenth_Constant) = round(abs(Vertical_Metre * OCS_Lenth_X_Y(1)))
            End If
            If OCS_Lenth_X_Y(1) = 0 Then //横线
                OC_Info(CL1_in_PD, CL2_in_PD, Lenth_Constant) = round(abs(Horizontal_Metre * OCS_Lenth_X_Y(0)))
            End If
            If (OCS_Lenth_X_Y(0) <> 0 and OCS_Lenth_X_Y(1) <> 0) Then //斜线
                OC_Info(CL1_in_PD, CL2_in_PD, Lenth_Constant) = round(abs(sqr((OCS_Lenth_X_Y(0) * Vertical_Metre) ^ 2 + (OCS_Lenth_X_Y(1) * Horizontal_Metre) ^ 2)))
            End If
        Next
    Next
    For CL1_in_PD = 1 To (Ubound(OCS_CC) - 1)//加载光缆段型号
        For CL2_in_PD = 1 To OCS_CC(CL1_in_PD)
            OC_Info(CL1_in_PD, CL2_in_PD, Type_Constant) = "GYTA-" & CStr(OC_Info(CL1_in_PD,CL2_in_PD,Fiber_Count_Constant))
        Next
    Next
    For CL1_in_PD = 1 To Ubound(Other_Info) //加载其他信息
        Other_Info(CL1_in_PD) = Plugin.Office.ReadXls(Info_Sheet, CL1_in_PD, Other_Info_Data_Column)
    Next
    For CL1_in_PD = 1 To (Ubound(FSB_CC) - 1) //加载分纤箱信息
        For CL2_in_PD = 1 To FSB_CC(CL1_in_PD)
            FSB_Info(CL1_in_PD, CL2_in_PD, FSB_Name_Constant) = Plugin.Office.ReadXls(FSB_Topology_Sheet, CL1_in_PD, CL2_in_PD)
            FSB_Info(CL1_in_PD, CL2_in_PD, FSB_Longitude_Constant) = Plugin.Office.ReadXls(Longitude_Sheet, CL1_in_PD, CL2_in_PD)
            FSB_Info(CL1_in_PD, CL2_in_PD, FSB_Latitude_Constant) = Plugin.Office.ReadXls(Latitude_Sheet, CL1_in_PD, CL2_in_PD)
        Next
    Next
    Dim temp_a_PD //加载光缆向量
    Dim temp_b_PD
    For CL1_in_PD = 1 To (Ubound(OCS_CC) - 1)
        For CL2_in_PD = 1 To OCS_CC(CL1_in_PD)
            OC_Info(CL1_in_PD, CL2_in_PD, Vector_Constant) = Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, (CL2_in_PD + 1))//光缆段向量(未处理)
        Next
        OC_Info(CL1_in_PD, 1, Start_Point_Constant) = Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, 1)//起始FSB坐标
        temp_a_PD = split(Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, 1), "&")
        OC_Info(CL1_in_PD, 1, Start_Point_Constant) = Plugin.Office.ReadXls(FSB_Topology_Sheet, CInt(temp_a_PD(1)), CInt(temp_a_PD(0)))//起始FSB名称
        For CL2_in_PD = 1 to OCS_CC(CL1_in_PD) //光缆段向量存入二维数组
            temp_b_PD = split(Plugin.Office.ReadXls(OC_Topology_Sheet, CL1_in_PD, (CL2_in_PD + 1)), "&") 
            Draw_Vector(CL1_in_PD, CL2_in_PD, 0) = temp_b_PD(0)
            Draw_Vector(CL1_in_PD, CL2_in_PD, 1) = temp_b_PD(1)
        Next
    Next
    Call Plugin.Office.CloseXls()
End Function
