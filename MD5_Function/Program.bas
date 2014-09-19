Attribute VB_Name = "Program"
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary      Program
'@@require
'@@reference
'@@license
'@@author
'@@create
'@@modify
'//////////////////////////////////////////////////////////////////////////////

'注意：1、由于VB6 IDE固有原因，调试控制台程序时IDE无响应属正常现象，请勿手动关
'         闭控制台窗口！否则将直接退出IDE，请随时保存您的代码。
'
'      2、在[工程属性>生成>条件编译参数]选项中设置 VB_DEBUG=1 可帮助调试程序，
'         防止运行结束后直接退出控制台而看不到运行结果，正式编译时，可设为0。
'
'      3、启用 PowerVB Console Application Add-In 插件，在[外接程序]菜单中，选
'         择 Complie As Windows Console Application，可编译为真正的控制台程序。
'
'      4、控制台程序一般不使用对话框，推荐勾选[工程属性>通用>无用户界面]选项。


'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------
'@summary   控制台对象
Public Console As CConsole



'------------------------------------------------------------------------------
'@summary   程序入口
'------------------------------------------------------------------------------
Public Sub Main()
    Set Console = New CConsole
    '--------------------------------------------------------------------------
    Dim intNum As Integer, intUnit As Integer
    'TODO: 在此编辑代码
    Console.WriteLine "Hello, Welcome to the world of Visual Basic"
    'intNum = Int(Console.ReadLine("请输入数值:"))
    'intUnit = Int(Console.ReadLine("请输入单位数值:"))
    'Console.WriteLine "比数值 " & intNum & " 大 " & CalcMiss(intNum, intUnit) & " 为 " & intUnit & " 的倍数"
    '【验证字符串滚动函数】
        'Console.WriteLine("abcdef roll left 7:=>" & ShiftStr("abcdef", 7, RollLeft, True))
        'Console.WriteLine("abcdef roll left 3:=>" & ShiftStr("abcdef", 3, RollLeft, True))
        'Console.WriteLine("abcdef move left 7:=>" & ShiftStr("abcdef", 7, RollLeft, False))
        'Console.WriteLine("abcdef move left 3:=>" & ShiftStr("abcdef", 3, RollLeft, False))

        'Console.WriteLine("abcdef roll right 7:=>" & ShiftStr("abcdef", 7, RollRight, True))
        'Console.WriteLine("abcdef roll right 3:=>" & ShiftStr("abcdef", 3, RollRight, True))
        'Console.WriteLine("abcdef move right 7:=>" & ShiftStr("abcdef", 7, RollRight, False))
        'Console.WriteLine("abcdef move right 3:=>" & ShiftStr("abcdef", 3, RollRight, False))
    '【验证字节转二进制文字】
    'Console.WriteLine "&H13:" & ByteToString(&HAF)
    '--------------------------------------------------------------------------
    #If VB_DEBUG Then
        Console.WriteText "Press any key to continue..."
        Console.ReadKey
    #End If
    Console.WriteLine
End Sub
