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

'ע�⣺1������VB6 IDE����ԭ�򣬵��Կ���̨����ʱIDE����Ӧ���������������ֶ���
'         �տ���̨���ڣ�����ֱ���˳�IDE������ʱ�������Ĵ��롣
'
'      2����[��������>����>�����������]ѡ�������� VB_DEBUG=1 �ɰ������Գ���
'         ��ֹ���н�����ֱ���˳�����̨�����������н������ʽ����ʱ������Ϊ0��
'
'      3������ PowerVB Console Application Add-In �������[��ӳ���]�˵��У�ѡ
'         �� Complie As Windows Console Application���ɱ���Ϊ�����Ŀ���̨����
'
'      4������̨����һ�㲻ʹ�öԻ����Ƽ���ѡ[��������>ͨ��>���û�����]ѡ�


'------------------------------------------------------------------------------
'       ���б���
'------------------------------------------------------------------------------
'@summary   ����̨����
Public Console As CConsole



'------------------------------------------------------------------------------
'@summary   �������
'------------------------------------------------------------------------------
Public Sub Main()
    Set Console = New CConsole
    '--------------------------------------------------------------------------
    Dim intNum As Integer, intUnit As Integer
    'TODO: �ڴ˱༭����
    Console.WriteLine "Hello, Welcome to the world of Visual Basic"
    'intNum = Int(Console.ReadLine("��������ֵ:"))
    'intUnit = Int(Console.ReadLine("�����뵥λ��ֵ:"))
    'Console.WriteLine "����ֵ " & intNum & " �� " & CalcMiss(intNum, intUnit) & " Ϊ " & intUnit & " �ı���"
    '����֤�ַ�������������
        'Console.WriteLine("abcdef roll left 7:=>" & ShiftStr("abcdef", 7, RollLeft, True))
        'Console.WriteLine("abcdef roll left 3:=>" & ShiftStr("abcdef", 3, RollLeft, True))
        'Console.WriteLine("abcdef move left 7:=>" & ShiftStr("abcdef", 7, RollLeft, False))
        'Console.WriteLine("abcdef move left 3:=>" & ShiftStr("abcdef", 3, RollLeft, False))

        'Console.WriteLine("abcdef roll right 7:=>" & ShiftStr("abcdef", 7, RollRight, True))
        'Console.WriteLine("abcdef roll right 3:=>" & ShiftStr("abcdef", 3, RollRight, True))
        'Console.WriteLine("abcdef move right 7:=>" & ShiftStr("abcdef", 7, RollRight, False))
        'Console.WriteLine("abcdef move right 3:=>" & ShiftStr("abcdef", 3, RollRight, False))
    '����֤�ֽ�ת���������֡�
    'Console.WriteLine "&H13:" & ByteToString(&HAF)
    '--------------------------------------------------------------------------
    #If VB_DEBUG Then
        Console.WriteText "Press any key to continue..."
        Console.ReadKey
    #End If
    Console.WriteLine
End Sub
