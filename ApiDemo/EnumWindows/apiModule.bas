Attribute VB_Name = "apiModule"

Public Declare Function GetWindowText _
               Lib "user32" _
               Alias "GetWindowTextA" (ByVal hwnd As Long, _
                                       ByVal lpString As String, _
                                       ByVal cch As Long) As Long

Public Declare Function EnumWindows _
               Lib "user32" (ByVal lpEnumFunc As Long, _
                             ByVal lParam As Long) As Long

Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Public Const GW_OWNER = 4

Public Declare Function GetWindow _
               Lib "user32" (ByVal hwnd As Long, _
                             ByVal wCmd As Long) As Long

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean

    '���ڵı���
    Dim Title         As String

    '�ɼ��ԣ�������ڲ��ɼ����򷵻�0
    Dim VisiableState As Long

    '���ڵĲ㼶����������и����ڣ��򷵻�ֵ���㣬��������ֻ��Ҫ�������ڵģ����Ե�������0��������������Ǹ����ڣ�����ʾ������
    Dim LevelState    As Long

    '�����ṩ��ʼֵ
    Title = String(80, 0)
    Call GetWindowText(hwnd, Title, 80)
    Title = Left(Title, InStr(Title, Chr$(0)) - 1)

    LevelState = GetWindow(hwnd, GW_OWNER)
    VisiableState = IsWindowVisible(hwnd)

    If Len(Title) > 0 Then

        Form1.List1.AddItem Title
    End If

    EnumWindowsProc = True

End Function
