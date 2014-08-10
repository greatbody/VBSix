Attribute VB_Name = "apiModule"
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Const GW_OWNER = 4
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    '窗口的标题
    Dim Title As String
    '可见性，如果窗口不可见，则返回0
    Dim VisiableState As Long
    '窗口的层级，如果窗口有父窗口，则返回值非零，这里我们只需要顶级窗口的，所以当窗口是0（即窗口自身就是父窗口）才显示窗口名
    Dim LevelState As Long
    '给它提供初始值
    Title = String(80, 0)
    Call GetWindowText(hwnd, Title, 80)
    Title = Left(Title, InStr(Title, Chr$(0)) - 1)

    LevelState = GetWindow(hwnd, GW_OWNER)
    VisiableState = IsWindowVisible(hwnd)
    If Len(Title) > 0 And VisiableState <> 0 And LevelState = 0 Then
        Form1.List1.AddItem Title
    End If
    EnumWindowsProc = True
End Function
