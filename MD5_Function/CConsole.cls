VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//////////////////////////////////////////////////////////////////////////////
'@@summary      控制台
'@@require
'@@reference
'@@license      BSD
'@@author
'@@create       2013-02-16
'@@modify       2013-05-30
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       接口继承
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有常量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有数据类型
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       公有API
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       事件声明
'------------------------------------------------------------------------------


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有声明
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       私有常量
'------------------------------------------------------------------------------
'API
Private Const STD_INPUT_HANDLE As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const STD_ERROR_HANDLE As Long = -12
Private Const CONSOLE_FULLSCREEN_MODE As Long = 1
Private Const CONSOLE_WINDOWED_MODE As Long = 2
Private Const CONSOLE_FULLSCREEN  As Long = 1
Private Const CONSOLE_FULLSCREEN_HARDWARE As Long = 2
Private Const ENABLE_ECHO_INPUT As Long = &H4
Private Const ENABLE_INSERT_MODE As Long = &H20
Private Const ENABLE_LINE_INPUT As Long = &H2
Private Const ENABLE_MOUSE_INPUT As Long = &H10
Private Const ENABLE_PROCESSED_INPUT As Long = &H1
Private Const ENABLE_QUICK_EDIT_MODE As Long = &H40
Private Const ENABLE_WINDOW_INPUT As Long = &H8
Private Const ENABLE_PROCESSED_OUTPUT As Long = &H1
Private Const ENABLE_WRAP_AT_EOL_OUTPUT As Long = &H2
Private Const COMMON_LVB_REVERSE_VIDEO As Integer = &H4000
Private Const KEY_EVENT As Integer = 1
'User
Private Const CONSOLE_BUFFER_MAXSIZE As Long = 8191 '缓冲区最大字符数


'------------------------------------------------------------------------------
'       私有数据类型
'------------------------------------------------------------------------------
Private Type COORD
    x As Integer
    y As Integer
End Type

Private Type SMALL_RECT
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize As COORD
    dwCursorPosition As COORD
    wAttributes As Integer
    srWindow As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Private Type CONSOLE_CURSOR_INFO
    dwSize As Long
    bVisible As Boolean
End Type

Private Type KEY_EVENT_RECORD
    bKeyDown As Long
    wRepeatCount As Integer
    wVirtualKeyCode As Integer
    wVirtualScanCode As Integer
    uChar As Integer
    dwControlKeyState As Long
End Type

Private Type INPUT_RECORD
    EventType As Integer
    EventContent As KEY_EVENT_RECORD
End Type


'------------------------------------------------------------------------------
'       私有变量
'------------------------------------------------------------------------------


'------------------------------------------------------------------------------
'       属性变量
'------------------------------------------------------------------------------
Private m_StdOut As Long
Private m_StdIn As Long
Private m_StdErr As Long
Private m_IsReady As Boolean
Private m_DefaultForeColor As Byte
Private m_DefaultBackColor As Byte


'------------------------------------------------------------------------------
'       私有API
'------------------------------------------------------------------------------
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function AllocConsole Lib "kernel32" () As Long
Private Declare Function FreeConsole Lib "kernel32" () As Long
Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function SetStdHandle Lib "kernel32" (ByVal nStdHandle As Long, ByVal hHandle As Long) As Long
Private Declare Function WriteConsole Lib "kernel32" Alias "WriteConsoleA" (ByVal hConsoleOutput As Long, ByVal lpBuffer As String, ByVal nNumberOfCharsToWrite As Long, lpNumberOfCharsWritten As Long, lpReserved As Any) As Long
Private Declare Function ReadConsole Lib "kernel32" Alias "ReadConsoleA" (ByVal hConsoleInput As Long, ByVal lpBuffer As Any, ByVal nNumberOfCharsToRead As Long, lpNumberOfCharsRead As Long, lpReserved As Any) As Long
Private Declare Function GetConsoleWindow Lib "kernel32" () As Long
Private Declare Function GetConsoleCP Lib "kernel32" () As Long
Private Declare Function SetConsoleCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Private Declare Function GetConsoleOutputCP Lib "kernel32" () As Long
Private Declare Function SetConsoleOutputCP Lib "kernel32" (ByVal wCodePageID As Long) As Long
Private Declare Function GetConsoleTitle Lib "kernel32" Alias "GetConsoleTitleA" (ByVal lpConsoleTitle As Any, ByVal nSize As Long) As Long
Private Declare Function SetConsoleTitle Lib "kernel32" Alias "SetConsoleTitleA" (ByVal lpConsoleTitle As String) As Long
Private Declare Function GetConsoleDisplayMode Lib "kernel32" (lpModeFlags As Long) As Long
Private Declare Function SetConsoleDisplayMode Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwFlags As Long, lpNewScreenBufferDimensions As COORD) As Long
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long
Private Declare Function GetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorInfo Lib "kernel32" (ByVal hConsoleOutput As Long, lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
Private Declare Function SetConsoleCursorPosition Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal dwCursorPosition As Long) As Long
Private Declare Function SetConsoleWindowInfo Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal bAbsolute As Long, lpConsoleWindow As SMALL_RECT) As Long
Private Declare Function GetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, lpMode As Long) As Long
Private Declare Function SetConsoleMode Lib "kernel32" (ByVal hConsoleHandle As Long, ByVal lpMode As Long) As Long
Private Declare Function SetConsoleTextAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long
Private Declare Function FillConsoleOutputCharacter Lib "kernel32" Alias "FillConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal cCharacter As Byte, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfCharsWritten As Long) As Long
Private Declare Function FillConsoleOutputAttribute Lib "kernel32" (ByVal hConsoleOutput As Long, ByVal wAttribute As Integer, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfCharsWritten As Long) As Long
Private Declare Function ReadConsoleOutputCharacter Lib "kernel32.dll" Alias "ReadConsoleOutputCharacterA" (ByVal hConsoleOutput As Long, ByVal lpCharacter As Any, ByVal nLength As Long, ByVal dwReadCoord As Long, lpNumberOfCharsRead As Long) As Long
Private Declare Function ReadConsoleOutputAttribute Lib "kernel32.dll" (ByVal hConsoleOutput As Long, lpAttribute As Any, ByVal nLength As Long, ByVal dwReadCoord As Long, lpNumberOfAttrsRead As Long) As Long
Private Declare Function WriteConsoleOutputAttribute Lib "kernel32.dll" (ByVal hConsoleOutput As Long, lpAttribute As Any, ByVal nLength As Long, ByVal dwWriteCoord As Long, lpNumberOfAttrsWritten As Long) As Long
Private Declare Function ReadConsoleInput Lib "kernel32.dll" Alias "ReadConsoleInputA" (ByVal hConsoleInput As Long, lpBuffer As INPUT_RECORD, ByVal nLength As Long, lpNumberOfEventsRead As Long) As Long
Private Declare Function FlushConsoleInputBuffer Lib "kernel32.dll" (ByVal hConsoleInput As Long) As Long


'//////////////////////////////////////////////////////////////////////////////
'//
'//      类
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'       初始化
'------------------------------------------------------------------------------
Private Sub Class_Initialize()
    AllocConsole
    m_StdOut = GetStdHandle(STD_OUTPUT_HANDLE)
    m_StdIn = GetStdHandle(STD_INPUT_HANDLE)
    m_StdErr = GetStdHandle(STD_ERROR_HANDLE)
    m_DefaultForeColor = ForeColor
    m_DefaultBackColor = BackColor
End Sub


'------------------------------------------------------------------------------
'       销毁
'------------------------------------------------------------------------------
Private Sub Class_Terminate()
    ForeColor = m_DefaultForeColor
    BackColor = m_DefaultBackColor
    FreeConsole
End Sub


'//////////////////////////////////////////////////////////////////////////////
'//
'//      事件处理
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      私有方法
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'@summary   剪切矩形
'------------------------------------------------------------------------------
'@param     被剪切的X坐标
'@param     被剪切的Y坐标
'@param     被剪切的宽度
'@param     被剪切的高度
'@param     最大宽度
'@param     最大高度
Private Sub ClipRect( _
    pX As Integer, _
    pY As Integer, _
    pW As Integer, _
    pH As Integer, _
    pMaxW As Integer, _
    pMaxH As Integer _
)
    Dim X2 As Integer, Y2 As Integer
    If pW < 0 Then GoTo ERR_OUTRANGE
    If pH < 0 Then GoTo ERR_OUTRANGE
    If pX > pMaxW Then GoTo ERR_OUTRANGE
    If pY > pMaxH Then GoTo ERR_OUTRANGE
    X2 = pX + pW - 1
    Y2 = pY + pH - 1
    If X2 < 0 Or Y2 < 0 Then GoTo ERR_OUTRANGE
    If pX < 0 Then pX = 0
    If pY < 0 Then pY = 0
    If X2 > pMaxW Then X2 = pMaxW
    If Y2 > pMaxH Then Y2 = pMaxH
    pW = X2 - pX + 1
    pH = Y2 - pY + 1
    Exit Sub
ERR_OUTRANGE:
    pX = 0: pY = 0: pW = 0: pH = 0
End Sub


'------------------------------------------------------------------------------
'@summary   由两个16位整数合成一个32位整数
'------------------------------------------------------------------------------
'@param     低16位整数
'@param     高16位整数
Private Function MakeDWord(pHi As Integer, pLow As Integer) As Long
    MakeDWord = (CLng(pHi) * &H10000) Or (pLow And &HFFFF&)
End Function


'//////////////////////////////////////////////////////////////////////////////
'//
'//      继承实现
'//
'//////////////////////////////////////////////////////////////////////////////


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有属性
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'@summary   背景色
'------------------------------------------------------------------------------
'@remark    值范围：0 - 15
Public Property Get BackColor() As Byte
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    BackColor = (tmp.wAttributes And &HF0) \ &HF
End Property
Public Property Let BackColor(v As Byte)
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    tmp.wAttributes = (tmp.wAttributes And (Not &HF0)) Or ((v And &HF) * &H10)
    SetConsoleTextAttribute m_StdOut, tmp.wAttributes
End Property


'------------------------------------------------------------------------------
'@summary   缓冲区高度
'------------------------------------------------------------------------------
Public Property Get BufferHeight() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    BufferHeight = tmp.dwSize.y
End Property


'------------------------------------------------------------------------------
'@summary   缓冲区宽度
'------------------------------------------------------------------------------
Public Property Get BufferWidth() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    BufferWidth = tmp.dwSize.x
End Property


'------------------------------------------------------------------------------
'@summary   与控制台相关的窗口句柄
'------------------------------------------------------------------------------
Public Property Get ConsoleWindow() As Long
    ConsoleWindow = GetConsoleWindow()
End Property


'------------------------------------------------------------------------------
'@summary   光标所在列
'------------------------------------------------------------------------------
Public Property Get CursorCol() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    CursorCol = tmp.dwCursorPosition.x
End Property
Public Property Let CursorCol(v As Integer)
    SetConsoleCursorPosition m_StdOut, MakeDWord(CursorRow, v)
End Property


'------------------------------------------------------------------------------
'@summary   光标所在行
'------------------------------------------------------------------------------
Public Property Get CursorRow() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    CursorRow = tmp.dwCursorPosition.y
End Property
Public Property Let CursorRow(v As Integer)
    SetConsoleCursorPosition m_StdOut, MakeDWord(v, CursorCol)
End Property


'------------------------------------------------------------------------------
'@summary   光标大小
'------------------------------------------------------------------------------
'@param     值范围：1 - 100
Public Property Get CursorSize() As Long
    Dim tmp As CONSOLE_CURSOR_INFO
    GetConsoleCursorInfo m_StdOut, tmp
    CursorSize = tmp.dwSize
End Property
Public Property Let CursorSize(v As Long)
    Dim tmp As CONSOLE_CURSOR_INFO
    GetConsoleCursorInfo m_StdOut, tmp
    tmp.dwSize = v
    SetConsoleCursorInfo m_StdOut, tmp
End Property


'------------------------------------------------------------------------------
'@summary   光标是否可见
'------------------------------------------------------------------------------
Public Property Get CursorVisable() As Boolean
    Dim tmp As CONSOLE_CURSOR_INFO
    GetConsoleCursorInfo m_StdOut, tmp
    CursorVisable = tmp.bVisible
End Property
Public Property Let CursorVisable(v As Boolean)
    Dim tmp As CONSOLE_CURSOR_INFO
    GetConsoleCursorInfo m_StdOut, tmp
    tmp.bVisible = v
    SetConsoleCursorInfo m_StdOut, tmp
End Property


'------------------------------------------------------------------------------
'@summary   前景色
'------------------------------------------------------------------------------
'@remark    值范围：0 - 15
Public Property Get ForeColor() As Byte
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    ForeColor = tmp.wAttributes And &HF
End Property
Public Property Let ForeColor(v As Byte)
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    tmp.wAttributes = (tmp.wAttributes And (Not &HF)) Or (v And &HF)
    SetConsoleTextAttribute m_StdOut, tmp.wAttributes
End Property


'------------------------------------------------------------------------------
'@summary   是否全屏
'------------------------------------------------------------------------------
Public Property Get FullScreen() As Boolean
    Dim ret As Long
    GetConsoleDisplayMode ret
    FullScreen = (ret <> 0)
End Property
Public Property Let FullScreen(v As Boolean)
    Dim tmp As COORD
    SetConsoleDisplayMode m_StdOut, _
        IIf(v, CONSOLE_FULLSCREEN_MODE, CONSOLE_WINDOWED_MODE), tmp
End Property


'------------------------------------------------------------------------------
'@summary   输入代码页
'------------------------------------------------------------------------------
Public Property Get InputCodePage() As Long
    InputCodePage = GetConsoleCP()
End Property
Public Property Let InputCodePage(v As Long)
    SetConsoleCP v
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_ECHO_INPUT
'------------------------------------------------------------------------------
Public Property Get InputModeEcho() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeEcho = ((ret And ENABLE_ECHO_INPUT) = ENABLE_ECHO_INPUT)
End Property
Public Property Let InputModeEcho(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_ECHO_INPUT, _
        ret And (Not ENABLE_ECHO_INPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_INSERT_MODE
'------------------------------------------------------------------------------
Public Property Get InputModeInsert() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeInsert = ((ret And ENABLE_INSERT_MODE) = ENABLE_INSERT_MODE)
End Property
Public Property Let InputModeInsert(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_INSERT_MODE, _
        ret And (Not ENABLE_INSERT_MODE))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_LINE_INPUT
'------------------------------------------------------------------------------
Public Property Get InputModeLine() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeLine = ((ret And ENABLE_LINE_INPUT) = ENABLE_LINE_INPUT)
End Property
Public Property Let InputModeLine(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_LINE_INPUT, _
        ret And (Not ENABLE_LINE_INPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_MOUSE_INPUT
'------------------------------------------------------------------------------
Public Property Get InputModeMouse() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeMouse = ((ret And ENABLE_MOUSE_INPUT) = ENABLE_MOUSE_INPUT)
End Property
Public Property Let InputModeMouse(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_MOUSE_INPUT, _
        ret And (Not ENABLE_MOUSE_INPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_PROCESSED_INPUT
'------------------------------------------------------------------------------
Public Property Get InputModeProcessed() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeProcessed = _
        ((ret And ENABLE_PROCESSED_INPUT) = ENABLE_PROCESSED_INPUT)
End Property
Public Property Let InputModeProcessed(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_PROCESSED_INPUT, _
        ret And (Not ENABLE_PROCESSED_INPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_QUICK_EDIT_MODE
'------------------------------------------------------------------------------
Public Property Get InputModeQuickEdit() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeQuickEdit = _
        ((ret And ENABLE_QUICK_EDIT_MODE) = ENABLE_QUICK_EDIT_MODE)
End Property
Public Property Let InputModeQuickEdit(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_QUICK_EDIT_MODE, _
        ret And (Not ENABLE_QUICK_EDIT_MODE))
End Property


'------------------------------------------------------------------------------
'@summary   输入模式 ENABLE_WINDOW_INPUT
'------------------------------------------------------------------------------
Public Property Get InputModeWindow() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    InputModeWindow = ((ret And ENABLE_WINDOW_INPUT) = ENABLE_WINDOW_INPUT)
End Property
Public Property Let InputModeWindow(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdIn, ret
    SetConsoleMode m_StdIn, IIf(v, ret Or ENABLE_WINDOW_INPUT, _
        ret And (Not ENABLE_WINDOW_INPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输出代码页
'------------------------------------------------------------------------------
Public Property Get OutputCodePage() As Long
    OutputCodePage = GetConsoleOutputCP()
End Property
Public Property Let OutputCodePage(v As Long)
    SetConsoleOutputCP v
End Property


'------------------------------------------------------------------------------
'@summary   输出模式 ENABLE_PROCESSED_OUTPUT
'------------------------------------------------------------------------------
Public Property Get OutputModeProcessed() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdOut, ret
    OutputModeProcessed = _
        ((ret And ENABLE_PROCESSED_OUTPUT) = ENABLE_PROCESSED_OUTPUT)
End Property
Public Property Let OutputModeProcessed(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdOut, ret
    SetConsoleMode m_StdOut, IIf(v, ret Or ENABLE_PROCESSED_OUTPUT, _
        ret And (Not ENABLE_PROCESSED_OUTPUT))
End Property


'------------------------------------------------------------------------------
'@summary   输出模式 ENABLE_WRAP_AT_EOL_OUTPUT
'------------------------------------------------------------------------------
Public Property Get OutputModeWrapAtEol() As Boolean
    Dim ret As Long
    GetConsoleMode m_StdOut, ret
    OutputModeWrapAtEol = _
        ((ret And ENABLE_WRAP_AT_EOL_OUTPUT) = ENABLE_WRAP_AT_EOL_OUTPUT)
End Property
Public Property Let OutputModeWrapAtEol(v As Boolean)
    Dim ret As Long
    GetConsoleMode m_StdOut, ret
    SetConsoleMode m_StdOut, IIf(v, ret Or ENABLE_WRAP_AT_EOL_OUTPUT, _
        ret And (Not ENABLE_WRAP_AT_EOL_OUTPUT))
End Property


'------------------------------------------------------------------------------
'@summary   标准错误句柄
'------------------------------------------------------------------------------
Public Property Get StdErr() As Long
    StdErr = m_StdErr
End Property
Public Property Let StdErr(v As Long)
    SetStdHandle STD_ERROR_HANDLE, v
    m_StdErr = GetStdHandle(STD_ERROR_HANDLE)
End Property


'------------------------------------------------------------------------------
'@summary   标准输入句柄
'------------------------------------------------------------------------------
Public Property Get StdIn() As Long
    StdIn = m_StdIn
End Property
Public Property Let StdIn(v As Long)
    SetStdHandle STD_INPUT_HANDLE, v
    m_StdIn = GetStdHandle(STD_INPUT_HANDLE)
End Property


'------------------------------------------------------------------------------
'@summary   标准输出句柄
'------------------------------------------------------------------------------
Public Property Get StdOut() As Long
    StdOut = m_StdOut
End Property
Public Property Let StdOut(v As Long)
    SetStdHandle STD_OUTPUT_HANDLE, v
    m_StdOut = GetStdHandle(STD_OUTPUT_HANDLE)
End Property


'------------------------------------------------------------------------------
'@summary   控制台标题
'------------------------------------------------------------------------------
Public Property Get Title() As String
    Dim buf As String
    Dim ret As Long
    buf = String(CONSOLE_BUFFER_MAXSIZE, Chr(0))
    ret = GetConsoleTitle(buf, CONSOLE_BUFFER_MAXSIZE)
    Title = LeftB(buf, ret + 1)
End Property
Public Property Let Title(ByVal v As String)
    SetConsoleTitle v
End Property


'------------------------------------------------------------------------------
'@summary   视口高度
'------------------------------------------------------------------------------
Public Property Get ViewHeight() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    ViewHeight = tmp.srWindow.Bottom - tmp.srWindow.Top + 1
End Property


'------------------------------------------------------------------------------
'@summary   最大视口高度
'------------------------------------------------------------------------------
Public Property Get ViewHeightMax() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    ViewHeightMax = tmp.dwMaximumWindowSize.y
End Property


'------------------------------------------------------------------------------
'@summary   视口宽度
'------------------------------------------------------------------------------
Public Property Get ViewWidth() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    ViewWidth = tmp.srWindow.Right - tmp.srWindow.Left + 1
End Property


'------------------------------------------------------------------------------
'@summary   最大视口宽度
'------------------------------------------------------------------------------
Public Property Get ViewWidthMax() As Integer
    Dim tmp As CONSOLE_SCREEN_BUFFER_INFO
    GetConsoleScreenBufferInfo m_StdOut, tmp
    ViewWidthMax = tmp.dwMaximumWindowSize.x
End Property


'//////////////////////////////////////////////////////////////////////////////
'//
'//      公有方法
'//
'//////////////////////////////////////////////////////////////////////////////


'------------------------------------------------------------------------------
'@summary   清屏
'------------------------------------------------------------------------------
'@param     是否恢复属性
Public Sub Clear(Optional ByVal pToDefault As Boolean = False)
    FillText 0, 0, BufferWidth * BufferHeight, 32
    If pToDefault Then
        FillTextAttribute 0, 0, BufferWidth * BufferHeight, _
            m_DefaultForeColor, m_DefaultBackColor, False
        ForeColor = m_DefaultForeColor
        BackColor = m_DefaultBackColor
    End If
    MoveCursor 0, 0
End Sub


'------------------------------------------------------------------------------
'@summary   填充矩形文本
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     宽度
'@param     高度
'@param     字符
'@return    实际填充数量
Public Function FillRectText( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pWidth As Integer, _
    ByVal pHeight As Integer, _
    ByVal pChar As Byte _
) As Long
    Dim i As Integer, ret As Long
    ClipRect pCol, pRow, pWidth, pHeight, BufferWidth - 1, BufferHeight - 1
    If pWidth <= 0 Or pHeight <= 0 Then Exit Function
    For i = pRow To pRow + pHeight - 1
        ret = ret + FillText(i, pCol, pWidth, pChar)
    Next
    FillRectText = ret
End Function


'------------------------------------------------------------------------------
'@summary   填充矩形属性
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     宽度
'@param     高度
'@param     前景色
'@param     背景色
'@param     反转色
'@return    实际填充数量
Public Function FillRectAttribute( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pWidth As Integer, _
    ByVal pHeight As Integer, _
    ByVal pForeColor As Byte, _
    ByVal pBackColor As Byte, _
    ByVal pReverseColor As Boolean _
) As Long
    Dim i As Integer, ret As Long
    ClipRect pCol, pRow, pWidth, pHeight, BufferWidth - 1, BufferHeight - 1
    If pWidth = 0 Or pHeight = 0 Then Exit Function
    For i = pRow To pRow + pHeight - 1
        ret = ret + FillTextAttribute(i, pCol, pWidth, pForeColor, _
            pBackColor, pReverseColor)
    Next
    FillRectAttribute = ret
End Function


'------------------------------------------------------------------------------
'@summary   填充矩形文本，扩展
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     宽度
'@param     高度
'@param     字符
'@param     前景色
'@param     背景色
'@param     反转色
'@return    实际填充数量
Public Function FillRectEx( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pWidth As Integer, _
    ByVal pHeight As Integer, _
    ByVal pChar As Byte, _
    ByVal pForeColor As Byte, _
    ByVal pBackColor As Byte, _
    ByVal pReverseColor As Boolean _
) As Long
    FillRectText pRow, pCol, pWidth, pHeight, pChar
    FillRectEx = FillRectAttribute(pRow, pCol, pWidth, pHeight, pForeColor, _
        pBackColor, pReverseColor)
End Function


'------------------------------------------------------------------------------
'@summary   填充文本
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     字符
'@param     数量
'@return    实际填充数量
Public Function FillText( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long, _
    ByVal pChar As Byte _
) As Long
    Dim ret As Long
    If pRow < 0 Or pCol < 0 Then Exit Function
    FillConsoleOutputCharacter m_StdOut, pChar, pCount, _
        MakeDWord(pRow, pCol), ret
    FillText = ret
End Function


'------------------------------------------------------------------------------
'@summary   填充属性
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     数量
'@param     前景色
'@param     背景色
'@param     反转色
'@return    实际填充数量
Public Function FillTextAttribute( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long, _
    ByVal pForeColor As Byte, _
    ByVal pBackColor As Byte, _
    ByVal pReverseColor As Boolean _
) As Long
    Dim ret As Long, attr As Long
    If pRow < 0 Or pCol < 0 Then Exit Function
    attr = MakeAttribute(pForeColor, pBackColor, pReverseColor)
    FillConsoleOutputAttribute m_StdOut, attr, pCount, _
        MakeDWord(pRow, pCol), ret
    FillTextAttribute = ret
End Function


'------------------------------------------------------------------------------
'@summary   填充文本，扩展
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     数量
'@param     字符
'@param     前景色
'@param     背景色
'@param     反转色
'@return    实际填充数量
Public Function FillTextEx( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long, _
    ByVal pChar As Byte, _
    ByVal pForeColor As Byte, _
    ByVal pBackColor As Byte, _
    ByVal pReverseColor As Boolean _
) As Long
    FillText pRow, pCol, pCount, pChar
    FillTextEx = FillTextAttribute(pRow, pCol, pCount, pForeColor, _
        pBackColor, pReverseColor)
End Function


'------------------------------------------------------------------------------
'@summary   合成字符属性
'------------------------------------------------------------------------------
'@param     前景色
'@param     背景色
'@param     反转色
'@return    合成的属性
Public Function MakeAttribute( _
    ByVal pForeColor As Byte, _
    ByVal pBackColor As Byte, _
    ByVal pReverseColor As Boolean _
) As Integer
    Dim ret As Integer
    If pReverseColor Then ret = ret Or COMMON_LVB_REVERSE_VIDEO
    ret = ret Or (pForeColor And &HF)
    ret = ret Or ((pBackColor And &HF) * &H10)
    MakeAttribute = ret
End Function


'------------------------------------------------------------------------------
'@summary   移动光标
'------------------------------------------------------------------------------
'@param     行
'@param     列
Public Sub MoveCursor(ByVal pRow As Integer, ByVal pCol As Integer)
    SetConsoleCursorPosition m_StdOut, MakeDWord(pRow, pCol)
End Sub


'------------------------------------------------------------------------------
'@summary   读取缓冲属性
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     读取多少个属性
'@return    属性数组，每个值1-4位为前景色，5-8位为背景色，15位为反转色
Public Function ReadBufferAttribute( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long _
) As Integer()
    Dim buf() As Integer, ret As Long
    If pCount <= 0 Then Exit Function
    ReDim buf(pCount - 1) As Integer
    ReadConsoleOutputAttribute _
        m_StdOut, buf(0), pCount, MakeDWord(pRow, pCol), ret
    ReDim Preserve buf(ret - 1) As Integer
    ReadBufferAttribute = buf
End Function


'------------------------------------------------------------------------------
'@summary   读取缓冲文本
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     读取多少个
'@return    读取的文本
Public Function ReadBufferText( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long _
) As String
    Dim buf As String, ret As Long
    If pCount <= 0 Then Exit Function
    buf = String(pCount, Chr(0))
    ReadConsoleOutputCharacter _
        m_StdOut, buf, pCount, MakeDWord(pRow, pCol), ret
    ReadBufferText = Left(buf, ret)
End Function


'------------------------------------------------------------------------------
'@summary   读取键盘按键
'------------------------------------------------------------------------------
'@return    按键的虚拟键码，参考KeyCodeConstants枚局常量
Public Function ReadKey() As Integer
    Dim buf As INPUT_RECORD
    Dim ret As Long
    FlushConsoleInputBuffer m_StdIn
    Do While True
        ReadConsoleInput m_StdIn, buf, 1, ret
        If ret = 1 And buf.EventType = KEY_EVENT Then Exit Do
    Loop
    ReadKey = buf.EventContent.wVirtualKeyCode
End Function


'------------------------------------------------------------------------------
'@summary   读取整行输入
'------------------------------------------------------------------------------
'@param     提示
'@return    输入文本，不包括回车换行
Public Function ReadLine(Optional pPrompt As String = "") As String
    Dim buf As String
    Dim ret As Long
    If pPrompt <> "" Then WriteText pPrompt
    buf = String(CONSOLE_BUFFER_MAXSIZE, Chr(0))
    ReadConsole m_StdIn, buf, CONSOLE_BUFFER_MAXSIZE, ret, 0
    ReadLine = Left(buf, ret - 2)
End Function


'------------------------------------------------------------------------------
'@summary   读取密码
'------------------------------------------------------------------------------
'@param     提示
'@return    输入文本，不包括回车换行
Public Function ReadPassword(Optional pPrompt As String = "") As String
    Dim buf As String
    Dim ret As Long, preEchoFlag As Boolean
    If pPrompt <> "" Then WriteText pPrompt
    preEchoFlag = InputModeEcho
    InputModeEcho = False
    buf = String(CONSOLE_BUFFER_MAXSIZE, Chr(0))
    ReadConsole m_StdIn, buf, CONSOLE_BUFFER_MAXSIZE, ret, 0
    InputModeEcho = preEchoFlag
    ReadPassword = Left(buf, ret - 2)
End Function


'------------------------------------------------------------------------------
'@summary   反转前景色和背景色
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     数量
'@param     是否反转
'@return    实际反转数量
Public Function ReverseTextColor( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pCount As Long, _
    ByVal pReverse As Boolean _
) As Long
    Dim attr() As Integer
    Dim i As Long, c As Long
    attr = ReadBufferAttribute(pRow, pCol, pCount)
    c = UBound(attr)
    If pReverse Then
        For i = 0 To c
            attr(i) = attr(i) Or COMMON_LVB_REVERSE_VIDEO
        Next
    Else
        For i = 0 To c
            attr(i) = attr(i) And (Not COMMON_LVB_REVERSE_VIDEO)
        Next
    End If
    WriteBufferAttribute pRow, pCol, attr, c
    ReverseTextColor = c
End Function


'------------------------------------------------------------------------------
'@summary   反转矩形前景色和背景色
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     宽度
'@param     高度
'@param     是否反转
'@return    实际反转数量
Public Function ReverseRectColor( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByVal pWidth As Integer, _
    ByVal pHeight As Integer, _
    ByVal pReverse As Boolean _
) As Long
    Dim i As Integer, ret As Long
    If pWidth < 0 Or pHeight < 0 Then Exit Function
    For i = pRow To pRow + pHeight - 1
        ret = ret + ReverseTextColor(i, pCol, pWidth, pReverse)
    Next
    ReverseRectColor = ret
End Function


'------------------------------------------------------------------------------
'@summary   设置视口尺寸
'------------------------------------------------------------------------------
'@param     宽
'@param     高
Public Sub SetViewSize(ByVal pWidth As Integer, ByVal pHeight As Integer)
    Dim tmp As SMALL_RECT
    tmp.Right = pWidth - 1
    tmp.Bottom = pHeight - 1
    SetConsoleWindowInfo m_StdOut, True, tmp
End Sub


'------------------------------------------------------------------------------
'@summary   分解字符属性
'------------------------------------------------------------------------------
'@param     属性
'@param     前景色
'@param     背景色
'@param     反转色
Public Function SplitAttribute( _
    ByVal pAttribute As Integer, _
    ByRef pForeColor As Byte, _
    ByRef pBackColor As Byte, _
    ByRef pReverseColor As Boolean _
)
    pForeColor = pAttribute And &HF
    pBackColor = (pAttribute And &HF0) \ &H10
    pReverseColor = ((pAttribute And COMMON_LVB_REVERSE_VIDEO) = _
        COMMON_LVB_REVERSE_VIDEO)
End Function


'------------------------------------------------------------------------------
'@summary   测试文本长度，单字节算1，双字节算2
'------------------------------------------------------------------------------
'@param     内容
Public Function StrLen(pStr As String) As Long
    StrLen = lstrlen(pStr)
End Function


'------------------------------------------------------------------------------
'@summary   写入缓冲属性
'------------------------------------------------------------------------------
'@param     行
'@param     列
'@param     属性数组，每个属性值1-4位为前景色，5-8位为背景色，15位为反转色
'@param     写入多少个字符
'@return    实际写入多少个字符
Public Function WriteBufferAttribute( _
    ByVal pRow As Integer, _
    ByVal pCol As Integer, _
    ByRef pAttributes() As Integer, _
    ByVal pCount As Long _
) As Long
    Dim ret As Long
    If pCount <= 0 Then Exit Function
    If UBound(pAttributes) < 0 Then Exit Function
    WriteConsoleOutputAttribute _
        m_StdOut, pAttributes(0), pCount, MakeDWord(pRow, pCol), ret
    WriteBufferAttribute = ret
End Function


'------------------------------------------------------------------------------
'@summary   写入整行文本
'------------------------------------------------------------------------------
'@param     内容
'@return    实际写入字符个数
Public Function WriteLine(Optional ByVal pBuf As String = "") As Long
    Dim ret As Long
    ret = WriteText(pBuf)
    ret = ret + WriteText(vbCrLf)
    WriteLine = ret
End Function


'------------------------------------------------------------------------------
'@summary   写入整行文本,扩展
'------------------------------------------------------------------------------
'@param     内容
'@param     前景色
'@param     背景色
'@return    实际写入字符个数
Public Function WriteLineEx( _
    ByVal pBuf As String, _
    Optional ByVal pForeColor As Integer = -1, _
    Optional ByVal pBackColor As Integer = -1 _
) As Long
    Dim ret As Long, foreIsSet As Boolean, backIsSet As Boolean
    Dim preForeColor As Byte, preBackColor As Byte
    foreIsSet = pForeColor > -1
    backIsSet = pBackColor > -1
    If foreIsSet Then
        preForeColor = ForeColor
        ForeColor = pForeColor
    End If
    If backIsSet Then
        preBackColor = BackColor
        BackColor = pBackColor
    End If
    ret = WriteLine(pBuf)
    If foreIsSet Then ForeColor = preForeColor
    If backIsSet Then BackColor = preBackColor
    WriteLineEx = ret
End Function


'------------------------------------------------------------------------------
'@summary   写入文本
'------------------------------------------------------------------------------
'@param     内容
'@return    实际写入字符个数
Public Function WriteText(pBuf As String) As Long
    Dim ret As Long
    If pBuf = "" Then Exit Function
    WriteConsole m_StdOut, pBuf, lstrlen(pBuf), ret, 0
    WriteText = ret
End Function


'------------------------------------------------------------------------------
'@summary   写入文本,扩展
'------------------------------------------------------------------------------
'@param     内容
'@param     前景色
'@param     背景色
'@return    实际写入字符个数
Public Function WriteTextEx( _
    ByVal pBuf As String, _
    Optional ByVal pForeColor As Integer = -1, _
    Optional ByVal pBackColor As Integer = -1 _
) As Long
    Dim ret As Long, foreIsSet As Boolean, backIsSet As Boolean
    Dim preForeColor As Byte, preBackColor As Byte
    foreIsSet = pForeColor > -1
    backIsSet = pBackColor > -1
    If foreIsSet Then
        preForeColor = ForeColor
        ForeColor = pForeColor
    End If
    If backIsSet Then
        preBackColor = BackColor
        BackColor = pBackColor
    End If
    ret = WriteText(pBuf)
    If foreIsSet Then ForeColor = preForeColor
    If backIsSet Then BackColor = preBackColor
    WriteTextEx = ret
End Function
