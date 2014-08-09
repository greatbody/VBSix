Attribute VB_Name = "winpro"
Public preFunc As Long '窗口程序的地址
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function AdjustWindowRect Lib "user32" (lpRect As RECT, ByVal dwStyle As Long, ByVal bMenu As Long) As Long
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type



Public Function Pro(ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case Msg
        Case WM_ACTIVATEAPP
            ShowText "appActivate"
        Case WM_CANCELMOD
            ShowText "模态"
            Pro = 0&
            Exit Function
        Case WM_CHILDACTIVATE
            ShowText "WM_CHILDACTIVATE"
        Case WM_CLOSE
            ShowText "Want me Close?"
            Pro = 1&
            Exit Function
        Case WM_DESTROY
            Pro = 0&
            Exit Function
        Case WM_MOVE
            Call ShowText("move" & Now, 2)
            Pro = 0&
            Exit Function
    End Select
    Pro = CallWindowProc(preFunc, hWnd, Msg, wParam, lParam)
End Function

Public Sub ShowText(ByVal s As String, Optional ByVal PicID As Integer = 1)
    Call Form1.DrawToPic(s, PicID)
End Sub
