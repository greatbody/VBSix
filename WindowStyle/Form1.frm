VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1320
      Top             =   2520
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function PaintRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function PathToRegion Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Command2_Click()
    BeginPath Me.hdc '在这个设备场景中启动路径分支[所有的GDI绘制操作都会描出需要展现的区域]

    Me.Line (Command2.Left, Command2.Top)-(Command2.Width, Command2.Height), vbRed, BF
    
    EndPath Me.hdc '在这个设备场景中关闭路径分支
    hdcID = PathToRegion(Me.hdc) '将画笔绘制的路径转换到其它设备场景中
    SetWindowRgn Me.hwnd, hdcID, 1 '将指定设备场景中的绘制区域设置为显示区域，其它场景中，例如窗口的绘制就会按照画笔绘制图形的路径来绘制，保证了窗口的形状
    DeleteObject hdcID '删除临时的设备场景
End Sub

Private Sub Form_Load()
    'PrintForms Me, "为了世界的和平！"
End Sub

Private Sub PrintForms(ByRef Frm As Form, ByVal PrintWord As String)
    'frm为改变的窗体
    'printword为窗体文字的形状
    Dim hdcID As Long
    BeginPath Frm.hdc '在这个设备场景中启动路径分支[所有的GDI绘制操作都会描出需要展现的区域]
    With Frm
        .CurrentX = 0
        .CurrentY = 220
        .FontSize = 170
    End With
    Frm.Print PrintWord '打印窗体的形状
    Frm.Circle (200, 200), 200
    EndPath Frm.hdc '在这个设备场景中关闭路径分支
    hdcID = PathToRegion(Frm.hdc) '将画笔绘制的路径转换到其它设备场景中
    SetWindowRgn Frm.hwnd, hdcID, 1 '将指定设备场景中的绘制区域设置为显示区域，其它场景中，例如窗口的绘制就会按照画笔绘制图形的路径来绘制，保证了窗口的形状
    DeleteObject hdcID '删除临时的设备场景
End Sub

Private Sub Timer1_Timer()
    'frm为改变的窗体
    'printword为窗体文字的形状
    Static k As Long
    Static lngAdds As Long
    If lngAdds = 0 Then lngAdds = 4
    Dim hdcID As Long
    BeginPath Me.hdc '在这个设备场景中启动路径分支[所有的GDI绘制操作都会描出需要展现的区域]

    Me.Circle (150, 100), k
    EndPath Me.hdc '在这个设备场景中关闭路径分支
    hdcID = PathToRegion(Me.hdc) '将画笔绘制的路径转换到其它设备场景中
    SetWindowRgn Me.hwnd, hdcID, 1 '将指定设备场景中的绘制区域设置为显示区域，其它场景中，例如窗口的绘制就会按照画笔绘制图形的路径来绘制，保证了窗口的形状
    DeleteObject hdcID '删除临时的设备场景
    k = k + lngAdds
    If k > 200 Then lngAdds = -4
    If k < 10 Then lngAdds = 4
End Sub
