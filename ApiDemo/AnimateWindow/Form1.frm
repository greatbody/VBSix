VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long


'hwnd只对Form有效，其他像Picture1都无法产生效果。
'　　　dwTime是动画持续的时间，默认为200。
'dwFlags可取以下值:
'　　　　　AW_HOR_POSITIVE （  &H1  ） 从左到右打开窗口
'　　　　　AW_HOR_NEGATIVE （  &H2  ） 从右到左打开窗口
'　　　　　AW_VER_POSITIVE （  &H4  ） 从上到下打开窗口
'　　　　　AW_VER_NEGATIVE （  &H8  ） 从下到上打开窗口
'　　　　　AW_CENTER 　　　（  &H10 ） 看不出任何效果
'　　　　　AW_HIDE 　　　　（&H10000） 在窗体卸载时若想使用本函数就得加上此常量
'　　　　　AW_ACTIVATE 　　（&H20000） 在窗体通过本函数打开后，默认情况下会失去焦点，除非加上本常量
'　　　　　AW_SLIDE 　　　 （&H40000） 看不出任何效果
'　　　　　AW_BLEND 　　　 （&H80000） 淡入淡出效果
Private Sub Form_Load()
    AnimateWindow Me.hwnd, 500, &H4
End Sub

