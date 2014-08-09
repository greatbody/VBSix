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
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   600
      Top             =   2160
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "没什么用"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

Private Sub Timer1_Timer()
    Dim r As Long
    r = BringWindowToTop(Me.hwnd)
    If r = 0 Then
        MsgBox "no"
    End If
End Sub
