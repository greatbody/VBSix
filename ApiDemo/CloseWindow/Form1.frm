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
   Begin VB.CommandButton Command1 
      Caption         =   "隐藏窗口"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "有效：可以最小化窗口"
      Height          =   255
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long


Private Sub Command1_Click()
    CloseWindow Me.hwnd
End Sub
