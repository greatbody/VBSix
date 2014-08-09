VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   ScaleHeight     =   1545
   ScaleWidth      =   4920
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   960
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2400
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    Picture1.AutoRedraw = True
    Picture1.Cls
    Picture1.Print Format(Now, "yyyy-MM-dd hh:mm:ss")
    Picture1.AutoRedraw = False
End Sub
