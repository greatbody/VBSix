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
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   2400
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "弹出模态消息"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   2640
      ScaleHeight     =   555
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "拦截消息"
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
Dim pric As New MessageProc

Private Sub Command1_Click()
    pric.Init Me.hWnd
    preFunc = pric.OriginProcFunc
    pric.SetNewFunc Me.hWnd, AddressOf Pro
End Sub

Private Sub Command2_Click()
    MsgBox "Mode"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pric.CancelWinProc
End Sub

Public Sub DrawText(ByVal txtContent As String)
    Picture1.AutoRedraw = True
    Picture1.Cls
    Picture1.Print txtContent
    Picture1.AutoRedraw = False
End Sub

Public Sub DrawToPic(ByVal Content As String, ByVal picID As Integer)
    Select Case picID
        Case 1
            Picture1.AutoRedraw = True
            Picture1.Cls
            Picture1.Print Content
            Picture1.AutoRedraw = False
        Case 2
            Picture2.AutoRedraw = True
            Picture2.Cls
            Picture2.Print Content
            Picture2.AutoRedraw = False
    End Select
End Sub
