VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ASP编程辅助工具"
   ClientHeight    =   4500
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   8520
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "Response填充"
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin VB.CommandButton Command1 
         Caption         =   "转换到剪贴板"
         Height          =   375
         Left            =   2880
         TabIndex        =   5
         Top             =   3720
         Width           =   1335
      End
      Begin VB.TextBox uOut 
         Height          =   1695
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1920
         Width           =   7455
      End
      Begin VB.TextBox uSource 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   7455
      End
      Begin VB.Label Label2 
         Caption         =   "网代码"
         Height          =   615
         Left            =   7800
         TabIndex        =   4
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "源代码"
         Height          =   615
         Left            =   7800
         TabIndex        =   3
         Top             =   480
         Width           =   255
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim r() As String
    uOut.Text = ""
    r = Split(uSource.Text, vbCrLf)
    For Each i In r
        If uOut.Text = "" Then
            uOut.Text = TransOne(i)
        Else
            uOut.Text = uOut.Text & vbCrLf & TransOne(i)
        End If
    Next i
    Clipboard.Clear
    Clipboard.SetText CStr(uOut.Text)
End Sub
Function TransOne(ByVal s As String) As String
    Dim r As String
    Dim outS As String
    r = Replace(s, """", """""")
    If outS = "" Then
        outS = "Response.Write(""" & r & """)"
    Else
        outS = outS & vbCrLf & "Response.Write(""" & r & """)"
    End If
    TransOne = outS
End Function
