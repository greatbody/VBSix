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
   StartUpPosition =   3  '����ȱʡ
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
    BeginPath Me.hdc '������豸����������·����֧[���е�GDI���Ʋ������������Ҫչ�ֵ�����]

    Me.Line (Command2.Left, Command2.Top)-(Command2.Width, Command2.Height), vbRed, BF
    
    EndPath Me.hdc '������豸�����йر�·����֧
    hdcID = PathToRegion(Me.hdc) '�����ʻ��Ƶ�·��ת���������豸������
    SetWindowRgn Me.hwnd, hdcID, 1 '��ָ���豸�����еĻ�����������Ϊ��ʾ�������������У����細�ڵĻ��ƾͻᰴ�ջ��ʻ���ͼ�ε�·�������ƣ���֤�˴��ڵ���״
    DeleteObject hdcID 'ɾ����ʱ���豸����
End Sub

Private Sub Form_Load()
    'PrintForms Me, "Ϊ������ĺ�ƽ��"
End Sub

Private Sub PrintForms(ByRef Frm As Form, ByVal PrintWord As String)
    'frmΪ�ı�Ĵ���
    'printwordΪ�������ֵ���״
    Dim hdcID As Long
    BeginPath Frm.hdc '������豸����������·����֧[���е�GDI���Ʋ������������Ҫչ�ֵ�����]
    With Frm
        .CurrentX = 0
        .CurrentY = 220
        .FontSize = 170
    End With
    Frm.Print PrintWord '��ӡ�������״
    Frm.Circle (200, 200), 200
    EndPath Frm.hdc '������豸�����йر�·����֧
    hdcID = PathToRegion(Frm.hdc) '�����ʻ��Ƶ�·��ת���������豸������
    SetWindowRgn Frm.hwnd, hdcID, 1 '��ָ���豸�����еĻ�����������Ϊ��ʾ�������������У����細�ڵĻ��ƾͻᰴ�ջ��ʻ���ͼ�ε�·�������ƣ���֤�˴��ڵ���״
    DeleteObject hdcID 'ɾ����ʱ���豸����
End Sub

Private Sub Timer1_Timer()
    'frmΪ�ı�Ĵ���
    'printwordΪ�������ֵ���״
    Static k As Long
    Static lngAdds As Long
    If lngAdds = 0 Then lngAdds = 4
    Dim hdcID As Long
    BeginPath Me.hdc '������豸����������·����֧[���е�GDI���Ʋ������������Ҫչ�ֵ�����]

    Me.Circle (150, 100), k
    EndPath Me.hdc '������豸�����йر�·����֧
    hdcID = PathToRegion(Me.hdc) '�����ʻ��Ƶ�·��ת���������豸������
    SetWindowRgn Me.hwnd, hdcID, 1 '��ָ���豸�����еĻ�����������Ϊ��ʾ�������������У����細�ڵĻ��ƾͻᰴ�ջ��ʻ���ͼ�ε�·�������ƣ���֤�˴��ڵ���״
    DeleteObject hdcID 'ɾ����ʱ���豸����
    k = k + lngAdds
    If k > 200 Then lngAdds = -4
    If k < 10 Then lngAdds = 4
End Sub
