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
   StartUpPosition =   3  '����ȱʡ
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Long


'hwndֻ��Form��Ч��������Picture1���޷�����Ч����
'������dwTime�Ƕ���������ʱ�䣬Ĭ��Ϊ200��
'dwFlags��ȡ����ֵ:
'����������AW_HOR_POSITIVE ��  &H1  �� �����Ҵ򿪴���
'����������AW_HOR_NEGATIVE ��  &H2  �� ���ҵ���򿪴���
'����������AW_VER_POSITIVE ��  &H4  �� ���ϵ��´򿪴���
'����������AW_VER_NEGATIVE ��  &H8  �� ���µ��ϴ򿪴���
'����������AW_CENTER ��������  &H10 �� �������κ�Ч��
'����������AW_HIDE ����������&H10000�� �ڴ���ж��ʱ����ʹ�ñ������͵ü��ϴ˳���
'����������AW_ACTIVATE ������&H20000�� �ڴ���ͨ���������򿪺�Ĭ������»�ʧȥ���㣬���Ǽ��ϱ�����
'����������AW_SLIDE ������ ��&H40000�� �������κ�Ч��
'����������AW_BLEND ������ ��&H80000�� ���뵭��Ч��
Private Sub Form_Load()
    AnimateWindow Me.hwnd, 500, &H4
End Sub

