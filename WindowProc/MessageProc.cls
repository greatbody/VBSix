VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageProc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private lngDefWindowFunc As Long '记录默认的窗口处理程序的地址，方便结束的时候交还

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()

End Sub

Public Sub Init(ByRef Frm As Form)
    
End Sub
