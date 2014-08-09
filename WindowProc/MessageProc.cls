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
'*************************************************************************
'**模 块 名：WindowProc - MessageProc
'**说    明：http://www.extencent.com 版权所有2014 - 2015(C)1
'**创 建 人：孙瑞
'**日    期：2014-08-09 15:05:14
'**修 改 人：孙瑞
'**日    期：2014-08-09 15:05:14
'**描    述：
'**版    本：V1.0.0
'*************************************************************************

'*********定义系统API
'获取窗口程序地址
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Const GWL_EXSTYLE = (-20) '获得扩展窗口风格
Private Const GWL_HINSTANCE = (-6) '获取应用实例句柄
Private Const GWL_HWNDPARENT = (-8) '如果父窗口存在，则获取父窗口句柄
Private Const GWL_ID = (-12) '获得窗口标示
Private Const GWL_STYLE = (-16) '获得窗口风格
Private Const GWL_USERDATA = (-21) '获得窗口有关的32位数值
Private Const GWL_WNDPROC = (-4) '获得窗口程序的地址
'设置窗口属性【所涉及参数值，同上】
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'指定的窗口数据是在缓存中保存的，因此在调用SetWindowLong之后再调用SetWindowPos函数才能使SetWindowLong函数所作的改变生效。
Private lngDefWindowFunc As Long '记录默认的窗口处理程序的地址，方便结束的时候交还
Private frmHwnd As Long '窗口的句柄
Public Property Get OriginProcFunc()
    OriginProcFunc = lngDefWindowFunc
End Property



'*************************************************************************
'**函 数 名：Init
'**输    入：ByRef Frm(Form) -
'**输    出：无
'**功能描述：初始化类，设置当前窗体的默认值，并且获取当前窗体的默认窗体处理程序
'**全局变量：
'**调用模块：
'**作    者：孙瑞
'**日    期：2014-08-09 15:05:31
'**修 改 人：孙瑞
'**日    期：2014-08-09 15:05:31
'**版    本：V1.0.0
'*************************************************************************
Public Sub Init(ByVal hWnd As Long)
    frmHwnd = hWnd
    lngDefWindowFunc = GetWindowLong(hWnd, GWL_WNDPROC)
End Sub

'*************************************************************************
'**函 数 名：SetNewFunc
'**输    入：ByRef Frm(Form)     -
'**        ：ByVal lngFunc(Long) -
'**输    出：无
'**功能描述：为窗口设置新的窗口程序
'**全局变量：
'**调用模块：
'**作    者：孙瑞
'**日    期：2014-08-09 15:04:08
'**修 改 人：孙瑞
'**日    期：2014-08-09 15:04:08
'**版    本：V1.0.0
'*************************************************************************
Public Sub SetNewFunc(ByVal hWnd As Long, ByVal lngFunc As Long)
    Dim lngRet As Long
    lngRet = SetWindowLong(hWnd, GWL_WNDPROC, lngFunc)
    If lngRet = 0 Then
        'MsgBox "出现故障", , "Window Proc Demo"
    End If
End Sub


'*************************************************************************
'**函 数 名：CancelWinProc
'**输    入：无
'**输    出：无
'**功能描述：恢复窗口原来的窗口函数
'**全局变量：
'**调用模块：
'**作    者：孙瑞
'**日    期：2014-08-09 15:02:49
'**修 改 人：孙瑞
'**日    期：2014-08-09 15:02:49
'**版    本：V1.0.0
'*************************************************************************
Public Sub CancelWinProc()
    SetNewFunc frmHwnd, lngDefWindowFunc
End Sub

