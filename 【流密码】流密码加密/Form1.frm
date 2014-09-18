VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "流密码加密实例代码"
   ClientHeight    =   4905
   ClientLeft      =   4380
   ClientTop       =   2265
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   4905
   ScaleWidth      =   9960
   Begin VB.CommandButton Command2 
      Caption         =   "解密转换文本"
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox KeyText 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   2760
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "加密信息"
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   9735
      Begin VB.TextBox CodeText 
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   240
         Width           =   9495
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密转换（HEX）"
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "原始信息"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox HexText 
         Height          =   1455
         Left            =   6000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   240
         Width           =   3615
      End
      Begin VB.TextBox SourceText 
         Height          =   1455
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   240
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "H E X 输入"
         Height          =   975
         Left            =   5760
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "文本输入"
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
   End
   Begin VB.Label Label1 
      Caption         =   "您的密钥"
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim keyByte() As Byte '密钥序列
Dim codeByte() As Byte '待加密序列
'the plan for stream code
'1,get the ascii code of all the string
'2.use 10 to chu the number or 100 to chu the number,to make it less
'3.use it as a seed to creat a codeByte
'4.combine then to creat a string
'创建一个函数，将源字节流处理后得到一组秘钥
Function CreatKey(ByRef source() As Byte) As Byte()
    Dim r() As Byte
    Dim i As Long, bLength As Long
    Dim tFloat As Single
    bLength = UBound(source)
    ReDim r(bLength)
    For i = 0 To bLength
        tFloat = source(i) / 255
        Randomize
        r(i) = CByte(Rnd(tFloat) * 255)
    Next i
    CreatKey = r
End Function

'将生成的key与source异或编码【代码有问题！】
Function XorGroup(ByRef source() As Byte, ByRef key() As Byte) As Byte()
    Dim r() As Byte
    Dim i As Long
    Dim length As Long
    If UBound(key) <= UBound(source) Then
        length = UBound(key)
    Else
        length = UBound(source)
    End If
    ReDim r(length)
    For i = 0 To length
        r(i) = source(i) Xor key(i)
    Next i
    XorGroup = r
End Function
'将字符密码替换为字节密码，填充到字节数组
'2014年9月18日08:50:00 孙瑞修改，大改
Function TransKeyString(ByVal strKey As String, ByRef org() As Byte)
    org = strKey
End Function
'将内容字节数组大小调整为秘钥长度的整数倍
'2014年9月18日08:53:16 孙瑞调整
Function ReArrange(ByRef contentByte() As Byte, ByVal keylen As Long) As Long
    Dim t As Single
    Dim targetSize As Long
    t = (UBound(contentByte) + 1) / (keylen)
    If cint() Then
        '整数
        Exit Function
    Else
        targetSize = (Int(t) + 1) * keylen - 1
    End If
    ReDim Preserve contentByte(targetSize)
    'if (ubound(contentbyte)+1)/keylen
End Function

Sub GetSection(ByRef sByte() As Byte, ByRef rt() As Byte, ByVal BeGins As Long, ByVal klen As Long)
    'sByte  :数据来源 ，rt：返回数据  klen：长度
    Dim i As Long
    ReDim rt(klen - 1)
    For i = 0 To klen - 1
        rt(i) = sByte(BeGins + i)
    Next i
End Sub

Sub SaveToSection(ByRef source() As Byte, ByRef sdata() As Byte, ByVal loc As Long)
    Dim i As Long
    For i = 0 To UBound(sdata)
        source(loc + i) = sdata(i)
    Next i
End Sub

Function FixIt(ByVal s As String) As String
    If Len(s) < 2 Then
        s = s & "0"
    End If
    FixIt = s
End Function

Private Sub Command1_Click()
    'step one
    'make up for empty
    '定义变量区
    Dim sourStr As String
    Dim showStr As String
    Dim KeyStr As String
    Dim keyByte() As Byte
    Dim countByte() As Byte
    Dim tempByte() As Byte
    Dim i As Long
    
    Dim keylen As Long
    Dim contentlen As Long
    '代码区

    sourStr = SourceText.Text       '原始文本
    KeyStr = KeyText.Text           '密码文本
    TransKeyString KeyStr, keyByte  '将字符密码替换成字节密码【字符密码初值】
    keylen = UBound(keyByte) + 1    '密码字节数组长度
    countByte = sourStr '获取待加密内容的字节数据
    
    Debug.Print "调整前数组长度：" & UBound(countByte) + 1
    ReArrange countByte, keylen '数据重整
    contentlen = UBound(countByte) + 1
    Debug.Print "调整后数组长度：" & contentlen
    Debug.Print "调整后数组内容：" & ShowBytes(countByte)
    
    ReDim tempByte(UBound(keyByte))
    '测试代码区
    '循环提取数据
    For i = 1 To contentlen / keylen
        GetSection countByte, tempByte, (i - 1) * keylen, keylen
        keyByte = CreatKey(keyByte)
        tempByte = XorGroup(tempByte, keyByte)
        SaveToSection countByte, tempByte, (i - 1) * keylen
    Next i
    '展示数据
    For i = 0 To UBound(countByte)
        If i Mod 20 = False Then showStr = showStr & vbCrLf
        showStr = showStr & " " & FixIt(Hex(countByte(i)))
    Next i
    CodeText.Text = showStr
End Sub

Private Sub Command2_Click()
    'step one
    'make up for empty
    '定义变量区
    Dim sourStr As String
    Dim midStr() As String
    Dim showStr As String
    Dim KeyStr As String
    Dim keyByte() As Byte
    Dim countByte() As Byte
    Dim tempByte() As Byte
    Dim i As Long
    
    Dim keylen As Long
    Dim contentlen As Long
    '代码区
    '获取原始数据，直接转换为字节数组
    sourStr = HexText.Text       '原始文本
    sourStr = Trim(sourStr)      '清除空格
    sourStr = Replace(sourStr, vbCrLf, "")
    midStr = Split(sourStr, " ") '根据空格分割
    Call GetBytes(countByte, midStr)
    Debug.Print ShowBytes(countByte) '到这一步都没问题
    '*************
    KeyStr = KeyText.Text           '密码文本
    TransKeyString KeyStr, keyByte  '将字符密码替换成字节密码【字符密码初值】
    keylen = UBound(keyByte) + 1    '密码字节数组长度
    Debug.Print "调整前数组长度：" & UBound(countByte) + 1
    ReArrange countByte, keylen '数据重整
    contentlen = UBound(countByte) + 1
    Debug.Print "调整后数组长度：" & contentlen
    
    ReDim tempByte(UBound(keyByte))
    '测试代码区
    '循环提取数据
    For i = 1 To contentlen / keylen
        GetSection countByte, tempByte, (i - 1) * keylen, keylen
        keyByte = CreatKey(keyByte)
        tempByte = XorGroup(tempByte, keyByte)
        SaveToSection countByte, tempByte, (i - 1) * keylen
    Next i
    '展示数据
    Debug.Print "调整后解码的内容：" & ShowBytes(countByte)
    showStr = countByte
    CodeText.Text = showStr
End Sub

Function ShowBytes(ByRef SourceBytes() As Byte) As String
    Dim strOut As String
    Dim i As Long
    For i = 0 To UBound(SourceBytes)
        If strOut = "" Then
            strOut = strOut & FixIt(Hex(SourceBytes(i)))
        Else
            strOut = strOut & " " & FixIt(Hex(SourceBytes(i)))
        End If
    Next i
    ShowBytes = strOut
End Function
