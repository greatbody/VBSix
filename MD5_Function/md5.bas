Attribute VB_Name = "md5"
Option Explicit
'�����������У���䵽����MD5Ҫ��
Public Function FillBin(ByRef source() As Byte) As Byte()
    Dim lngCodeLen As Long
    lngCodeLen = UBound(source) + 1 '��ȡ��ǰ�ĳ���
    
End Function
'����һ�����־���ĳ�����ֵı����������
Public Function CalcMiss(ByVal lngNum As Long, ByVal lngUnitNum As Long) As Long
    'lngNum : the number we have
    'lngUnit: the unit number we have
    'example:lngNum=7 lngUnit=2,we need a number that is n Times the value of lngUnit
    Dim fDiv As Double, lngDiv As Long
    If lngNum = lngUnitNum Then
        CalcMiss = 0
        Exit Function
    End If
    fDiv = lngNum / lngUnitNum
    lngDiv = Int(fDiv)
    If IsZs(fDiv) Then
        CalcMiss = 0
    Else
        CalcMiss = (lngDiv + 1) * lngUnitNum - lngNum
    End If
End Function

Public Function IsZs(ByRef value As Variant) As Boolean
    If IsNumeric(value) = False Then
        Err.Raise 11, , "ֵ��������"
    End If
    Dim dblValue As Double
    Dim intValue As Long
    dblValue = CDbl(value)
    intValue = Int(value)
    If dblValue - intValue = 0 Then
        IsZs = True
    Else
        IsZs = False
    End If
End Function

Public Function MakeByteFromHaxStr(ByVal strHexCode As String) As Byte()
    'step one:verify string
    Dim i As Long '��������ѭ���ı���
    Dim lngLength As Long '������¼���ȵı���
    Dim intAsc As Integer
    Dim ret() As Byte   '�����õ�����
    lngLength = Len(strHexCode)
    For i = 1 To lngLength
        intAsc = Asc(Mid(strHexCode, i, 1))
        If ValueIn(intAsc, 48, 57) Or ValueIn(intAsc, 65, 70) Or ValueIn(intAsc, 97, 102) Then
        Else
            Err.Raise 12, , "���ṩ���ַ�����ʮ�������ַ���"
        End If
    Next i
    '��ʼ�ϳ�
    '��Ȳ���������ʮ�������ַ�Ϊ1���ֽ�
    ReDim ret(Int(lngLength / 2) + 1)
End Function

Public Function ValueIn(ByVal value As Long, ByVal lngMin As Long, ByVal lngMax As Long) As Boolean
    If value >= lngMin And value <= lngMax Then
        ValueIn = True
    Else
        ValueIn = False
    End If
End Function

