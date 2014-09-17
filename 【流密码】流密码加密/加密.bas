Attribute VB_Name = "º”√‹"
Sub GetBytes(ByRef bufByte() As Byte, ByRef strIn() As String)
'
Dim k As Long
ReDim bufByte(k)
For i = 0 To UBound(strIn)
    ReDim Preserve bufByte(k)
    If Trim(strIn(i)) <> "" Then
        bufByte(k) = CByte(HEX_to_DEC(MyTrim(strIn(i))))
        k = k + 1
    End If
Next i
End Sub
Function HEX_to_DEC(ByVal Hex As String) As Long
    Dim i As Long
    Dim sums As Long
    Hex = UCase(Hex)
    For i = 1 To Len(Hex)
        Select Case Mid(Hex, Len(Hex) - i + 1, 1)
            Case "0": sums = sums + 16 ^ (i - 1) * 0
            Case "1": sums = sums + 16 ^ (i - 1) * 1
            Case "2": sums = sums + 16 ^ (i - 1) * 2
            Case "3": sums = sums + 16 ^ (i - 1) * 3
            Case "4": sums = sums + 16 ^ (i - 1) * 4
            Case "5": sums = sums + 16 ^ (i - 1) * 5
            Case "6": sums = sums + 16 ^ (i - 1) * 6
            Case "7": sums = sums + 16 ^ (i - 1) * 7
            Case "8": sums = sums + 16 ^ (i - 1) * 8
            Case "9": sums = sums + 16 ^ (i - 1) * 9
            Case "A": sums = sums + 16 ^ (i - 1) * 10
            Case "B": sums = sums + 16 ^ (i - 1) * 11
            Case "C": sums = sums + 16 ^ (i - 1) * 12
            Case "D": sums = sums + 16 ^ (i - 1) * 13
            Case "E": sums = sums + 16 ^ (i - 1) * 14
            Case "F": sums = sums + 16 ^ (i - 1) * 15
        End Select
    Next i
    HEX_to_DEC = sums
End Function
Function MyTrim(ByVal s As String)
    Dim i As Long, vaLues As Long
    Dim iStart As Long, iEnd As Long
    Dim outS As String
    For i = 1 To Len(s)
        vaLues = Asc(Mid(s, i, 1))
        If vaLues = 32 Or vaLues = 13 Or vaLues = 10 Then
        Else
            iStart = i
            Exit For
        End If
    Next i
    
    For i = Len(s) To 1 Step -1
        vaLues = Asc(Mid(s, i, 1))
        If vaLues = 32 Or vaLues = 13 Or vaLues = 10 Then
        Else
            iEnd = i
            Exit For
        End If
    Next i
    
    If iStart < iEnd Then
        outS = Mid(s, iStart, iEnd - iStart + 1)
    Else
        outS = ""
    End If
    MyTrim = outS
End Function
