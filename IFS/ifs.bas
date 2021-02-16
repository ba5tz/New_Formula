Function IFS(ParamArray arguments() As Variant)
REM Support : ExcelNoob.com
Dim j As Integer: j = UBound(arguments)
Dim c As Integer: c = 1
Dim k As Integer: k = (j + 1) / 2
Dim a As Integer
 
If (j + 1) Mod 2 = 1 Then
    IFS = CVErr(xlErrValue)
    Exit Function  
End If
 
For a = 1 To k
    If arguments(c - 1) Then
        IFS = arguments(c)
        Exit Function
    End If
    c = c + 2
Next a
IFS = CVErr(xlErrNA)
End Function
