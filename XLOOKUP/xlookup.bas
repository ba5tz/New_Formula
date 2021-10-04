Function XLOOKUP(rng1 As Variant, rng2 As Range, rng3 As Variant, Optional arg1 As Variant, Optional arg2 As Variant) As Range
Rem Auth : ExcelNoob.com
If IsMissing(arg1) Then arg1 = 0
If IsMissing(arg2) Then arg2 = 0
Dim rsult As Variant 'Untuk Hasil array Akhir
Dim r2width As Integer: r2width = rng2.Columns.Count
Dim r3width As Integer: r3width = rng3.Columns.Count
Dim rtnHeaderColumn As Boolean: rtnHeaderColumn = r2width > 1
If r2width > 1 And r2width <> r3width Then
   XLOOKUP = CVErr(xlErrRef)
   Exit Function
End If
Dim srchVal As Variant: srchVal = rng1.Value 'Nilai yg dicari'
Dim sIndex As Double: sIndex = rng2.Row - 1 
Dim n As Long 'for array loop
If (arg1 <> 2 And VarType(rng1) = vbString) Then srchVal = Replace(Replace(Replace(srchVal, "*", "~*"), "?", "~?"), "#", "~#") 'untuk wildcard switch'
'-----------------------'
Dim srchType As String
Dim matchArg As Integer
Dim lDirection As String
Dim nextSize As String
Select Case arg1 
    Case 0, 2
        If arg2 = 0 Or arg2 = 1 Then
            srchType = "im"
            matchArg = 0
        End If
    Case 1, -1
        nextSize = IIf(arg1 = -1, "s", "l") 
        If arg2 = 0 Or arg2 = 1 Then
            srchType = "lp"
            lDirection = "forward"
        End If
End Select
Select Case arg2 
    Case -1
        srchType = "lp": lDirection = "reverse"
    Case 2
        srchType = "im": matchArg = 1
    Case -2
        srchType = "im": matchArg = -1
End Select
If srchType = "im" Then 
    If rtnHeaderColumn Then
        Set XLOOKUP = rng3.Columns(WorksheetFunction.Match(srchVal, rng2, matchArg))
    Else
        Set XLOOKUP = rng3.Rows(WorksheetFunction.Match(srchVal, rng2, matchArg))
    End If
    Exit Function
Else  
    Dim vArr As Variant: vArr = IIf(rtnHeaderColumn, WorksheetFunction.Transpose(rng2), rng2) 
    Dim nsml As Variant: ' nsmal - next smallest value
    Dim nlrg As Variant: ' nlrg - next largest value
    Dim nStart As Double: nStart = IIf(lDirection = "forward", 1, UBound(vArr))
    Dim nEnd As Double: nEnd = IIf(lDirection = "forward", UBound(vArr), 1)
    Dim nStep As Integer: nStep = IIf(lDirection = "forward", 1, -1)
        For n = nStart To nEnd Step nStep
            If vArr(n, 1) Like srchVal Then Set XLOOKUP = IIf(rtnHeaderColumn, rng3.Columns(n), rng3.Rows(n)): Exit Function 
            If nsml < vArr(n, 1) And vArr(n, 1) < srchVal Then 
                Set nsml = rng2.Rows(n)
            End If
            If vArr(n, 1) > srchVal And (IsEmpty(nlrg) Or nlrg > vArr(n, 1)) Then 
                Set nlrg = IIf(rtnHeaderColumn, rng2.Columns(n), rng2.Rows(n))
            End If
        Next
End If
If arg1 = -1 Then 
    Set XLOOKUP = rng3.Rows(nsml.Row - sIndex)
ElseIf arg1 = 1 Then 
    Set XLOOKUP = rng3.Rows(nlrg.Row - sIndex)
End If
End Function
