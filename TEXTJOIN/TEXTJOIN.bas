Function TEXTJOIN(delim As String, skipblank As Boolean, arr)
    Dim d As Long
    Dim c As Long
    Dim arr2()
    Dim t As Long, y As Long
    t = -1
    y = -1
    If TypeName(arr) = "Range" Then
        arr2 = arr.Value
    Else
        arr2 = arr
    End If
    On Error Resume Next
    t = UBound(arr2, 2)
    y = UBound(arr2, 1)
    On Error GoTo 0

    If t >= 0 And y >= 0 Then
        For c = LBound(arr2, 1) To UBound(arr2, 1)
            For d = LBound(arr2, 1) To UBound(arr2, 2)
                If arr2(c, d) <> "" Or Not skipblank Then
                    TEXTJOIN = TEXTJOIN & arr2(c, d) & delim
                End If
            Next d
        Next c
    Else
        For c = LBound(arr2) To UBound(arr2)
            If arr2(c) <> "" Or Not skipblank Then
                TEXTJOIN = TEXTJOIN & arr2(c) & delim
            End If
        Next c
    End If
    TEXTJOIN = Left(TEXTJOIN, Len(TEXTJOIN) - Len(delim))
End Function

