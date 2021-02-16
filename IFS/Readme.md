

Fungsi IFS memeriksa apakah satu atau beberapa kondisi terpenuhi dan mengembalikan nilai yang sesuai dengan kondisi TRUE pertama. 
IFS dapat menggantikan beberapa pernyataan IF yang bertumpuk, dan jauh lebih mudah dibaca dengan beberapa kondisi.


Secara umum, sintaks fungsi IFS adalah:
```
=IFS([Logika test1, Nilai jika Benar1, Logika test2, Nilai jika Benar2, Logika test3, Nilai jika Benar3)

```

## Script
Module1.Bas
```
Function IFS(ParamArray arguments() As Variant)
REM Auth : ExcelNoob.com
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
```

