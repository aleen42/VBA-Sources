Sub 宏H()
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
With Sheets("出货记录")
   Call 宏2
     arr = .Range("a2:p" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And .Range("g" & i) = Sheets("签收单").Range("j" & 3) And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And .Range("g" & i) = Sheets("签收单").Range("j" & 3) And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then
        n = i - 1
        Exit For
        End If
    Next
End With
With Sheets("签收单")
Sheets("签收单").Select
Columns("b:g").Select
Selection.Cells.Clear
Sheets("签收单单头").Range("a1:f5").Copy .Range("b1:g5")
num = 6
.Range("c" & 2) = Sheets("出货记录").Range("f" & n + 1)
.Range("c" & 3) = Sheets("出货记录").Cells(n + 1, "e").Value & " " & Sheets("出货记录").Cells(n + 1, "h").Value
.Range("e" & 2) = Sheets("出货记录").Range("i" & n + 1)
.Range("g" & 2) = .Range("j" & 2)
.Range("G2").Select
Selection.NumberFormatLocal = "yyyy/m/d"
For i = n To n + j - 1
    .Range(Cells(num + i - n, "b"), Cells(num + i - n, "c")).Select
    Selection.Merge
    .Range("b" & num + i - n) = Sheets("出货记录").Range("j" & i + 1)
    .Range("d" & num + i - n) = Sheets("出货记录").Range("k" & i + 1)
    .Range("e" & num + i - n) = Sheets("出货记录").Range("l" & i + 1)
    Next
.Range(Cells(num, "f"), Cells(num + j - 1, "f")).Select
Selection.Merge
.Range(Cells(num, "g"), Cells(num + j - 1, "g")).Select
Selection.Merge
.Range("f" & num) = Sheets("出货记录").Range("m" & n + 1)
.Range("g" & num) = Sheets("出货记录").Range("n" & n + 1)
.Range("b2:g" & .[b65536].End(3).Row).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
Sheets("签收单单尾").Select
    Range("A1:F4").Select
    Selection.Copy
Sheets("签收单").Select
.Range(Cells(num + j, "b"), Cells(num + j + 4, "g")).Select
ActiveSheet.Paste
.Range("g" & .[b65536].End(3).Row) = Format(.Range("j" & 2), "yyyy") + Format(.Range("j" & 2), "mm") + Format(.Range("j" & 2), "dd")

Selection.NumberFormatLocal = "@"

.Range("g" & .[b65536].End(3).Row) = .Range("g" & .[b65536].End(3).Row) & Right(Sheets("出货记录（仙人掌）").Range("g" & n + 1), 2)
.Range("A" & 1).Select
.Range("A" & 1).Select
End With
End Sub

Sub 宏i()
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
With Sheets("出货记录（仙人掌）")
   Call 宏9
     arr = .Range("a2:p" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And Left(.Range("g" & i), 1) = "自" And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And Left(.Range("g" & i), 1) = "自" And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then
        n = i - 1
        Exit For
        End If
    Next
End With
With Sheets("签收单")
Sheets("签收单").Select
Columns("b:g").Select
Selection.Cells.Clear
Sheets("签收单单头").Range("a1:f5").Copy .Range("b1:g5")
num = 6
.Range("c" & 2) = Sheets("出货记录（仙人掌）").Range("f" & n + 1)
.Range("c" & 3) = Sheets("出货记录（仙人掌）").Cells(n + 1, "e").Value & " " & Sheets("出货记录（仙人掌）").Cells(n + 1, "h").Value
.Range("e" & 2) = Sheets("出货记录（仙人掌）").Range("i" & n + 1)
.Range("g" & 2) = .Range("j" & 2)

.Range("G2").Select
Selection.NumberFormatLocal = "yyyy/m/d"
For i = n To n + j - 1
    .Range(Cells(num + i - n, "b"), Cells(num + i - n, "c")).Select
    Selection.Merge
    .Range("b" & num + i - n) = Sheets("出货记录（仙人掌）").Range("j" & i + 1)
    .Range("d" & num + i - n) = Sheets("出货记录（仙人掌）").Range("k" & i + 1)
    .Range("e" & num + i - n) = Sheets("出货记录（仙人掌）").Range("l" & i + 1)
    Next
.Range(Cells(num, "f"), Cells(num + j - 1, "f")).Select
Selection.Merge
.Range(Cells(num, "g"), Cells(num + j - 1, "g")).Select
Selection.Merge
.Range("f" & num) = Sheets("出货记录（仙人掌）").Range("m" & n + 1)
.Range("g" & num) = Sheets("出货记录（仙人掌）").Range("n" & n + 1)


.Range("b2:g" & .[b65536].End(3).Row).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone

With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
Sheets("签收单单尾").Select
    Range("A1:F4").Select
    Selection.Copy
Sheets("签收单").Select
.Range(Cells(num + j, "b"), Cells(num + j + 4, "g")).Select



ActiveSheet.Paste
.Range("g" & .[b65536].End(3).Row) = Format(.Range("j" & 2), "yyyy") + Format(.Range("j" & 2), "mm") + Format(.Range("j" & 2), "dd")

Selection.NumberFormatLocal = "@"

.Range("g" & .[b65536].End(3).Row) = .Range("g" & .[b65536].End(3).Row) & Right(Sheets("出货记录（仙人掌）").Range("g" & n + 1), 2)
.Range("A" & 1).Select
End With
End Sub

Sub 宏j()
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
With Sheets("出货记录（季节风）")
   Call 宏9
     arr = .Range("a2:p" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And .Range("g" & i) = Sheets("签收单").Range("j" & 3) And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("签收单").Range("j" & 2) And .Range("g" & i) = Sheets("签收单").Range("j" & 3) And .Range("f" & i) = Sheets("签收单").Range("j" & 4) Then
        n = i - 1
        Exit For
        End If
    Next
End With
With Sheets("签收单")
Sheets("签收单").Select
Columns("b:g").Select
Selection.Cells.Clear
Sheets("签收单单头").Range("a1:f5").Copy .Range("b1:g5")
num = 6
.Range("c" & 2) = Sheets("出货记录（季节风）").Range("f" & n + 1)
.Range("c" & 3) = Sheets("出货记录（季节风）").Cells(n + 1, "e").Value & " " & Sheets("出货记录（季节风）").Cells(n + 1, "h").Value
.Range("e" & 2) = Sheets("出货记录（季节风）").Range("i" & n + 1)
.Range("g" & 2) = .Range("j" & 2)
.Range("G2").Select
Selection.NumberFormatLocal = "yyyy/m/d"
For i = n To n + j - 1
    .Range(Cells(num + i - n, "b"), Cells(num + i - n, "c")).Select
    Selection.Merge
    .Range("b" & num + i - n) = Sheets("出货记录（季节风）").Range("j" & i + 1)
    .Range("d" & num + i - n) = Sheets("出货记录（季节风）").Range("k" & i + 1)
    .Range("e" & num + i - n) = Sheets("出货记录（季节风）").Range("l" & i + 1)
    Next
.Range(Cells(num, "f"), Cells(num + j - 1, "f")).Select
Selection.Merge
.Range(Cells(num, "g"), Cells(num + j - 1, "g")).Select
Selection.Merge
.Range("f" & num) = Sheets("出货记录（季节风）").Range("m" & n + 1)
.Range("g" & num) = Sheets("出货记录（季节风）").Range("n" & n + 1)
.Range("b2:g" & .[b65536].End(3).Row).Select
Selection.Borders(xlDiagonalDown).LineStyle = xlNone
Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
Sheets("签收单单尾").Select
    Range("A1:F4").Select
    Selection.Copy
Sheets("签收单").Select
.Range(Cells(num + j, "b"), Cells(num + j + 4, "g")).Select
ActiveSheet.Paste
.Range("g" & .[b65536].End(3).Row) = Format(.Range("j" & 2), "yyyy") + Format(.Range("j" & 2), "mm") + Format(.Range("j" & 2), "dd")

Selection.NumberFormatLocal = "@"

.Range("g" & .[b65536].End(3).Row) = .Range("g" & .[b65536].End(3).Row) & Right(Sheets("出货记录（仙人掌）").Range("g" & n + 1), 2)
.Range("A" & 1).Select
.Range("A" & 1).Select
End With
End Sub