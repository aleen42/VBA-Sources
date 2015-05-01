Sub 宏a()
Application.ScreenUpdating = False
Call 宏16
Call 宏4
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
Call 按钮2_Click
With Sheets("出货记录")
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n = i - 1
        Exit For
        End If
    Next
    
End With
Sheets("配货单").Select
num = 1
With Sheets("配货单")
.Range("m" & 1) = Sheets("配货单").Range("p" & 1)
For i = n To n + j - 1
    If i = n Then LRow = 3: Str = "": a = arr(i, 14) Else LRow = .[c65536].End(3).Row + a: LRow1 = .[g65536].End(3).Row: Str = .Range("m" & 1)
    If arr(i, 1) <> Str Or k = a Then
    
        For b = n + LRow - 2 To n + LRow + arr(i, 14) - 3
        If Sheets("出货记录").Cells(b, "g") = Sheets("出货记录").Cells(b + 1, "g") And b <> n + LRow + arr(i, 14) - 3 Then
        Else: Exit For
        End If
        Next
        If b = n + LRow + arr(i, 14) - 3 Then a = arr(i, 14) Else a = 1
        
       .Range(Cells(LRow, "a"), Cells(LRow + a - 1, "a")).Select
       
       Selection.Merge
       Selection = num
       num = num + 1
       
       .Range(Cells(LRow, "b"), Cells(LRow + a - 1, "b")).Select
        Selection.Merge
      
       .Range(Cells(LRow, "c"), Cells(LRow + a - 1, "c")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "d"), Cells(LRow + a - 1, "d")).Select
       Selection.Merge
       
       
       For b = n + LRow - 2 To n + LRow + arr(i, 14) - 3
        If Sheets("出货记录").Cells(b, "e") = Sheets("出货记录").Cells(b + 1, "e") And Sheets("出货记录").Cells(b, "f") = Sheets("出货记录").Cells(b + 1, "f") Then
        Else: Exit For
        End If
        Next
       
       .Range(Cells(LRow, "e"), Cells(LRow + a - 1, "e")).Select
        If b = n + LRow + arr(i, 14) - 3 Then Selection.Merge
        
       .Range(Cells(LRow, "f"), Cells(LRow + a - 1, "f")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "j"), Cells(LRow + a - 1, "j")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "k"), Cells(LRow + a - 1, "k")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "l"), Cells(LRow + a - 1, "l")).Select
       If arr(i, 16) = "到付" Then
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       Selection.Merge
       
       .Range(Cells(LRow, "m"), Cells(LRow + a - 1, "m")).Select
       Selection.Merge
       
       .Range("b" & LRow) = arr(i, 7)
       .Range("c" & LRow) = arr(i, 6)
       .Range("d" & LRow) = arr(i, 9)
       .Range("e" & LRow) = arr(i, 5)
       .Range("f" & LRow) = arr(i, 8)
       If InStr(1, .Range("f" & LRow), "送") > 0 And InStr(1, .Range("f" & LRow), "货") > 0 And InStr(1, .Range("f" & LRow), "上") > 0 Then '设置送货上门颜色
            With .Range("f" & LRow).Characters(start:=InStr(1, .Range("f" & LRow), "送"), length:=4).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       .Range("g" & LRow) = arr(i, 10)
       .Range("h" & LRow) = arr(i, 11)
       .Range("i" & LRow) = arr(i, 12)
       .Range("j" & LRow) = arr(i, 13)
       .Range("k" & LRow) = arr(i, 14)
       .Range("l" & LRow) = arr(i, 15) & arr(i, 16)
       .Range("m" & LRow) = arr(i, 17)
       k = 1
    Else
       k = k + 1
       .Range("b" & LRow1 + 1) = arr(i, 7)
       .Range("e" & LRow1 + 1) = arr(i, 5)
       .Range("g" & LRow1 + 1) = arr(i, 10)
       .Range("h" & LRow1 + 1) = arr(i, 11)
       .Range("i" & LRow1 + 1) = arr(i, 12)
       a = arr(i, 14)
    End If
Next

.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("g2:g" & .[g65536].End(3).Row).Select
With Selection
    .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
End With
.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Rows.AutoFit
End With

Call 宏3

.Range(Cells(1, "a"), Cells(1, "m")).Select
Selection.RowHeight = 42
.Range("a" & 1).Select
End With



Application.ScreenUpdating = True
Sheets("配货单").Activate

    
End Sub




Sub 宏2()
    Sheets("出货记录").Select
    Sheets("出货记录").Select
    Columns("A:T").Select
    ActiveWorkbook.Worksheets("出货记录").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("出货记录").Sort.SortFields.Add Key:=Range( _
        "A2:A4680"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录").Sort.SortFields.Add Key:=Range( _
        "G2:G4680"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录").Sort.SortFields.Add Key:=Range( _
        "H2:H4680"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录").Sort.SortFields.Add Key:=Range( _
        "F2:F4680"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("出货记录").Sort
        .SetRange Range("A1:T4680")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub


Sub 宏3()
    With Sheets("配货单")
    Range("A2:M" & .[a65536].End(3).Row).Select
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
    End With
    
End Sub

Sub 宏4()
    Sheets("配货单").Select
    With Sheets("配货单")
    Sheets("配货单表头").Range("a1:m2").Copy .Range("a" & .[g65536].End(3).Row)
    End With
End Sub
Sub 宏5()
    Sheets("配货单").Select
    With Sheets("配货单")
    Sheets("配货单表头1").Range("a1:m2").Copy .Range("a" & .[g65536].End(3).Row)
    End With
End Sub
Sub 宏6()
    Sheets("配货单").Select
    With Sheets("配货单")
    Sheets("配货单表头2").Range("a1:m2").Copy .Range("a" & .[g65536].End(3).Row)
    End With
End Sub

Sub 宏25()
    Sheets("配货单").Select
    With Sheets("配货单")
    Sheets("配货单表头1").Range("a1:m2").Copy .Range("a" & .[g65536].End(3).Row + 1)
    End With
End Sub
Sub 宏26()
    Sheets("配货单").Select
    With Sheets("配货单")
    Sheets("配货单表头2").Range("a1:m2").Copy .Range("a" & .[g65536].End(3).Row + 1)
    End With
End Sub
Sub 宏16()
    Sheets("配货单").Select
    With Sheets("配货单")
    Columns("a:M").Select
    Selection.Cells.Clear
    .Range("a:m").Select
    With Selection
        Selection.RowHeight = 13.5
    End With
End With
End Sub



Sub 宏7()
Call 按钮2_Click

Dim i As Integer, num As Integer
num = 1
Dim j1 As Integer, j2 As Integer, j3 As Integer
Dim n1 As Integer, n2 As Integer, n3 As Integer
With Sheets("出货记录")
   Sheets("出货记录").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j1 = 0
     n1 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            j1 = j1 + 1
            .Range("u" & i) = num
            num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n1 = i - 1
        Exit For
        End If
    Next
End With

With Sheets("出货记录（仙人掌）")
    Sheets("出货记录（仙人掌）").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j2 = 0
     n2 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            j2 = j2 + 1
            .Range("u" & i) = num
            num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            n2 = i - 1
        Exit For
        End If
    Next
End With

With Sheets("出货记录（季节风）")
    Sheets("出货记录（季节风）").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j3 = 0
     n3 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        j3 = j3 + 1
        .Range("u" & i) = num
        num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n3 = i - 1
        Exit For
        End If
    Next
End With
If n1 = 0 And n2 = 0 And n3 = 0 Then
    Call 宏16
End If
If n1 <> 0 And n2 = 0 And n3 = 0 Then
    Call 宏a
End If
If n2 <> 0 And n1 = 0 And n3 = 0 Then
    Call 宏b
End If
If n3 <> 0 And n1 = 0 And n2 = 0 Then
    Call 宏c
End If
If n1 <> 0 And n2 <> 0 And n3 = 0 Then
    Call 宏a
    Call 宏bb
End If
If n1 <> 0 And n3 <> 0 And n2 = 0 Then
    Call 宏a
    Call 宏cc
End If
If n2 <> 0 And n3 <> 0 And n1 = 0 Then
    Call 宏b
    Call 宏cc
End If
If n2 <> 0 And n3 <> 0 And n1 <> 0 Then
    Call 宏a
    Call 宏bb
    Call 宏cc
End If

Columns("M:M").EntireColumn.AutoFit
End Sub

Sub 宏8()

    Call 按钮2_Click

Dim i As Integer, num As Integer
num = 1
Dim j1 As Integer, j2 As Integer, j3 As Integer
Dim n1 As Integer, n2 As Integer, n3 As Integer
With Sheets("出货记录")
   Sheets("出货记录").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j1 = 0
     n1 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            j1 = j1 + 1
            .Range("u" & i) = num
            num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n1 = i - 1
        Exit For
        End If
    Next
End With

With Sheets("出货记录（仙人掌）")
    Sheets("出货记录（仙人掌）").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j2 = 0
     n2 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            j2 = j2 + 1
            .Range("u" & i) = num
            num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
            n2 = i - 1
        Exit For
        End If
    Next
End With

With Sheets("出货记录（季节风）")
    Sheets("出货记录（季节风）").Select
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j3 = 0
     n3 = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        j3 = j3 + 1
        .Range("u" & i) = num
        num = num + 1
        End If
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n3 = i - 1
        Exit For
        End If
    Next
End With
If n1 = 0 And n2 = 0 And n3 = 0 Then
    Call 宏16
End If
If n1 <> 0 And n2 = 0 And n3 = 0 Then
    Call 宏a
    Call 宏table
End If
If n2 <> 0 And n1 = 0 And n3 = 0 Then
    Call 宏b
    Call 宏table
End If
If n3 <> 0 And n1 = 0 And n2 = 0 Then
    Call 宏c
    Call 宏table
End If
If n1 <> 0 And n2 <> 0 And n3 = 0 Then
    Call 宏a
    Call 宏bb
    Call 宏table
End If
If n1 <> 0 And n3 <> 0 And n2 = 0 Then
    Call 宏a
    Call 宏cc
    Call 宏table
End If
If n2 <> 0 And n3 <> 0 And n1 = 0 Then
    Call 宏b
    Call 宏cc
    Call 宏table
End If
If n2 <> 0 And n3 <> 0 And n1 <> 0 Then
    Call 宏a
    Call 宏bb
    Call 宏cc
    Call 宏table
End If

Columns("M:M").EntireColumn.AutoFit


End Sub

Sub 宏b()
Application.ScreenUpdating = False
Call 宏16
Call 宏5
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer

With Sheets("出货记录（仙人掌）")
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n = i - 1
        Exit For
        End If
    Next
    
End With
Sheets("配货单").Select
num = 1
With Sheets("配货单")
Dim sign As Integer

sign = .[g65536].End(3).Row - 1

.Range("m" & sign) = Sheets("配货单").Range("p" & 1)
For i = n To n + j - 1
    If i = n Then LRow = .[c65536].End(3).Row + 1: Str = "": a = arr(i, 14) Else LRow = .[c65536].End(3).Row + a: LRow1 = .[g65536].End(3).Row: Str = .Range("m" & 1)
    If arr(i, 1) <> Str Or k = a Then
    
        For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（仙人掌）").Cells(b, "g") = Sheets("出货记录（仙人掌）").Cells(b + 1, "g") And b <> (n + LRow + arr(i, 14) - 3 - sign + 1) Then
        Else: Exit For
        End If
        Next
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then a = arr(i, 14) Else a = 1
        
       .Range(Cells(LRow, "a"), Cells(LRow + a - 1, "a")).Select
       
       Selection.Merge
       Selection = num
       num = num + 1
       
       .Range(Cells(LRow, "b"), Cells(LRow + a - 1, "b")).Select
        Selection.Merge
      
       .Range(Cells(LRow, "c"), Cells(LRow + a - 1, "c")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "d"), Cells(LRow + a - 1, "d")).Select
       Selection.Merge
       
       
       For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（仙人掌）").Cells(b, "e") = Sheets("出货记录（仙人掌）").Cells(b + 1, "e") And Sheets("出货记录（仙人掌）").Cells(b, "f") = Sheets("出货记录（仙人掌）").Cells(b + 1, "f") Then
        Else: Exit For
        End If
        Next
       
       .Range(Cells(LRow, "e"), Cells(LRow + a - 1, "e")).Select
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then Selection.Merge
        
       .Range(Cells(LRow, "f"), Cells(LRow + a - 1, "f")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "j"), Cells(LRow + a - 1, "j")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "k"), Cells(LRow + a - 1, "k")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "l"), Cells(LRow + a - 1, "l")).Select
       If arr(i, 16) = "到付" Then
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       Selection.Merge
       
       .Range(Cells(LRow, "m"), Cells(LRow + a - 1, "m")).Select
       Selection.Merge
       
       .Range("b" & LRow) = arr(i, 7)
       .Range("c" & LRow) = arr(i, 6)
       .Range("d" & LRow) = arr(i, 9)
       .Range("e" & LRow) = arr(i, 5)
       .Range("f" & LRow) = arr(i, 8)
       If InStr(1, .Range("f" & LRow), "送") > 0 And InStr(1, .Range("f" & LRow), "货") > 0 And InStr(1, .Range("f" & LRow), "上") > 0 Then '设置送货上门颜色
             With .Range("f" & LRow).Characters(start:=InStr(1, .Range("f" & LRow), "送"), length:=4).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       .Range("g" & LRow) = arr(i, 10)
       .Range("h" & LRow) = arr(i, 11)
       .Range("i" & LRow) = arr(i, 12)
       .Range("j" & LRow) = arr(i, 13)
       .Range("k" & LRow) = arr(i, 14)
       .Range("l" & LRow) = arr(i, 15) & arr(i, 16)
       .Range("m" & LRow) = arr(i, 17)
       k = 1
    Else
       k = k + 1
       .Range("b" & LRow1 + 1) = arr(i, 7)
       .Range("e" & LRow1 + 1) = arr(i, 5)
       .Range("g" & LRow1 + 1) = arr(i, 10)
       .Range("h" & LRow1 + 1) = arr(i, 11)
       .Range("i" & LRow1 + 1) = arr(i, 12)
       a = arr(i, 14)
    End If
Next

.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("g2:g" & .[g65536].End(3).Row).Select
With Selection
    .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
End With
.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Rows.AutoFit
End With
Call 宏10

.Range("a" & 1).Select
.Range(Cells(sign, "a"), Cells(sign, "m")).Select
    Selection.RowHeight = 42
End With



Application.ScreenUpdating = True
Sheets("配货单").Activate

    
End Sub

Sub 宏bb()
Application.ScreenUpdating = False
Call 宏25
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer

With Sheets("出货记录（仙人掌）")
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n = i - 1
        Exit For
        End If
    Next
    
End With
Sheets("配货单").Select
num = 1
With Sheets("配货单")
Dim sign As Integer

sign = .[g65536].End(3).Row - 1

.Range("m" & sign) = Sheets("配货单").Range("p" & 1)
For i = n To n + j - 1
    If i = n Then LRow = .[c65536].End(3).Row + 1: Str = "": a = arr(i, 14) Else LRow = .[c65536].End(3).Row + a: LRow1 = .[g65536].End(3).Row: Str = .Range("m" & 1)
    If arr(i, 1) <> Str Or k = a Then
    
        For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（仙人掌）").Cells(b, "g") = Sheets("出货记录（仙人掌）").Cells(b + 1, "g") And b <> (n + LRow + arr(i, 14) - 3 - sign + 1) Then
        Else: Exit For
        End If
        Next
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then a = arr(i, 14) Else a = 1
        
       .Range(Cells(LRow, "a"), Cells(LRow + a - 1, "a")).Select
       
       Selection.Merge
       Selection = num
       num = num + 1
       
       .Range(Cells(LRow, "b"), Cells(LRow + a - 1, "b")).Select
        Selection.Merge
      
       .Range(Cells(LRow, "c"), Cells(LRow + a - 1, "c")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "d"), Cells(LRow + a - 1, "d")).Select
       Selection.Merge
       
       
       For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（仙人掌）").Cells(b, "e") = Sheets("出货记录（仙人掌）").Cells(b + 1, "e") And Sheets("出货记录（仙人掌）").Cells(b, "f") = Sheets("出货记录（仙人掌）").Cells(b + 1, "f") Then
        Else: Exit For
        End If
        Next
       
       .Range(Cells(LRow, "e"), Cells(LRow + a - 1, "e")).Select
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then Selection.Merge
        
       .Range(Cells(LRow, "f"), Cells(LRow + a - 1, "f")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "j"), Cells(LRow + a - 1, "j")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "k"), Cells(LRow + a - 1, "k")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "l"), Cells(LRow + a - 1, "l")).Select
       If arr(i, 16) = "到付" Then
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       Selection.Merge
       
       .Range(Cells(LRow, "m"), Cells(LRow + a - 1, "m")).Select
       Selection.Merge
       
       .Range("b" & LRow) = arr(i, 7)
       .Range("c" & LRow) = arr(i, 6)
       .Range("d" & LRow) = arr(i, 9)
       .Range("e" & LRow) = arr(i, 5)
       .Range("f" & LRow) = arr(i, 8)
       If InStr(1, .Range("f" & LRow), "送") > 0 And InStr(1, .Range("f" & LRow), "货") > 0 And InStr(1, .Range("f" & LRow), "上") > 0 Then '设置送货上门颜色
             With .Range("f" & LRow).Characters(start:=InStr(1, .Range("f" & LRow), "送"), length:=4).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       .Range("g" & LRow) = arr(i, 10)
       .Range("h" & LRow) = arr(i, 11)
       .Range("i" & LRow) = arr(i, 12)
       .Range("j" & LRow) = arr(i, 13)
       .Range("k" & LRow) = arr(i, 14)
       .Range("l" & LRow) = arr(i, 15) & arr(i, 16)
       .Range("m" & LRow) = arr(i, 17)
       k = 1
    Else
       k = k + 1
       .Range("b" & LRow1 + 1) = arr(i, 7)
       .Range("e" & LRow1 + 1) = arr(i, 5)
       .Range("g" & LRow1 + 1) = arr(i, 10)
       .Range("h" & LRow1 + 1) = arr(i, 11)
       .Range("i" & LRow1 + 1) = arr(i, 12)
       a = arr(i, 14)
    End If
Next

.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("g2:g" & .[g65536].End(3).Row).Select
With Selection
        .VerticalAlignment = xlCenter
End With
.Range(Cells(sign, "a"), Cells(.[g65536].End(3).Row, "m")).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Rows.AutoFit
End With
Call 宏10

.Range("a" & 1).Select
.Range(Cells(sign, "a"), Cells(sign, "m")).Select
    Selection.RowHeight = 42
End With



Application.ScreenUpdating = True
Sheets("配货单").Activate

    
End Sub



Sub 宏9()


Sheets("出货记录（仙人掌）").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-30
    Range("A1:T1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort.SortFields.Add Key:=Range( _
        "A2:A420"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort.SortFields.Add Key:=Range( _
        "G2:G420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort.SortFields.Add Key:=Range( _
        "H2:H420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort.SortFields.Add Key:=Range( _
        "F2:F420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("出货记录（仙人掌）").Sort
        .SetRange Range("A1:T420")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Sub 宏10()
    With Sheets("配货单")
    Range("a2:m" & .[c65536].End(3).Row).Select
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
    End With
    
End Sub

Sub 宏cc()

Application.ScreenUpdating = False
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
Call 宏26
With Sheets("出货记录（季节风）")
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n = i - 1
        Exit For
        End If
    Next
    
End With
Sheets("配货单").Select
num = 1
With Sheets("配货单")
Dim sign As Integer
sign = .[g65536].End(3).Row - 1
.Range("m" & sign) = Sheets("配货单").Range("p" & 1)
For i = n To n + j - 1
    If i = n Then LRow = .[c65536].End(3).Row + 1: Str = "": a = arr(i, 14) Else LRow = .[c65536].End(3).Row + a: LRow1 = .[g65536].End(3).Row: Str = .Range("m" & 1)
    If arr(i, 1) <> Str Or k = a Then
    
        For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（季节风）").Cells(b, "g") = Sheets("出货记录（季节风）").Cells(b + 1, "g") And b <> (n + LRow + arr(i, 14) - 3 - sign + 1) Then
        Else: Exit For
        End If
        Next
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then a = arr(i, 14) Else a = 1
        
       .Range(Cells(LRow, "a"), Cells(LRow + a - 1, "a")).Select
       
       Selection.Merge
       Selection = num
       num = num + 1
       
       .Range(Cells(LRow, "b"), Cells(LRow + a - 1, "b")).Select
        Selection.Merge
      
       .Range(Cells(LRow, "c"), Cells(LRow + a - 1, "c")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "d"), Cells(LRow + a - 1, "d")).Select
       Selection.Merge
       
       
       For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（季节风）").Cells(b, "e") = Sheets("出货记录（季节风）").Cells(b + 1, "e") And Sheets("出货记录（季节风）").Cells(b, "f") = Sheets("出货记录（季节风）").Cells(b + 1, "f") Then
        Else: Exit For
        End If
        Next
       
       .Range(Cells(LRow, "e"), Cells(LRow + a - 1, "e")).Select
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then Selection.Merge
        
       .Range(Cells(LRow, "f"), Cells(LRow + a - 1, "f")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "j"), Cells(LRow + a - 1, "j")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "k"), Cells(LRow + a - 1, "k")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "l"), Cells(LRow + a - 1, "l")).Select
       If arr(i, 16) = "到付" Then
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       Selection.Merge
       
       .Range(Cells(LRow, "m"), Cells(LRow + a - 1, "m")).Select
       Selection.Merge
       
       .Range("b" & LRow) = arr(i, 7)
       .Range("c" & LRow) = arr(i, 6)
       .Range("d" & LRow) = arr(i, 9)
       .Range("e" & LRow) = arr(i, 5)
       .Range("f" & LRow) = arr(i, 8)
       If InStr(1, .Range("f" & LRow), "送") > 0 And InStr(1, .Range("f" & LRow), "货") > 0 And InStr(1, .Range("f" & LRow), "上") > 0 Then '设置送货上门颜色
             With .Range("f" & LRow).Characters(start:=InStr(1, .Range("f" & LRow), "送"), length:=4).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       .Range("g" & LRow) = arr(i, 10)
       .Range("h" & LRow) = arr(i, 11)
       .Range("i" & LRow) = arr(i, 12)
       .Range("j" & LRow) = arr(i, 13)
       .Range("k" & LRow) = arr(i, 14)
       .Range("l" & LRow) = arr(i, 15) & arr(i, 16)
       .Range("m" & LRow) = arr(i, 17)
       k = 1
    Else
       k = k + 1
       .Range("b" & LRow1 + 1) = arr(i, 7)
       .Range("e" & LRow1 + 1) = arr(i, 5)
       .Range("g" & LRow1 + 1) = arr(i, 10)
       .Range("h" & LRow1 + 1) = arr(i, 11)
       .Range("i" & LRow1 + 1) = arr(i, 12)
       a = arr(i, 14)
    End If
Next

.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("g2:g" & .[g65536].End(3).Row).Select
With Selection
        .VerticalAlignment = xlCenter
End With
.Range(Cells(sign, "a"), Cells(.[g65536].End(3).Row, "m")).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Rows.AutoFit
End With
Call 宏15

.Range("a" & 1).Select
.Range(Cells(sign, "a"), Cells(sign, "m")).Select
    Selection.RowHeight = 42
End With



Application.ScreenUpdating = True
Sheets("配货单").Activate


End Sub
Sub 宏c()

Application.ScreenUpdating = False
Dim Str As String, i As Integer, j As Integer, k As Integer, n As Integer, a As Integer, b As Integer, num As Integer
Call 宏16
Call 宏6
With Sheets("出货记录（季节风）")
     arr = .Range("a2:q" & .[a65536].End(3).Row)
     j = 0
     n = 0
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j = j + 1
    Next
    
    For i = 2 To UBound(arr) + 1
        If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
        n = i - 1
        Exit For
        End If
    Next
    
End With
Sheets("配货单").Select
num = 1
With Sheets("配货单")
Dim sign As Integer
sign = .[g65536].End(3).Row - 1
.Range("m" & sign) = Sheets("配货单").Range("p" & 1)
For i = n To n + j - 1
    If i = n Then LRow = .[c65536].End(3).Row + 1: Str = "": a = arr(i, 14) Else LRow = .[c65536].End(3).Row + a: LRow1 = .[g65536].End(3).Row: Str = .Range("m" & 1)
    If arr(i, 1) <> Str Or k = a Then
    
        For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（季节风）").Cells(b, "g") = Sheets("出货记录（季节风）").Cells(b + 1, "g") And b <> (n + LRow + arr(i, 14) - 3 - sign + 1) Then
        Else: Exit For
        End If
        Next
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then a = arr(i, 14) Else a = 1
        
       .Range(Cells(LRow, "a"), Cells(LRow + a - 1, "a")).Select
       
       Selection.Merge
       Selection = num
       num = num + 1
       
       .Range(Cells(LRow, "b"), Cells(LRow + a - 1, "b")).Select
        Selection.Merge
      
       .Range(Cells(LRow, "c"), Cells(LRow + a - 1, "c")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "d"), Cells(LRow + a - 1, "d")).Select
       Selection.Merge
       
       
       For b = n + LRow - 2 - sign + 1 To n + LRow + arr(i, 14) - 3 - sign + 1
        If Sheets("出货记录（季节风）").Cells(b, "e") = Sheets("出货记录（季节风）").Cells(b + 1, "e") And Sheets("出货记录（季节风）").Cells(b, "f") = Sheets("出货记录（季节风）").Cells(b + 1, "f") Then
        Else: Exit For
        End If
        Next
       
       .Range(Cells(LRow, "e"), Cells(LRow + a - 1, "e")).Select
        If b = n + LRow + arr(i, 14) - 3 - sign + 1 Then Selection.Merge
        
       .Range(Cells(LRow, "f"), Cells(LRow + a - 1, "f")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "j"), Cells(LRow + a - 1, "j")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "k"), Cells(LRow + a - 1, "k")).Select
       Selection.Merge
       
       .Range(Cells(LRow, "l"), Cells(LRow + a - 1, "l")).Select
       If arr(i, 16) = "到付" Then
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       Selection.Merge
       
       .Range(Cells(LRow, "m"), Cells(LRow + a - 1, "m")).Select
       Selection.Merge
       
       .Range("b" & LRow) = arr(i, 7)
       .Range("c" & LRow) = arr(i, 6)
       .Range("d" & LRow) = arr(i, 9)
       .Range("e" & LRow) = arr(i, 5)
       .Range("f" & LRow) = arr(i, 8)
       If InStr(1, .Range("f" & LRow), "送") > 0 And InStr(1, .Range("f" & LRow), "货") > 0 And InStr(1, .Range("f" & LRow), "上") > 0 Then '设置送货上门颜色
             With .Range("f" & LRow).Characters(start:=InStr(1, .Range("f" & LRow), "送"), length:=4).Font
                .Color = -16776961
                .TintAndShade = 0
            End With
       End If
       .Range("g" & LRow) = arr(i, 10)
       .Range("h" & LRow) = arr(i, 11)
       .Range("i" & LRow) = arr(i, 12)
       .Range("j" & LRow) = arr(i, 13)
       .Range("k" & LRow) = arr(i, 14)
       .Range("l" & LRow) = arr(i, 15) & arr(i, 16)
       .Range("m" & LRow) = arr(i, 17)
       k = 1
    Else
       k = k + 1
       .Range("b" & LRow1 + 1) = arr(i, 7)
       .Range("e" & LRow1 + 1) = arr(i, 5)
       .Range("g" & LRow1 + 1) = arr(i, 10)
       .Range("h" & LRow1 + 1) = arr(i, 11)
       .Range("i" & LRow1 + 1) = arr(i, 12)
       a = arr(i, 14)
    End If
Next

.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("g2:g" & .[g65536].End(3).Row).Select
With Selection
    .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
End With
.Range("a2:m" & .[c65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Rows.AutoFit
End With
Call 宏15

.Range("a" & 1).Select
.Range(Cells(sign, "a"), Cells(sign, "m")).Select
    Selection.RowHeight = 42
End With



Application.ScreenUpdating = True
Sheets("配货单").Activate


End Sub


Sub 宏14()
Sheets("出货记录（季节风）").Select
    Range("A2:T2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-30
    Range("A1:T1").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWorkbook.Worksheets("出货记录（季节风）").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("出货记录（季节风）").Sort.SortFields.Add Key:=Range( _
        "A2:A420"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（季节风）").Sort.SortFields.Add Key:=Range( _
        "G2:G420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（季节风）").Sort.SortFields.Add Key:=Range( _
        "H2:H420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    ActiveWorkbook.Worksheets("出货记录（季节风）").Sort.SortFields.Add Key:=Range( _
        "F2:F420"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("出货记录（季节风）").Sort
        .SetRange Range("A1:T420")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub


Sub 宏15()
     With Sheets("配货单")
    Range("a2:m" & .[c65536].End(3).Row).Select
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
    End With
    
End Sub





Sub 宏abc()

    ActiveWorkbook.Worksheets("出货记录").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("出货记录").AutoFilter.Sort.SortFields.Add Key:=Range( _
        "A2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("出货记录").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub 按钮1_Click()
    Dim l As Integer, m As Integer, o As Integer, p As Integer, num As Integer, i As Integer
    Dim j1 As Integer, j2 As Integer, j3 As Integer, j4 As Integer, j6 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, n4 As Integer, n5 As Integer
    Dim copyFromFileName As String, copyFromFileName1 As String, copyFromFileName2 As String
    Dim myPath As String
    Dim myFile As String
    Dim openFile As Workbook
    
    Dim signal As Boolean
    
    
    Dim endRow As Integer
    Dim endColumn As Integer
    Dim endColumnChar As String
    Dim rang As String
    
    
    copyFromFileName = "落单表-梁子婷.xls"    '这个地方设置被复制的excel文件
    copyFromFileName1 = "落单表-陆伟东.xls"    '这个地方设置被复制的excel文件
    copyFromFileName2 = "落单表-温碧莹.xls"    '这个地方设置被复制的excel文件

    myPath = "D:\用户目录\Documents\落单表" & "/" '把文件路径定义给变量
    myFile = Dir(myPath & "*.xls")   '依次找寻指定路径中的*.xls文件
    
    Do While myFile <> ""
        If myFile = copyFromFileName Or myFile = copyFromFileName1 Or myFile = copyFromFileName2 Then '假如遍历到需要复制的文件
            If myFile = copyFromFileName Then
                Set openFile = Workbooks.Open(myPath & myFile) '打开符合要求的文件
            End If
            If myFile = copyFromFileName1 Then
                Set openFile = Workbooks.Open(myPath & myFile) '打开符合要求的文件
            End If
            If myFile = copyFromFileName2 Then
                Set openFile = Workbooks.Open(myPath & myFile) '打开符合要求的文件
            End If
            'For i = 1 To openFile.Sheets.Count '复制所有的sheet
                'endRow = openFile.Sheets(i).Range("a65536").End(xlUp).Row   '根据第一列来确定有数据的最后一行
                'endColumn = openFile.Sheets(i).Cells(1255).End(xlToLeft).Column '根据第一行来确定有数据的最后一列
               ' endColumnChar = VBA.Split(Columns(endColumn).Address, "$")(2)   '取得最后一列对应的字母
                'rang = "A1:" & endColumnChar & endRow   '构建成标准的范围格式 例：“A1：C1”
                'openFile.Sheets(i).Range(rang).Copy ThisWorkbook.Sheets(i).Range(rang) '实现拷贝
            'Next
            Windows("每天出货信息.xls").Activate
            With Sheets("出货记录")
                Sheets("出货记录").Select
                Call 宏2
                arr3 = .Range("a2:t" & .[a65536].End(3).Row)
                j3 = 0
                n3 = 0
                For i = 2 To UBound(arr3) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j3 = j3 + 1 'j stnds for which is the end of those dates
                Next
                j3 = j3 + 1
                For i = 2 To UBound(arr3) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n3 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
            End With
            
            With Sheets("出货记录（仙人掌）")
                Sheets("出货记录（仙人掌）").Select
                Call 宏2
                arr4 = .Range("a2:t" & .[a65536].End(3).Row)
                j4 = 0
                n4 = 0
                For i = 2 To UBound(arr4) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j4 = j4 + 1 'j stnds for which is the end of those dates
                Next
                j4 = j4 + 1
                For i = 2 To UBound(arr4) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n4 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
            End With
            
             With Sheets("出货记录（季节风）")
                Sheets("出货记录（季节风）").Select
                Call 宏2
                arr5 = .Range("a2:t" & .[a65536].End(3).Row)
                j5 = 0
                n5 = 0
                For i = 2 To UBound(arr5) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j5 = j5 + 1 'j stnds for which is the end of those dates
                Next
                j5 = j5 + 1
                For i = 2 To UBound(arr5) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n5 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
            End With
            
            Windows(myFile).Activate
               Call 宏abc
                With Sheets("出货记录")
                    Sheets("出货记录").Select
                    arr2 = .Range("a2:t" & .[a65536].End(3).Row)
                    For l = 2 To UBound(arr2) + 1
                        If .Range("a" & l) <> ThisWorkbook.Sheets("配货单").Range("p" & 1) Then
                            Goto 100
                        End If
                        
                        signal = True
                        If .Range("s" & l) = "曼途建材" Then
                            If n3 = 0 Then
                                Goto 101
                            End If
                            For o = n3 To j3 + n3 - 2
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录").Range("f" & o) Then
                                    num = 0
                                    For p = n3 To j3 + n3 - 2
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录").Range("f" & p) And .Range("i" & l) = ThisWorkbook.Sheets("出货记录").Range("i" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录").Range("a" & p) Then
                                            num = num + 1
                                        End If
                                    Next
                                    If .Range("n" & l) = num Then
                                        signal = False
                                    End If
                                End If
                            Next
                        ElseIf .Range("s" & l) = "仙人掌" Then
                            If n4 = 0 Then
                                Goto 101
                            End If
                            For o = n4 To j4 + n4 - 2
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("f" & o) Then
                                    num = 0
                                    For p = n4 To j4 + n4 - 2
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("f" & p) And .Range("i" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("i" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("a" & p) Then
                                            num = num + 1
                                        End If
                                    Next
                                    If .Range("n" & l) = num Then
                                        signal = False
                                    End If
                                End If
                            Next
                        ElseIf .Range("s" & l) = "季节风" Then
                            If n5 = 0 Then
                                Goto 101
                            End If
                            For o = n5 To j5 + n5 - 2
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("f" & o) Then
                                    num = 0
                                    For p = n5 To j5 + n5 - 2
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("f" & p) And .Range("i" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("i" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("a" & p) Then
                                            num = num + 1
                                        End If
                                    Next
                                    If .Range("n" & l) = num Then
                                        signal = False
                                    End If
                                End If
                            Next
                        End If
                        
101:                    If .Range("a" & l) = ThisWorkbook.Sheets("配货单").Range("p" & 1) And .Range("s" & l) = "曼途建材" And signal = True Then
                                Windows(myFile).Activate
                                Sheets("出货记录").Select
                                .Range(Cells(l, "a"), Cells(l, "t")).Select
                                Selection.Copy
                                Windows("每天出货信息.xls").Activate
                                Sheets("出货记录").Select
                                Rows("2:2").Select
                                Selection.Insert Shift:=xlDown
                                signal = True
                                Range("e" & 2) = "=B2&" & Chr(34) & " " & Chr(34) & "&C2&" & Chr(34) & " " & Chr(34) & "&D2"
                                n3 = n3 + 1
                                j3 = j3 + 1
                        ElseIf .Range("a" & l) = ThisWorkbook.Sheets("配货单").Range("p" & 1) And .Range("s" & l) = "仙人掌" And signal = True Then
                                Windows(myFile).Activate
                                Sheets("出货记录").Select
                                .Range(Cells(l, "a"), Cells(l, "t")).Select
                                Selection.Copy
                                Windows("每天出货信息.xls").Activate
                                Sheets("出货记录（仙人掌）").Select
                                Rows("2:2").Select
                                Selection.Insert Shift:=xlDown
                                signal = True
                                Range("e" & 2) = "=B2&" & Chr(34) & " " & Chr(34) & "&C2&" & Chr(34) & " " & Chr(34) & "&D2"
                                n4 = n4 + 1
                                j4 = j4 + 1
                        ElseIf .Range("a" & l) = ThisWorkbook.Sheets("配货单").Range("p" & 1) And .Range("s" & l) = "季节风" And signal = True Then
                                 Windows(myFile).Activate
                                Sheets("出货记录").Select
                                .Range(Cells(l, "a"), Cells(l, "t")).Select
                                Selection.Copy
                                Windows("每天出货信息.xls").Activate
                                Sheets("出货记录（季节风）").Select
                                Rows("2:2").Select
                                Selection.Insert Shift:=xlDown
                                signal = True
                                Range("e" & 2) = "=B2&" & Chr(34) & " " & Chr(34) & "&C2&" & Chr(34) & " " & Chr(34) & "&D2"
                                n5 = n5 + 1
                                j5 = j5 + 1
                        End If
100:                    Next
                End With
            Windows(myFile).Activate
            Range("a" & 1).Select
            Windows("每天出货信息.xls").Activate
            Workbooks(myFile).Save
            Workbooks(myFile).Close False         '关闭源工作簿,并不作修改
        End If
        myFile = Dir
    Loop
    Sheets("配货单").Select
    Range("a" & 1).Select
End Sub




Sub 按钮2_Click()

    On Error Goto Error
    
    Dim i As Integer, j1 As Integer, j2 As Integer, j3 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, index As Integer
    Dim D_num As Integer, H_num As Integer, Z_num As Integer, S_num As Integer, F_num, grasp_num As Integer
    Dim lookfromfile As String
    Dim myPath As String
    Dim myFile As String
    Dim openFile As Workbook
    
    Dim temp As Integer
    
    Dim start As Integer
    
    D_num = 1
    H_num = 1
    Z_num = 1
    S_num = 1
    F_num = 1
    grasp_num = 1
    
    lookfromfilename = "物流公司对应表.xls"
    myPath = "E:\马赛克\马赛克" & "/"
    myFile = Dir(myPath & "*.xls")
    Do While myFile <> ""
        If myFile = lookfromfilename Then
            If myFile = lookfromfilename Then
                Set openFile = Workbooks.Open(myPath & myFile)
            End If
            Windows("每天出货信息.xls").Activate
            With Sheets("出货记录")
                Sheets("出货记录").Select
                Call 宏2
                     arr1 = .Range("a2:t" & .[a65536].End(3).Row)
                     j1 = 0
                     n1 = 0
                For i = 2 To UBound(arr1) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j1 = j1 + 1 'j stands for how many in that date
                Next
                
                For i = 2 To UBound(arr1) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n1 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
                j1 = j1 + n1 - 1
                    
                For i = n1 To j1
                    If n1 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "专线" Then
                        Windows(myFile).Activate
                        With Sheets("Sheet1")
                            If ThisWorkbook.Sheets("出货记录").Range("d" & i) <> "" Then
                                For start = 3 To .[d65536].End(3).Row
                                    If .Range("d" & start) = ThisWorkbook.Sheets("出货记录").Range("d" & i) Then
                                        temp = start
                                        Do While .Range("c" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        
                                        If Left(.Range("c" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录").Range("c" & i), 2) Then
                                            temp = start
                                            Do While .Range("a" & temp) = ""
                                                temp = temp - 1
                                            Loop
                                            
                                            If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录").Range("b" & i), 2) Then
                                                If (.Range("e" & start) <> "") Then
                                                    ThisWorkbook.Sheets("出货记录").Range("h" & i) = .Range("e" & start)
                                                End If
                                                Goto 200:
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                For start = 3 To .[c65536].End(3).Row
                                    If Left(.Range("c" & start), 2) = Left(ThisWorkbook.Sheets("出货记录").Range("c" & i), 2) Then
                                        temp = start
                                        Do While .Range("a" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录").Range("b" & i), 2) Then
                                            If (.Range("e" & start) <> "") Then
                                                ThisWorkbook.Sheets("出货记录").Range("h" & i) = .Range("e" & start)
                                            End If
                                            Goto 200:
                                        End If
                                    End If
                                Next
                            End If
                            
                                
200:                    End With
                        Windows("每天出货信息.xls").Activate
                    End If
                Next
                
                Call 宏2
                
                For i = n1 To j1
                    If n1 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "德邦" Or InStr(.Range("g" & i), "D") > 0 Then
                        If D_num < 10 Then
                            .Range("g" & i) = "D0" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        Else
                            .Range("g" & i) = "D" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "自提" Or InStr(.Range("g" & i), "自提") > 0 Then
                        If grasp_num < 10 Then
                            .Range("g" & i) = "自提0" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        Else
                            .Range("g" & i) = "自提" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "专线" Or InStr(.Range("g" & i), "Z") > 0 Then
                        If Z_num < 10 Then
                            .Range("g" & i) = "Z0" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        Else
                            .Range("g" & i) = "Z" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        End If
                        
                        
                    ElseIf .Range("g" & i) = "汇通" Or InStr(.Range("g" & i), "H") > 0 Then
                        If H_num < 10 Then
                            .Range("g" & i) = "H0" & H_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        Else
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "申通" Or InStr(.Range("g" & i), "S") > 0 Then
                        If S_num < 10 Then
                            .Range("g" & i) = "S0" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        Else
                            .Range("g" & i) = "S" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "顺丰" Or InStr(.Range("g" & i), "F") > 0 Then
                        If F_num < 10 Then
                            .Range("g" & i) = "F0" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        Else
                            .Range("g" & i) = "F" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        End If
                    End If
                Next
                
                Call 宏2
            End With
            With Sheets("出货记录（仙人掌）")
                Sheets("出货记录（仙人掌）").Select
                Call 宏9
                     arr2 = .Range("a2:t" & .[a65536].End(3).Row)
                     j2 = 0
                     n2 = 0
                For i = 2 To UBound(arr2) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j2 = j2 + 1 'j stands for how many in that date
                Next
                
                For i = 2 To UBound(arr2) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n2 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
                j2 = j2 + n2 - 1
                
                For i = n2 To j2
                    If n2 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "专线" Then
                        Windows(myFile).Activate
                        With Sheets("Sheet1")
                            If ThisWorkbook.Sheets("出货记录（仙人掌）").Range("d" & i) <> "" Then
                                For start = 3 To .[d65536].End(3).Row
                                    If .Range("d" & start) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("d" & i) Then
                                        temp = start
                                        Do While .Range("c" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        
                                        If Left(.Range("c" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（仙人掌）").Range("c" & i), 2) Then
                                            temp = start
                                            Do While .Range("a" & temp) = ""
                                                temp = temp - 1
                                            Loop
                                            
                                            If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（仙人掌）").Range("b" & i), 2) Then
                                                If (.Range("e" & start) <> "") Then
                                                    ThisWorkbook.Sheets("出货记录（仙人掌）").Range("h" & i) = .Range("e" & start)
                                                End If
                                                Goto 201:
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                For start = 3 To .[c65536].End(3).Row
                                    If Left(.Range("c" & start), 2) = Left(ThisWorkbook.Sheets("出货记录（仙人掌）").Range("c" & i), 2) Then
                                        temp = start
                                        Do While .Range("a" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        
                                        If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（仙人掌）").Range("b" & i), 2) Then
                                            If (.Range("e" & start) <> "") Then
                                                ThisWorkbook.Sheets("出货记录（仙人掌）").Range("h" & i) = .Range("e" & start)
                                            End If
                                            Goto 201:
                                        End If
                                    End If
                                Next
                            End If
                            
                                
201:                    End With
                        Windows("每天出货信息.xls").Activate
                    End If
                Next
                
                Call 宏9
                
                For i = n2 To j2
                    If n2 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "德邦" Or InStr(.Range("g" & i), "D") > 0 Then
                        If D_num < 10 Then
                            .Range("g" & i) = "D0" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        Else
                            .Range("g" & i) = "D" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "自提" Or InStr(.Range("g" & i), "自提") > 0 Then
                        If grasp_num < 10 Then
                            .Range("g" & i) = "自提0" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        Else
                            .Range("g" & i) = "自提" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "专线" Or InStr(.Range("g" & i), "Z") > 0 Then
                        If Z_num < 10 Then
                            .Range("g" & i) = "Z0" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        Else
                            .Range("g" & i) = "Z" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "汇通" Or InStr(.Range("g" & i), "H") > 0 Then
                        If H_num < 10 Then
                            .Range("g" & i) = "H0" & H_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        Else
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "申通" Or InStr(.Range("g" & i), "S") > 0 Then
                        If S_num < 10 Then
                            .Range("g" & i) = "S0" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        Else
                            .Range("g" & i) = "S" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "顺丰" Or InStr(.Range("g" & i), "F") > 0 Then
                        If F_num < 10 Then
                            .Range("g" & i) = "F0" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        Else
                            .Range("g" & i) = "F" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        End If
                    End If
                Next
                                
                Call 宏9
            End With
            With Sheets("出货记录（季节风）")
                Sheets("出货记录（季节风）").Select
                Call 宏14
                     arr3 = .Range("a2:t" & .[a65536].End(3).Row)
                     j3 = 0
                     n3 = 0
                For i = 2 To UBound(arr3) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then j3 = j3 + 1 'j stands for how many in that date
                Next
                
                For i = 2 To UBound(arr3) + 1
                    If .Range("a" & i) = Sheets("配货单").Range("p" & 1) Then
                        n3 = i 'n stands for which is the head of those dates
                    Exit For
                    End If
                Next
                j3 = j3 + n3 - 1
                
                For i = n3 To j3
                    If n3 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "专线" Then
                        Windows(myFile).Activate
                        With Sheets("Sheet1")
                            If ThisWorkbook.Sheets("出货记录（季节风）").Range("d" & i) <> "" Then
                                For start = 3 To .[d65536].End(3).Row
                                    If .Range("d" & start) = ThisWorkbook.Sheets("出货记录（季节风）").Range("d" & i) Then
                                        temp = start
                                        Do While .Range("c" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        
                                        If Left(.Range("c" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（季节风）").Range("c" & i), 2) Then
                                            temp = start
                                            Do While .Range("a" & temp) = ""
                                                temp = temp - 1
                                            Loop
                                            
                                            If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（季节风）").Range("b" & i), 2) Then
                                                If (.Range("e" & start) <> "") Then
                                                    ThisWorkbook.Sheets("出货记录（季节风）").Range("h" & i) = .Range("e" & start)
                                                End If
                                                Goto 202:
                                            End If
                                        End If
                                    End If
                                Next
                            Else
                                For start = 3 To .[c65536].End(3).Row
                                    If Left(.Range("c" & start), 2) = Left(ThisWorkbook.Sheets("出货记录（季节风）").Range("c" & i), 2) Then
                                        temp = start
                                        Do While .Range("a" & temp) = ""
                                            temp = temp - 1
                                        Loop
                                        If Left(.Range("a" & temp), 2) = Left(ThisWorkbook.Sheets("出货记录（季节风）").Range("b" & i), 2) Then
                                            If (.Range("e" & start) <> "") Then
                                                ThisWorkbook.Sheets("出货记录（季节风）").Range("h" & i) = .Range("e" & start)
                                            End If
                                            Goto 202:
                                        End If
                                    End If
                                Next
                            End If
                            
                                
202:                    End With
                        Windows("每天出货信息.xls").Activate
                    End If
                Next
                
                Call 宏14
                
                For i = n3 To j3
                    If n3 = 0 Then
                        Exit For
                    End If
                    If .Range("g" & i) = "德邦" Or InStr(.Range("g" & i), "D") > 0 Then
                        If D_num < 10 Then
                            .Range("g" & i) = "D0" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        Else
                            .Range("g" & i) = "D" & D_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                D_num = D_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "自提" Or InStr(.Range("g" & i), "自提") > 0 Then
                        If grasp_num < 10 Then
                            .Range("g" & i) = "自提0" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        Else
                            .Range("g" & i) = "自提" & grasp_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                grasp_num = grasp_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "专线" Or InStr(.Range("g" & i), "Z") > 0 Then
                        If Z_num < 10 Then
                            .Range("g" & i) = "Z0" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        Else
                            .Range("g" & i) = "Z" & Z_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                Z_num = Z_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "汇通" Or InStr(.Range("g" & i), "H") > 0 Then
                        If H_num < 10 Then
                            .Range("g" & i) = "H0" & H_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        Else
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                H_num = H_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "申通" Or InStr(.Range("g" & i), "S") > 0 Then
                        If S_num < 10 Then
                            .Range("g" & i) = "S0" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        Else
                            .Range("g" & i) = "S" & S_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                S_num = S_num + 1
                            End If
                        End If
                    ElseIf .Range("g" & i) = "顺丰" Or InStr(.Range("g" & i), "F") > 0 Then
                        If F_num < 10 Then
                            .Range("g" & i) = "F0" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        Else
                            .Range("g" & i) = "F" & F_num
                            If .Range("f" & i) <> .Range("f" & i + 1) Then
                                F_num = F_num + 1
                            End If
                        End If
                    End If
                Next
                
                Call 宏14
            End With
            Windows(myFile).Close False
        End If
        myFile = Dir
    Loop
    
    'Workbooks(myFile).Save
    'Workbooks(myFile).Close False         '关闭源工作簿,并不作修改
    
Error: Resume Next
    
End Sub


Sub 宏i_曼途建材()   '提取信息并拷贝公共信息

Sheets("落单").Select
With Sheets("落单")

    
    .Range("B1").Select
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, TAB:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "，", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 1), Array(5, 1), _
        Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    
    
    
    .Cells(1, .Columns.count).End(xlToLeft).Select  '选择第一行最后一个
    
    Selection.TextToColumns Destination:=Range("V2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
    .Cells(1, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("V1").Select     '选出全名
    Selection.TextToColumns Destination:=Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="）", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    .Range("AD1").Select
    Selection.TextToColumns Destination:=Range("AF1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="（", FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True
        
    .Range(Cells(1, "V"), Cells(2, "AF")).Select
    For Each ug In Selection
        ug.Value = Trim(ug)
    Next
    
End With
    Call 宏j_曼途建材
End Sub
Sub 宏j_曼途建材()
    Dim i, num As Integer 'i计数,num代表总箱数
    Dim signal As Boolean '判断是否送货上门
    signal = False
With Sheets("落单")
    '拷贝公共信息
    .Cells(2, .Columns.count).End(xlToLeft).Select
     If Right(Selection, 1) = "门" Then
        signal = True
        Selection.ClearContents
     End If
    num = .Cells(2, .Columns.count).End(xlToLeft)
    For i = 1 To num
        Rows("3:3").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    
    .Range("A3") = Format(Date, "m/d")  '日期
    .Range("A3").Select
    Selection.Copy
    .Range(Cells(3, "A"), Cells(3 + num - 1, "A")).Select
    ActiveSheet.Paste
    
    .Range("AF1").Select '經辦人
    Selection.Copy
    .Range(Cells(3, "T"), Cells(3 + num - 1, "T")).Select
    ActiveSheet.Paste
    
    .Range("AE1").Select '客户id
    Selection.Copy
    .Range(Cells(3, "R"), Cells(3 + num - 1, "R")).Select
    ActiveSheet.Paste
    
    .Range("W1").Select '姓名
    Selection.Copy
    .Range(Cells(3, "F"), Cells(3 + num - 1, "F")).Select
    ActiveSheet.Paste
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '箱數
    Selection.Copy
    .Range(Cells(3, "N"), Cells(3 + num - 1, "N")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '物流費用
    Selection.Copy
    .Range(Cells(3, "Q"), Cells(3 + num - 1, "Q")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '付款方式
    Selection.Copy
    .Range(Cells(3, "P"), Cells(3 + num - 1, "P")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("X1").Select '電話
    Selection.Copy
    .Range(Cells(3, "I"), Cells(3 + num - 1, "I")).Select
    ActiveSheet.Paste
    
    .Range("Y1").Select '省
    Selection.Copy
    .Range(Cells(3, "B"), Cells(3 + num - 1, "B")).Select
    ActiveSheet.Paste
    .Range("Y1").Select
    Selection.ClearContents
    
    .Range("Z1").Select '市
    Selection.Copy
    .Range(Cells(3, "C"), Cells(3 + num - 1, "C")).Select
    ActiveSheet.Paste
    .Range("Z1").Select
    Selection.ClearContents
    
    .Range(Cells(1, "AD"), Cells(3 + num - 1, "AF")).Select
    Selection.ClearContents
    .Cells(1, .Columns.count).End(xlToLeft).Select '網點
    If Right(Selection, 1) <> "区" And Right(Selection, 1) <> "县" And (Right(Selection, 1) <> "市" Or Right(Selection, 2) = "超市") And Right(Selection, 1) <> "省" Then
        Selection.Copy
        .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
        ActiveSheet.Paste
        .Cells(1, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
    End If
    
    '拷贝不同的信息
    .Cells(1, .Columns.count).End(xlToLeft).Select '縣/區/市
    If Right(Selection, 1) = "县" Or Right(Selection, 1) = "区" Or Right(Selection, 1) = "市" Then    '判断最右一个字符是否含区或县
        Selection.Copy
        .Range(Cells(3, "D"), Cells(3 + num - 1, "D")).Select
        ActiveSheet.Paste
    End If
    
    .Cells(2, .Columns.count).End(xlToLeft).Select  '快车/慢车
    If Right(Selection, 1) = "车" Then    '判断最右一个字符是否含车,德邦
        Selection.Copy
        .Range(Cells(3, "O"), Cells(3 + num - 1, "O")).Select
        ActiveSheet.Paste
        .Range("G3") = "德邦"
        .Range("G3").Select
        Selection.Copy
        .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
        ActiveSheet.Paste
        .Cells(2, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
        End If
    ElseIf Right(Selection, 1) = "线" Then '專線
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        Else
            .Range("H3") = "专线物流"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    Else                                '其他
        If Selection = "顺丰" Then
            .Range("G3") = "顺丰"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "申通" Then
            .Range("G3") = "申通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "汇通" Then
            .Range("G3") = "汇通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    End If
    Columns("V:AF").Select
    Selection.Cells.Clear
    
    .Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    Selection.Cut
    Sheets("出货记录").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Range("E" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "e"), Cells(3 + num - 1, "e")), Type:=xlFillDefault
    Range("S" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "s"), Cells(3 + num - 1, "s")), Type:=xlFillDefault
    Range("b" & 1).Select
End With
End Sub

    


Sub 宏i_季节风()   '提取信息并拷贝公共信息

Sheets("落单").Select
With Sheets("落单")

    
    .Range("B1").Select
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, TAB:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "，", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 1), Array(5, 1), _
        Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    
    
    
    .Cells(1, .Columns.count).End(xlToLeft).Select  '选择第一行最后一个
    
    Selection.TextToColumns Destination:=Range("V2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
    .Cells(1, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("V1").Select     '选出全名
    Selection.TextToColumns Destination:=Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="）", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    .Range("AD1").Select
    Selection.TextToColumns Destination:=Range("AF1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="（", FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True
        
    .Range(Cells(1, "V"), Cells(2, "AF")).Select
    For Each ug In Selection
        ug.Value = Trim(ug)
    Next
    
End With
    Call 宏j_季节风
End Sub
Sub 宏j_季节风()
    Dim i, num As Integer 'i计数,num代表总箱数
    Dim signal As Boolean '判断是否送货上门
    signal = False
With Sheets("落单")
    '拷贝公共信息
    .Cells(2, .Columns.count).End(xlToLeft).Select
     If Right(Selection, 1) = "门" Then
        signal = True
        Selection.ClearContents
     End If
    num = .Cells(2, .Columns.count).End(xlToLeft)
    For i = 1 To num
        Rows("3:3").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    
    .Range("A3") = Format(Date, "m/d")  '日期
    .Range("A3").Select
    Selection.Copy
    .Range(Cells(3, "A"), Cells(3 + num - 1, "A")).Select
    ActiveSheet.Paste
    
    .Range("AF1").Select '經辦人
    Selection.Copy
    .Range(Cells(3, "T"), Cells(3 + num - 1, "T")).Select
    ActiveSheet.Paste
    
    .Range("AE1").Select '客户id
    Selection.Copy
    .Range(Cells(3, "R"), Cells(3 + num - 1, "R")).Select
    ActiveSheet.Paste
    
    .Range("W1").Select '姓名
    Selection.Copy
    .Range(Cells(3, "F"), Cells(3 + num - 1, "F")).Select
    ActiveSheet.Paste
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '箱數
    Selection.Copy
    .Range(Cells(3, "N"), Cells(3 + num - 1, "N")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '物流費用
    Selection.Copy
    .Range(Cells(3, "Q"), Cells(3 + num - 1, "Q")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '付款方式
    Selection.Copy
    .Range(Cells(3, "P"), Cells(3 + num - 1, "P")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("X1").Select '電話
    Selection.Copy
    .Range(Cells(3, "I"), Cells(3 + num - 1, "I")).Select
    ActiveSheet.Paste
    
    .Range("Y1").Select '省
    Selection.Copy
    .Range(Cells(3, "B"), Cells(3 + num - 1, "B")).Select
    ActiveSheet.Paste
    .Range("Y1").Select
    Selection.ClearContents
    
    .Range("Z1").Select '市
    Selection.Copy
    .Range(Cells(3, "C"), Cells(3 + num - 1, "C")).Select
    ActiveSheet.Paste
    .Range("Z1").Select
    Selection.ClearContents
    
    .Range(Cells(1, "AD"), Cells(3 + num - 1, "AF")).Select
    Selection.ClearContents
    .Cells(1, .Columns.count).End(xlToLeft).Select '網點
    If Right(Selection, 1) <> "区" And Right(Selection, 1) <> "县" And (Right(Selection, 1) <> "市" Or Right(Selection, 2) = "超市") And Right(Selection, 1) <> "省" Then
        Selection.Copy
        .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
        ActiveSheet.Paste
        .Cells(1, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
    End If
    
    '拷贝不同的信息
    .Cells(1, .Columns.count).End(xlToLeft).Select '縣/區/市
    If Right(Selection, 1) = "县" Or Right(Selection, 1) = "区" Or Right(Selection, 1) = "市" Then    '判断最右一个字符是否含区或县
        Selection.Copy
        .Range(Cells(3, "D"), Cells(3 + num - 1, "D")).Select
        ActiveSheet.Paste
    End If
    
    .Cells(2, .Columns.count).End(xlToLeft).Select  '快车/慢车
    If Right(Selection, 1) = "车" Then    '判断最右一个字符是否含车,德邦
        Selection.Copy
        .Range(Cells(3, "O"), Cells(3 + num - 1, "O")).Select
        ActiveSheet.Paste
        .Range("G3") = "德邦"
        .Range("G3").Select
        Selection.Copy
        .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
        ActiveSheet.Paste
        .Cells(2, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
        End If
    ElseIf Right(Selection, 1) = "线" Then '專線
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        Else
            .Range("H3") = "专线物流"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    Else                                '其他
        If Selection = "顺丰" Then
            .Range("G3") = "顺丰"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "申通" Then
            .Range("G3") = "申通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "汇通" Then
            .Range("G3") = "汇通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    End If
    Columns("V:AF").Select
    Selection.Cells.Clear
    
    .Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    Selection.Cut
    Sheets("出货记录（季节风）").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Range("E" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "e"), Cells(3 + num - 1, "e")), Type:=xlFillDefault
    Range("S" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "s"), Cells(3 + num - 1, "s")), Type:=xlFillDefault
    Range("b" & 1).Select
End With
End Sub

    
Sub 宏i_仙人掌()   '提取信息并拷贝公共信息

Sheets("落单").Select
With Sheets("落单")

    
    .Range("B1").Select
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, TAB:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "，", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 1), Array(5, 1), _
        Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    
    
    
    .Cells(1, .Columns.count).End(xlToLeft).Select  '选择第一行最后一个
    
    Selection.TextToColumns Destination:=Range("V2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
    .Cells(1, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("V1").Select     '选出全名
    Selection.TextToColumns Destination:=Range("AD1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="）", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
    .Range("AD1").Select
    Selection.TextToColumns Destination:=Range("AF1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="（", FieldInfo:=Array(Array(1, 9), Array(2, 1)), TrailingMinusNumbers:=True
        
    .Range(Cells(1, "V"), Cells(2, "AF")).Select
    For Each ug In Selection
        ug.Value = Trim(ug)
    Next
    
End With
    Call 宏j_仙人掌
End Sub
Sub 宏j_仙人掌()
    Dim i, num As Integer 'i计数,num代表总箱数
    Dim signal As Boolean '判断是否送货上门
    signal = False
With Sheets("落单")
    '拷贝公共信息
    .Cells(2, .Columns.count).End(xlToLeft).Select
     If Right(Selection, 1) = "门" Then
        signal = True
        Selection.ClearContents
     End If
    num = .Cells(2, .Columns.count).End(xlToLeft)
    For i = 1 To num
        Rows("3:3").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Next
    
    .Range("A3") = Format(Date, "m/d")  '日期
    .Range("A3").Select
    Selection.Copy
    .Range(Cells(3, "A"), Cells(3 + num - 1, "A")).Select
    ActiveSheet.Paste
    
    .Range("AF1").Select '經辦人
    Selection.Copy
    .Range(Cells(3, "T"), Cells(3 + num - 1, "T")).Select
    ActiveSheet.Paste
    
    .Range("AE1").Select '客户id
    Selection.Copy
    .Range(Cells(3, "R"), Cells(3 + num - 1, "R")).Select
    ActiveSheet.Paste
    
    .Range("W1").Select '姓名
    Selection.Copy
    .Range(Cells(3, "F"), Cells(3 + num - 1, "F")).Select
    ActiveSheet.Paste
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '箱數
    Selection.Copy
    .Range(Cells(3, "N"), Cells(3 + num - 1, "N")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '物流費用
    Selection.Copy
    .Range(Cells(3, "Q"), Cells(3 + num - 1, "Q")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Cells(2, .Columns.count).End(xlToLeft).Select '付款方式
    Selection.Copy
    .Range(Cells(3, "P"), Cells(3 + num - 1, "P")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Range("X1").Select '電話
    Selection.Copy
    .Range(Cells(3, "I"), Cells(3 + num - 1, "I")).Select
    ActiveSheet.Paste
    
    .Range("Y1").Select '省
    Selection.Copy
    .Range(Cells(3, "B"), Cells(3 + num - 1, "B")).Select
    ActiveSheet.Paste
    .Range("Y1").Select
    Selection.ClearContents
    
    .Range("Z1").Select '市
    Selection.Copy
    .Range(Cells(3, "C"), Cells(3 + num - 1, "C")).Select
    ActiveSheet.Paste
    .Range("Z1").Select
    Selection.ClearContents
    
    .Range(Cells(1, "AD"), Cells(3 + num - 1, "AF")).Select
    Selection.ClearContents
    .Cells(1, .Columns.count).End(xlToLeft).Select '網點
    If Right(Selection, 1) <> "区" And Right(Selection, 1) <> "县" And (Right(Selection, 1) <> "市" Or Right(Selection, 2) = "超市") And Right(Selection, 1) <> "省" Then
        Selection.Copy
        .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
        ActiveSheet.Paste
        .Cells(1, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
    End If
    
    '拷贝不同的信息
    .Cells(1, .Columns.count).End(xlToLeft).Select '縣/區/市
    If Right(Selection, 1) = "县" Or Right(Selection, 1) = "区" Or Right(Selection, 1) = "市" Then    '判断最右一个字符是否含区或县
        Selection.Copy
        .Range(Cells(3, "D"), Cells(3 + num - 1, "D")).Select
        ActiveSheet.Paste
    End If
    
    .Cells(2, .Columns.count).End(xlToLeft).Select  '快车/慢车
    If Right(Selection, 1) = "车" Then    '判断最右一个字符是否含车,德邦
        Selection.Copy
        .Range(Cells(3, "O"), Cells(3 + num - 1, "O")).Select
        ActiveSheet.Paste
        .Range("G3") = "德邦"
        .Range("G3").Select
        Selection.Copy
        .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
        ActiveSheet.Paste
        .Cells(2, .Columns.count).End(xlToLeft).Select
        Selection.ClearContents
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
        End If
    ElseIf Right(Selection, 1) = "线" Then '專線
        If signal = True Then
            .Range("H3") = .Range("H3") + "（送货上门）"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        Else
            .Range("H3") = "专线物流"
            .Range("H3").Select
            Selection.Copy
            .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
            ActiveSheet.Paste
            .Range("G3") = "专线"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    Else                                '其他
        If Selection = "顺丰" Then
            .Range("G3") = "顺丰"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "申通" Then
            .Range("G3") = "申通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        ElseIf Selection = "汇通" Then
            .Range("G3") = "汇通"
            .Range("G3").Select
            Selection.Copy
            .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
            ActiveSheet.Paste
        End If
    End If
    Columns("V:AF").Select
    Selection.Cells.Clear
    
    .Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("B1").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range(Cells(3, "a"), Cells(3 + num - 1, "t")).Select
    Selection.Cut
    Sheets("出货记录（仙人掌）").Select
    Rows("2:2").Select
    Selection.Insert Shift:=xlDown
    Range("E" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "e"), Cells(3 + num - 1, "e")), Type:=xlFillDefault
    Range("S" & 3 + num - 1).Select
    Selection.AutoFill Destination:=Range(Cells(2, "s"), Cells(3 + num - 1, "s")), Type:=xlFillDefault
    Range("b" & 1).Select
End With
End Sub

Sub 宏table()
    
'    On Error GoTo erro
    Application.ScreenUpdating = False '关闭屏幕刷新
    Set wordAppl = CreateObject("Word.Application")  '定义一个Word对象变量
    Sheets("配货单").Select
    Dim path, Day, store As String
    Dim label As String
    Dim sub_string As String
    Dim length As Integer
    Dim isStart As Boolean
    Dim count As Integer
    Dim num As Integer, num_box As Integer
    
    count = 0
    num = 1
    num_box = -1
    isStart = False
    
    
With Sheets("配货单")
    Day = Format(.Range("p" & 1), "yyyy年m月d日")
    path = .Range("u" & 1) & "/" & Day & ".doc"
End With

With wordAppl
'    Set openFile = .Documents.Open(path & "模板" & ".doc")
    .Visible = True
    Set objWord = .Documents.Add
    objWord.Activate
    objWord.SaveAs path
    'operation
    With wordAppl.Selection.PageSetup
        .LineNumbering.Active = False
        .Orientation = wdOrientPortrait
        .TopMargin = CentimetersToPoints(1.27)
        .BottomMargin = CentimetersToPoints(1.27)
        .LeftMargin = CentimetersToPoints(1.27)
        .RightMargin = CentimetersToPoints(1.27)
        .Gutter = CentimetersToPoints(0)
        .HeaderDistance = CentimetersToPoints(1.5)
        .FooterDistance = CentimetersToPoints(1.75)
        .PageWidth = CentimetersToPoints(21)
        .PageHeight = CentimetersToPoints(29.7)
        .FirstPageTray = wdPrinterDefaultBin
        .OtherPagesTray = wdPrinterDefaultBin
        .SectionStart = wdSectionNewPage
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
        .BookFoldPrinting = False
        .BookFoldRevPrinting = False
        .BookFoldPrintingSheets = 1
        .GutterPos = wdGutterPosLeft
        .LinesPage = 44
        .LayoutMode = wdLayoutModeLineGrid
    End With
    
    ActiveDocument.Tables.Add Range:=wordAppl.Selection.Range, NumRows:=1, NumColumns:= _
        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    With wordAppl.Selection.Tables(1)
        If .Style <> "网格型" Then
            .Style = "网格型"
        End If
        .ApplyStyleHeadingRows = True
        .ApplyStyleLastRow = False
        .ApplyStyleFirstColumn = True
        .ApplyStyleLastColumn = False
        .ApplyStyleRowBands = True
        .ApplyStyleColumnBands = False
    End With
    
    For i = 1 To [g65536].End(3).Row
        If isStart = True Then
            
            
            If Range("b" & i) <> "" Then
                label = Range("b" & i)
                num_box = Range("k" & i)
                Range("n" & 1) = label
                Range("n" & 1).Select
                Selection.Copy
                wordAppl.Selection.PasteExcelTable False, False, True
                Application.CutCopyMode = False
                Range("n" & 1).Clear
                wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
            End If
            
            If Range("b" & i) = "" Then
                wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
            End If

            Range("g" & i).Select
            Selection.Copy
            wordAppl.Selection.PasteExcelTable False, False, True
            Application.CutCopyMode = False
            wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
            
            store = Range("g" & i)
            Do While InStr(store, "+") <> 0
                
                sub_string = Left(store, InStr(store, "+") - 1)
                length = Len(sub_string)
                
                For j = 2 To Sheets("产品对应表").[b65536].End(3).Row
                    If Sheets("产品对应表").Range("b" & j) = "" Then
                        Goto 997
                    End If
                  
                    If sub_string Like "*" & Sheets("产品对应表").Range("b" & j) & "*" Then
                        sub_string = Replace(sub_string, Sheets("产品对应表").Range("b" & j), "")
                        Range("n" & i) = sub_string
                        Range("n" & i).Select
                        Selection.Copy
                        wordAppl.Selection.PasteExcelTable False, False, True
                        Application.CutCopyMode = False
                        Range("n" & i).Clear
                        Sheets("产品对应表").Select
                        If Sheets("产品对应表").Shapes.Range(Array("图片 " & Sheets("产品对应表").Range("b" & j))) Is Nothing Then Goto 32
                        Sheets("产品对应表").Shapes.Range(Array("图片 " & Sheets("产品对应表").Range("b" & j))).Select
32:                     Selection.Copy
                        wordAppl.Selection.Paste
                        Sheets("配货单").Select
                        Range("n" & 1) = " +  "
                        Range("n" & 1).Select
                        Selection.Copy
                        wordAppl.Selection.PasteExcelTable False, False, True
                        Application.CutCopyMode = False
                        Range("n" & 1).Clear
                        Exit For
                    End If
                    
997:            Next

                store = Mid(store, InStr(store, "+") + 1, Len(store) - length)
            Loop
            
                For j = 2 To Sheets("产品对应表").[b65536].End(3).Row
                    If Sheets("产品对应表").Range("b" & j) = "" Then
                        Goto 998
                    End If
                    
                    If store Like "*" & Sheets("产品对应表").Range("b" & j) & "*" Then
                        store = Replace(store, Sheets("产品对应表").Range("b" & j), "")
                        Range("n" & i) = store
                        Range("n" & i).Select
                        Selection.Copy
                        wordAppl.Selection.PasteExcelTable False, False, True
                        Application.CutCopyMode = False
                        Range("n" & i).Clear
                        Sheets("产品对应表").Select
                        If Sheets("产品对应表").Shapes.Range(Array("图片 " & Sheets("产品对应表").Range("b" & j))) Is Nothing Then Goto 33
                        Sheets("产品对应表").Shapes.Range(Array("图片 " & Sheets("产品对应表").Range("b" & j))).Select
33:                     Selection.Copy
                        wordAppl.Selection.Paste
                        Sheets("配货单").Select
                        Exit For
                    End If
998:            Next
                
            wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
            
            Range("n" & 1) = num
            Range("n" & 1).Select
            Selection.Copy
            wordAppl.Selection.PasteExcelTable False, False, True
            Application.CutCopyMode = False
            Range("n" & 1).Clear
            num = num + 1
            count = count + 1
            wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
            
            wordAppl.Selection.InsertRows 1
            wordAppl.Selection.Collapse Direction:=wdCollapseStart
            
'            If count = num_box Then
'                If count = 1 Then
'                    GoTo 999
'                End If
'                wordAppl.Selection.MoveUp Unit:=wdLine, count:=count
'                wordAppl.Selection.MoveDown Unit:=wdLine, count:=count - 1, Extend:=wdExtend
'                wordAppl.Selection.Cells.Merge
'                wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
'                wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
'                wordAppl.Selection.TypeBackspace
'999:            count = 0
'            End If
        End If
        
        If Range("g" & i) = "产品名称" Then
            isStart = True
        End If
        
        If Range("g" & i + 1) = Range("g" & 1) Then
            isStart = False
        End If
    Next
    
    
    
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Font.Size = 24
    wordAppl.Selection.SelectCell
    wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wordAppl.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
'    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
'    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Font.Size = 14
'    wordAppl.Selection.SelectCell
'    wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
'    wordAppl.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Font.Size = 16
    wordAppl.Selection.SelectCell
    wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wordAppl.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
    wordAppl.Selection.SelectCell
    wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wordAppl.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitContent)
    wordAppl.Selection.Font.Size = 18
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    
    wordAppl.Selection.MoveLeft Unit:=wdCharacter, count:=3, Extend:=wdExtend
    wordAppl.Selection.Rows.Delete
    
    wordAppl.Selection.WholeStory
    wordAppl.Selection.Font.Bold = wdToggle
    wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    wordAppl.Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
    
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=1, Extend:=wdExtend
    wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=1
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
    wordAppl.Selection.Columns.PreferredWidthType = wdPreferredWidthPoints
    wordAppl.Selection.Columns.PreferredWidth = CentimetersToPoints(4.9)
    
    wordAppl.Selection.MoveDown Unit:=wdLine, count:=10000
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=1
    wordAppl.Selection.MoveUp Unit:=wdLine, count:=10000, Extend:=wdExtend
    wordAppl.Selection.MoveRight Unit:=wdCharacter, count:=3, Extend:=wdExtend
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    wordAppl.Selection.Tables(1).AutoFitBehavior (wdAutoFitWindow)
    
    objWord.Save
    objWord.Close
    .Quit False
End With

    Set wordAppl = Nothing '释放存储空间
    Application.ScreenUpdating = False '关闭屏幕刷新
End
Exit Sub
erro:
    Selection.Copy
    Resume
End Sub