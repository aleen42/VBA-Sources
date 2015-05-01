Sub 宏1()
Dim i As Integer, j As Integer, store As Integer, num As Integer, store1 As Integer, store2 As Integer
With Sheets("发货价格表")
     arr1 = .Range("a5:v" & .[a65536].End(3).Row)
     arr2 = .Range("w3:ad" & .[w65536].End(3).Row)
     For i = 1 To UBound(arr2)
        For j = 1 To UBound(arr1)
            If InStr(arr1(j, 2), arr2(i, 1)) = 1 And InStr(arr1(j, 1), "广东") = 0 Then
                
                store = arr2(i, 2) * arr1(j, 5) + 14
                If store > arr1(j, 6) Then
                    .Range("z" & 2 + i) = store
                Else: .Range("z" & 2 + i) = arr1(j, 6)
                End If
                
                store = arr2(i, 2) * arr1(j, 9) + 14
                If store > arr1(j, 10) Then
                    .Range("aa" & 2 + i) = store
                Else: .Range("aa" & 2 + i) = arr1(j, 10)
                End If
                
                num = arr2(i, 3)
                If num <= 3 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 11)
                ElseIf num > 3 And num <= 10 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 12)
                ElseIf num > 10 And num <= 20 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 13)
                Else: .Range("ab" & 2 + i) = num * arr1(j, 14)
                End If
                
                .Range("p" & j + 4).Select
                .Range("v" & j + 4) = Selection.Cells.Value
                store1 = .Range("v" & j + 4)
                .Range("q" & j + 4).Select
                .Range("v" & j + 5) = Selection.Cells.Value
                store2 = .Range("v" & j + 5)
                .Range("v" & j + 4).Clear
                .Range("v" & j + 5).Clear
                .Range("ac" & 2 + i) = store1 + (arr2(i, 2) - 1) * store2
                
                .Range("s" & j + 4).Select
                .Range("v" & j + 4) = Selection.Cells.Value
                store1 = .Range("v" & j + 4)
                .Range("t" & j + 4).Select
                .Range("v" & j + 5) = Selection.Cells.Value
                store2 = .Range("v" & j + 5)
                .Range("v" & j + 4).Clear
                .Range("v" & j + 5).Clear
                .Range("ad" & 2 + i) = store1 + (arr2(i, 2) - 1) * store2
                
                Exit For
            ElseIf InStr(arr1(j, 2), arr2(i, 1)) = 1 And InStr(arr1(j, 2), arr2(i, 1)) = 1 Then
                
                store = arr2(i, 2) * arr1(j, 5) + 12
                If store > arr1(j, 6) Then
                    .Range("z" & 2 + i) = store
                Else: .Range("z" & 2 + i) = arr1(j, 6)
                End If
                
                store = arr2(i, 2) * arr1(j, 9) + 12
                If store > arr1(j, 10) Then
                    .Range("aa" & 2 + i) = store
                Else: .Range("aa" & 2 + i) = arr1(j, 10)
                End If
                
                num = arr2(i, 3)
                If num <= 3 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 11)
                ElseIf num > 3 And num <= 10 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 12)
                ElseIf num > 10 And num <= 20 Then
                    .Range("ab" & 2 + i) = num * arr1(j, 13)
                Else: .Range("ab" & 2 + i) = num * arr1(j, 14)
                End If
                
                .Range("p" & j + 4).Select
                .Range("v" & j + 4) = Selection.Cells.Value
                store1 = .Range("v" & j + 4)
                .Range("q" & j + 4).Select
                .Range("v" & j + 5) = Selection.Cells.Value
                store2 = .Range("v" & j + 5)
                .Range("v" & j + 4).Clear
                .Range("v" & j + 5).Clear
                .Range("ac" & 2 + i) = store1 + (arr2(i, 2) - 1) * store2
                
                .Range("s" & j + 4).Select
                .Range("v" & j + 4) = Selection.Cells.Value
                store1 = .Range("v" & j + 4)
                .Range("t" & j + 4).Select
                .Range("v" & j + 5) = Selection.Cells.Value
                store2 = .Range("v" & j + 5)
                .Range("v" & j + 4).Clear
                .Range("v" & j + 5).Clear
                .Range("ad" & 2 + i) = store1 + (arr2(i, 2) - 1) * store2
                
                Exit For
            Else:
            End If
        Next
    Next
.Range("w1:ad" & .[w65536].End(3).Row).Select
With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
End With
.Range("z" & 5).Select
.Range("z" & 1).Select
End With
End Sub


