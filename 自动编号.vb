Sub 按钮2_Click()
    Dim i As Integer, j1 As Integer, j2 As Integer, j3 As Integer, n1 As Integer, n2 As Integer, n3 As Integer, index As Integer
    Dim D_num As Integer, H_num As Integer, Z_num As Integer, S_num As Integer, F_num As Integer
    D_num = 1
    H_num = 1
    Z_num = 1
    S_num = 1
    F_num = 1
    
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
                                
    End With
    With Sheets("出货记录（仙人掌）")
        Sheets("出货记录（仙人掌）").Select
        Call 宏2
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
                        
    End With
    With Sheets("出货记录（季节风）")
        Sheets("出货记录（季节风）").Select
        Call 宏2
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
                        
    End With
End Sub






