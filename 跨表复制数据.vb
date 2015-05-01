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
    Dim copyFromFileName As String
    Dim myPath As String
    Dim myFile As String
    Dim openFile As Workbook
    
    Dim signal As Boolean
    
    
    Dim endRow As Integer
    Dim endColumn As Integer
    Dim endColumnChar As String
    Dim rang As String
    
    
    copyFromFileName = "落单表-梁子婷.xls"    '这个地方设置被复制的excel文件
    copyFromFileName1 = "落单表.xls"    '这个地方设置被复制的excel文件

    myPath = "D:\用户目录\Documents\落单表" & "/" '把文件路径定义给变量
    myFile = Dir(myPath & "*.xls")   '依次找寻指定路径中的*.xls文件
    
    Do While myFile <> ""
        If myFile = copyFromFileName Or myFile = copyFromFileName1 Then   '假如遍历到需要复制的文件
            If myFile = copyFromFileName Then
                Set openFile = Workbooks.Open(myPath & myFile) '打开符合要求的文件
            End If
            If myFile = copyFromFileName1 Then
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
                            For o = n3 To j3
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录").Range("f" & o) Then
                                    num = 0
                                    For p = n3 To j3
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录").Range("f" & p) And .Range("h" & l) = ThisWorkbook.Sheets("出货记录").Range("h" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录").Range("a" & p) Then
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
                            For o = n4 To j4
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("f" & o) Then
                                    num = 0
                                    For p = n4 To j4
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("f" & p) And .Range("h" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("h" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录（仙人掌）").Range("a" & p) Then
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
                            For o = n5 To j5
                                If .Range("a" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("a" & o) And .Range("f" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("f" & o) Then
                                    num = 0
                                    For p = n5 To j5
                                        If .Range("f" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("f" & p) And .Range("h" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("h" & p) And .Range("a" & l) = ThisWorkbook.Sheets("出货记录（季节风）").Range("a" & p) Then
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
    
    lookfromfilename = "中国各个市县名称汇总--史上最全(最新更新).xls"
    myPath = "F:\Business\季节风" & "/"
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
                                        If Left(.Range("a" & start), 2) = Left(ThisWorkbook.Sheets("出货记录").Range("b" & i), 2) Then
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
                                        If Left(.Range("a" & start), 2) = Left(ThisWorkbook.Sheets("出货记录（仙人掌）").Range("b" & i), 2) Then
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
                                        If Left(.Range("a" & start), 2) = Left(ThisWorkbook.Sheets("出货记录（季节风）").Range("b" & i), 2) Then
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
            End With
            Windows(myFile).Close False
        End If
        myFile = Dir
    Loop
    
    'Workbooks(myFile).Save
    'Workbooks(myFile).Close False         '关闭源工作簿,并不作修改
    
    
    
End Sub
