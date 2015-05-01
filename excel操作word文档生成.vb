Public Function iRnd(a As Integer, b As Integer) As Integer         '产生随机数
    iRnd = Int(Rnd * (Abs(a - b) + 1)) + IIf(a >= b, b, a)
End Function


Sub 宏1()

    On Error Goto erro1
    
    Application.ScreenUpdating = False '关闭屏幕刷新
    Set wordAppl = CreateObject("Word.Application")  '定义一个Word对象变量
    Sheets("Sheet1").Select
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=ROUNDUP(RC[-26]*R1C8,0)"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA38"), Type:=xlFillDefault
    Range("BA2:BA38").Select
    Selection.AutoFill Destination:=Range("BA2:BZ38"), Type:=xlFillDefault
    
    Dim path As String
    Dim column As Integer
    
    Dim clock As Integer, hour_top As Integer, hour_bottom As Integer, interval As Integer
    
    
    Dim name As String
    Dim num As Integer
    Dim v_num As Integer '变量数
    Dim rand_num As Integer
    Dim big_pic As String
    Dim small_pic As String
    Dim i As Integer, j As Integer
    Dim ch As String
    Dim selected As Integer
    Dim selected_num As Integer
    
    Dim column_str As String
    
    Dim openFile As Word.Document
    
    Dim paste_string As String
    
With Sheets("Sheet1")
    path = .Range("n" & 1) & "/"
    num = .Range("h" & 1)
    hour_top = .Range("j" & 1) * 60
    hour_bottom = .Range("l" & 1) * 60
    If num <> 1 Then
        interval = (hour_bottom - hour_top) / (num - 1)
    Else
        interval = hour_bottom - hour_top
    End If
End With
    

With wordAppl
'    Set openFile = .Documents.Open(path & "模板" & ".doc")
    .Visible = False
'    .Activate
'
'    wordAppl.Selection.WholeStory
'    wordAppl.Selection.Copy
    
    For i = 0 To (num - 1)
        Set objWord = .Documents.Add
        objWord.Activate
        If (hour_top + i * interval) Mod 60 = 0 Then
            name = (hour_top + i * interval) / 60 & "点"
        Else
            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
        End If
        objWord.SaveAs path & name & ".doc"
'        wordAppl.Selection.PasteAndFormat (wdPasteDefault)
'        Application.ScreenUpdating = False '关闭屏幕刷新
        
        'operations
        Range("a" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        
'        Dim objData As New DataObject
'        objData.SetText ""
'        objData.PutInClipboard
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 2) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 2) <> 0 Then
                Range(column_str & 2) = Range(column_str & 2) - 1
                Range(Chr(selected + 96) & 2).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        wordAppl.Selection.TypeParagraph
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 3) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 3).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 4) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 4).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 5) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 5).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 6) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 6).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 7) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 7) <> 0 Then
                Range(column_str & 7) = Range(column_str & 7) - 1
                Range(Chr(selected + 96) & 7).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 8) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 8).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 9) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 9) <> 0 Then
                Range(column_str & 9) = Range(column_str & 9) - 1
                Range(Chr(selected + 96) & 9).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 10) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 10).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 11) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 11) <> 0 Then
                Range(column_str & 11) = Range(column_str & 11) - 1
                Range(Chr(selected + 96) & 11).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
       
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 12) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 12).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array("图片 8")).Select
        Selection.Copy
        wordAppl.Selection.Paste
        wordAppl.Selection.TypeParagraph
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 13) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 13).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 14) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 14).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("a" & 15).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range("a" & 16).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 26) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 26) <> 0 Then
                Range(column_str & 26) = Range(column_str & 26) - 1
                Exit Do
            Else
            End If
        Loop

        big_pic = "图片" & Range(Chr(selected + 96) & 26)
        small_pic = "图片小" & Range(Chr(selected + 96) & 26)
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        wordAppl.Selection.TypeParagraph


        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 17) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 17) <> 0 Then
                Range(column_str & 17) = Range(column_str & 17) - 1
                Exit Do
            Else
            End If
        Loop
        Range(Chr(selected + 96) & 17).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 24) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected_num = iRnd(1, column)
            column_str = "b" & Chr(selected_num + 96)
            If Range(column_str & 24) <> 0 Then
                Range(column_str & 24) = Range(column_str & 24) - 1
                Exit Do
            Else
            End If
        Loop
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 19) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 19) <> 0 Then
                Range(column_str & 19) = Range(column_str & 19) - 1
                Exit Do
            Else
            End If
        Loop
        
        If Range(Chr(selected + 96) & 19) = "平方米" Or Range(Chr(selected + 96) & 19) = "平方" Or Range(Chr(selected + 96) & 19) = "方" Then
            Range(Chr(selected + 96) & 18) = Range(Chr(selected_num + 96) & 24) / 11
        Else
            Range(Chr(selected + 96) & 18) = Range(Chr(selected_num + 96) & 24)
        End If
        
        Range(Chr(selected + 96) & 18).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range(Chr(selected + 96) & 19).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range(Chr(selected + 96) & 20).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 21) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 21).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 22) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 22).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("a" & 23).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range(Chr(selected_num + 96) & 24).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range("a" & 25).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        
        Range(Chr(selected + 96) & 26).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False

        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 33) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        
        selected = iRnd(1, column)
        Range(Chr(selected + 96) & 33).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Range("b" & 1).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False

        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 34) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Range(Chr(iRnd(1, column) + 96) & 34).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
'            If iRnd(1, 2) = 1 Then
'            big_pic = Range("a" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("a" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("a" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        Else
'            big_pic = Range("b" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("b" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("b" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        End If
'
        
        
        'operations
        
    
        
        
        objWord.Save
        objWord.Close
    Next
    
'    For i = 0 To (num - 1)
'
'        If (hour_top + i * interval) Mod 60 = 0 Then
'            name = (hour_top + i * interval) / 60 & "点"
'        Else
'            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
'        End If
'
'        .Documents.Open (path & name & ".doc")
'
'
'        Operation
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'        Operation
'
'
'        .objWord.Save
'        .objWord.Close
'    Next

    '.objWord.Close False
    .Quit False
End With
    Set wordAppl = Nothing '释放存储空间
    Application.ScreenUpdating = False '关闭屏幕刷新
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-26]*R1C8"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA32"), Type:=xlFillDefault
    Range("BA2:BA32").Select
    Selection.AutoFill Destination:=Range("BA2:BP32"), Type:=xlFillDefault
    Range("BA2:BP32").Select
    Range("a" & 1).Select
    
End
Exit Sub
erro1:
    Selection.Copy
    Resume
End Sub


Sub 宏2()

    On Error Goto erro2

    Application.ScreenUpdating = False '关闭屏幕刷新
    Set wordAppl = CreateObject("Word.Application")  '定义一个Word对象变量
    Sheets("Sheet1").Select
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=ROUNDUP(RC[-26]*R1C8,0)"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA38"), Type:=xlFillDefault
    Range("BA2:BA38").Select
    Selection.AutoFill Destination:=Range("BA2:BZ38"), Type:=xlFillDefault
    
    Dim path As String
    Dim column As Integer
    
    Dim clock As Integer, hour_top As Integer, hour_bottom As Integer, interval As Integer
    
    
    Dim name As String
    Dim num As Integer
    Dim v_num As Integer '变量数
    Dim rand_num As Integer
    Dim big_pic As String
    Dim small_pic As String
    Dim i As Integer, j As Integer
    Dim ch As String
    Dim selected As Integer
    Dim selected1 As Integer, selected2 As Integer
    
    Dim selected_num As Integer
    Dim column_str As String
    
    Dim openFile As Word.Document
    
    Dim err As Boolean
    
    Dim paste_string As String
    
    
With Sheets("Sheet1")
    path = .Range("o" & 1) & "/"
    num = .Range("h" & 1)
    hour_top = .Range("j" & 1) * 60
    hour_bottom = .Range("l" & 1) * 60
    If num <> 1 Then
        interval = (hour_bottom - hour_top) / (num - 1)
    Else
        interval = hour_bottom - hour_top
    End If
End With
    

With wordAppl
'    Set openFile = .Documents.Open(path & "模板" & ".doc")
    .Visible = False
'    .Activate
'
'    wordAppl.Selection.WholeStory
'    wordAppl.Selection.Copy
    
    For i = 0 To (num - 1)
        Set objWord = .Documents.Add
        objWord.Activate
        If (hour_top + i * interval) Mod 60 = 0 Then
            name = (hour_top + i * interval) / 60 & "点"
        Else
            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
        End If
        objWord.SaveAs path & name & ".doc"
'        wordAppl.Selection.PasteAndFormat (wdPasteDefault)
'        Application.ScreenUpdating = False '关闭屏幕刷新
        
        'operations
        
        Do
            err = Range("a" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        
'        Dim objData As New DataObject
'        objData.SetText ""
'        objData.PutInClipboard
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 2) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 2) <> 0 Then
                Range(column_str & 2) = Range(column_str & 2) - 1
                Do
                    err = Range(Chr(selected + 96) & 2).Select
                    Selection.Copy
                    If err = True Then
                        Exit Do
                    End If
                Loop
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        wordAppl.Selection.TypeParagraph
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 3) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 3).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 4) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 4).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 5) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 5).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 6) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 6).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 7) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 7) <> 0 Then
                Range(column_str & 7) = Range(column_str & 7) - 1
                Do
                    err = Range(Chr(selected + 96) & 7).Select
                    Selection.Copy
                    If err = True Then
                        Exit Do
                    End If
                Loop
                
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 8) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 8).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 9) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 9) <> 0 Then
                Range(column_str & 9) = Range(column_str & 9) - 1
                Do
                    err = Range(Chr(selected + 96) & 9).Select
                    Selection.Copy
                    If err = True Then
                        Exit Do
                    End If
                Loop
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 10) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 10).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 11) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 11) <> 0 Then
                Range(column_str & 11) = Range(column_str & 11) - 1
                Do
                    err = Range(Chr(selected + 96) & 11).Select
                    Selection.Copy
                    If err = True Then
                        Exit Do
                    End If
                Loop
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
       
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 12) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 12).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array("图片 8")).Select
        Selection.Copy
        wordAppl.Selection.Paste
        wordAppl.Selection.TypeParagraph
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 13) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 13).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 14) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 14).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("a" & 15).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 16).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 26) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        
        Do
            selected1 = iRnd(1, column)
            column_str = "b" & Chr(selected1 + 96)
            If Range(column_str & 26) <> 0 Then
                Range(column_str & 26) = Range(column_str & 26) - 1
                Exit Do
            Else
            End If
        Loop
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 32) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        
        Do
            selected2 = iRnd(1, column)
            column_str = "b" & Chr(selected2 + 96)
            If Range(column_str & 32) <> 0 Then
                Range(column_str & 32) = Range(column_str & 32) - 1
                Exit Do
            Else
            End If
        Loop

        big_pic = "图片" & Range(Chr(selected1 + 96) & 26)
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        big_pic = "图片" & Range(Chr(selected2 + 96) & 32)
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        
        wordAppl.Selection.TypeParagraph

        Do
            err = Range("a" & 27).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("a" & 23).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 24) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected_num = iRnd(1, column)
            column_str = "b" & Chr(selected_num + 96)
            If Range(column_str & 24) <> 0 Then
                Range(column_str & 24) = Range(column_str & 24) - 1
                Exit Do
            Else
            End If
        Loop
        Range(Chr(selected_num + 96) & 24).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Do
            err = Range("a" & 25).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        small_pic = "图片小" & Range(Chr(selected1 + 96) & 26)
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        
        Do
            err = Range("a" & 29).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 30) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected_num = iRnd(1, column)
            column_str = "b" & Chr(selected_num + 96)
            If Range(column_str & 30) <> 0 Then
                Range(column_str & 30) = Range(column_str & 30) - 1
                Exit Do
            Else
            End If
        Loop
        Range(Chr(selected_num + 96) & 30).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Do
            err = Range("a" & 31).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        small_pic = "图片小" & Range(Chr(selected2 + 96) & 32)
        ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        
        Do
            err = Range(Chr(selected1 + 96) & 26).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        wordAppl.Selection.TypeBackspace
        Do
            err = Range("c" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        Do
            err = Range(Chr(selected2 + 96) & 32).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        wordAppl.Selection.MoveLeft Unit:=wdCharacter, Count:=Len(Range(Chr(selected2 + 96) & 32)) - 1
        wordAppl.Selection.TypeBackspace
        wordAppl.Selection.MoveRight Unit:=wdCharacter, Count:=Len(Range(Chr(selected2 + 96) & 32))
        wordAppl.Selection.TypeParagraph
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False

        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 33) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 33).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        Do
            err = Range("b" & 1).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False

        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 34) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            err = Range(Chr(iRnd(1, column) + 96) & 34).Select
            Selection.Copy
            If err = True Then
                Exit Do
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
'            If iRnd(1, 2) = 1 Then
'            big_pic = Range("a" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("a" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("a" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        Else
'            big_pic = Range("b" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("b" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("b" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        End If
'
        
        
        'operations
        
    
        
        
        objWord.Save
        objWord.Close
    Next
    
'    For i = 0 To (num - 1)
'
'        If (hour_top + i * interval) Mod 60 = 0 Then
'            name = (hour_top + i * interval) / 60 & "点"
'        Else
'            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
'        End If
'
'        .Documents.Open (path & name & ".doc")
'
'
'        Operation
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'        Operation
'
'
'        .objWord.Save
'        .objWord.Close
'    Next

    '.objWord.Close False
    .Quit False
End With
    Set wordAppl = Nothing '释放存储空间
    Application.ScreenUpdating = False '关闭屏幕刷新
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-26]*R1C8"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA32"), Type:=xlFillDefault
    Range("BA2:BA32").Select
    Selection.AutoFill Destination:=Range("BA2:BP32"), Type:=xlFillDefault
    Range("BA2:BP32").Select
    Range("a" & 1).Select
End
Exit Sub
erro2:
    Selection.Copy
    Resume

End Sub


Sub 宏3()
    
    On Error Goto erro3
    
    Application.ScreenUpdating = False '关闭屏幕刷新
    Set wordAppl = CreateObject("Word.Application")  '定义一个Word对象变量
    Sheets("Sheet2").Select
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=ROUNDUP(RC[-26]*R1C8,0)"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA38"), Type:=xlFillDefault
    Range("BA2:BA38").Select
    Selection.AutoFill Destination:=Range("BA2:BZ38"), Type:=xlFillDefault
    
    Dim path As String
    Dim column As Integer
    
    Dim clock As Integer, hour_top As Integer, hour_bottom As Integer, interval As Integer
    
    
    Dim name As String
    Dim num As Integer
    Dim v_num As Integer '变量数
    Dim rand_num As Integer
    Dim big_pic As String
    Dim small_pic As String
    Dim i As Integer, j As Integer
    Dim ch As String
    Dim selected As Integer
    Dim selected_num As Integer
    Dim selected_product As Integer
    Dim column_str As String
    
    Dim openFile As Word.Document
    
    Dim paste_string As String
    
    
With Sheets("Sheet2")
    path = .Range("n" & 1) & "/"
    num = .Range("h" & 1)
    hour_top = .Range("j" & 1) * 60
    hour_bottom = .Range("l" & 1) * 60
    If num <> 1 Then
        interval = (hour_bottom - hour_top) / (num - 1)
    Else
        interval = hour_bottom - hour_top
    End If
    
    column = 0
    For j = 0 To 100
        If Range(Chr(j + 97) & 26) <> "" Then
            column = column + 1
        Else
            Exit For
        End If
    Next
    For j = 1 To column
        Range(Chr(j + 96) & 32) = "（" + Range(Chr(j + 96) & 26) + "）"
    Next
End With
    

With wordAppl
'    Set openFile = .Documents.Open(path & "模板" & ".doc")
    .Visible = False
'    .Activate
'
'    wordAppl.Selection.WholeStory
'    wordAppl.Selection.Copy
    
    For i = 0 To (num - 1)
        Set objWord = .Documents.Add
        objWord.Activate
        If (hour_top + i * interval) Mod 60 = 0 Then
            name = (hour_top + i * interval) / 60 & "点"
        Else
            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
        End If
        objWord.SaveAs path & name & ".doc"
'        wordAppl.Selection.PasteAndFormat (wdPasteDefault)
'        Application.ScreenUpdating = False '关闭屏幕刷新
        
        'operations
        
'        Dim objData As New DataObject
'        objData.SetText ""
'        objData.PutInClipboard
        
        
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 2) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 2) <> 0 Then
                Range(column_str & 2) = Range(column_str & 2) - 1
                Range(Chr(selected + 96) & 2).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 3) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 3) <> 0 Then
                Range(column_str & 3) = Range(column_str & 3) - 1
                Range(Chr(selected + 96) & 3).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 4) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 4) <> 0 Then
                Range(column_str & 4) = Range(column_str & 4) - 1
                Range(Chr(selected + 96) & 4).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 5) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 5) <> 0 Then
                Range(column_str & 5) = Range(column_str & 5) - 1
                Range(Chr(selected + 96) & 5).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 6) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 6) <> 0 Then
                Range(column_str & 6) = Range(column_str & 6) - 1
                Range(Chr(selected + 96) & 6).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
'        column = 0
'        For j = 0 To 100
'            If Range(Chr(j + 97) & 7) <> "" Then
'                column = column + 1
'            Else
'                Exit For
'            End If
'        Next
'        Do
'            selected = iRnd(1, column)
'            column_str = "b" & Chr(selected + 96)
'            If Range(column_str & 7) <> 0 Then
'                Range(column_str & 7) = Range(column_str & 7) - 1
'                Range(Chr(selected + 96) & 7).Select
'                Selection.Copy
'                Exit Do
'            Else
'            End If
'        Loop
'        wordAppl.Selection.PasteExcelTable False, False, True
'        wordAppl.Selection.TypeParagraph
'        wordAppl.Selection.TypeParagraph
'        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 8) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 8) <> 0 Then
                Range(column_str & 8) = Range(column_str & 8) - 1
                Range(Chr(selected + 96) & 8).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 9) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 9) <> 0 Then
                Range(column_str & 9) = Range(column_str & 9) - 1
                Range(Chr(selected + 96) & 9).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        ThisWorkbook.Sheets("Sheet2").Shapes.Range(Array("图片 8")).Select
        Selection.Copy
        wordAppl.Selection.Paste
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 10) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 10) <> 0 Then
                Range(column_str & 10) = Range(column_str & 10) - 1
                Range(Chr(selected + 96) & 10).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 11) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 11) <> 0 Then
                Range(column_str & 11) = Range(column_str & 11) - 1
                Range(Chr(selected + 96) & 11).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 12) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 12) <> 0 Then
                Range(column_str & 12) = Range(column_str & 12) - 1
                Range(Chr(selected + 96) & 12).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 13) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 13) <> 0 Then
                Range(column_str & 13) = Range(column_str & 13) - 1
                Range(Chr(selected + 96) & 13).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 14) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 14) <> 0 Then
                Range(column_str & 14) = Range(column_str & 14) - 1
                Range(Chr(selected + 96) & 14).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 15) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 15) <> 0 Then
                Range(column_str & 15) = Range(column_str & 15) - 1
                Range(Chr(selected + 96) & 15).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 16) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 16) <> 0 Then
                Range(column_str & 16) = Range(column_str & 16) - 1
                Range(Chr(selected + 96) & 16).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 17) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 17) <> 0 Then
                Range(column_str & 17) = Range(column_str & 17) - 1
                Range(Chr(selected + 96) & 17).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 18) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 18) <> 0 Then
                Range(column_str & 18) = Range(column_str & 18) - 1
                Range(Chr(selected + 96) & 18).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 19) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 19) <> 0 Then
                Range(column_str & 19) = Range(column_str & 19) - 1
                Range(Chr(selected + 96) & 19).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 20) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 20) <> 0 Then
                Range(column_str & 20) = Range(column_str & 20) - 1
                Range(Chr(selected + 96) & 20).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 21) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 21) <> 0 Then
                Range(column_str & 21) = Range(column_str & 21) - 1
                Range(Chr(selected + 96) & 21).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 22) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 22) <> 0 Then
                Range(column_str & 22) = Range(column_str & 22) - 1
                Range(Chr(selected + 96) & 22).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
    
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 23) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 23) <> 0 Then
                Range(column_str & 23) = Range(column_str & 23) - 1
                Range(Chr(selected + 96) & 23).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
    
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 30) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected_num = iRnd(1, column)
            column_str = "b" & Chr(selected_num + 96)
            If Range(column_str & 30) <> 0 Then
                Range(column_str & 30) = Range(column_str & 30) - 1
                Exit Do
            Else
            End If
        Loop
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 25) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 25) <> 0 Then
                Range(column_str & 25) = Range(column_str & 25) - 1
                Exit Do
            Else
            End If
        Loop
        
        If Range(Chr(selected + 96) & 25) = "平方米" Or Range(Chr(selected + 96) & 25) = "平方" Or Range(Chr(selected + 96) & 25) = "方" Then
            Range(Chr(selected + 96) & 24) = Range(Chr(selected_num + 96) & 30) / 11
        Else
            Range(Chr(selected + 96) & 24) = Range(Chr(selected_num + 96) & 30)
        End If
        
        

        Range(Chr(selected + 96) & 24).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range(Chr(selected + 96) & 25).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 32) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected_product = iRnd(1, column)
            column_str = "b" & Chr(selected_product + 96)
            If Range(column_str & 32) <> 0 Then
                Range(column_str & 32) = Range(column_str & 32) - 1
                Exit Do
            Else
            End If
        Loop
        
        Range(Chr(selected_product + 96) & 26).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 27) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 27) <> 0 Then
                Range(column_str & 27) = Range(column_str & 27) - 1
                Range(Chr(selected + 96) & 27).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 28) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 28) <> 0 Then
                Range(column_str & 28) = Range(column_str & 28) - 1
                Range(Chr(selected + 96) & 28).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeParagraph
        Application.CutCopyMode = False
        wordAppl.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 29) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 29) <> 0 Then
                Range(column_str & 29) = Range(column_str & 29) - 1
                Range(Chr(selected + 96) & 29).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        Range(Chr(selected_num + 96) & 30).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        column = 0
        For j = 0 To 100
            If Range(Chr(j + 97) & 31) <> "" Then
                column = column + 1
            Else
                Exit For
            End If
        Next
        Do
            selected = iRnd(1, column)
            column_str = "b" & Chr(selected + 96)
            If Range(column_str & 31) <> 0 Then
                Range(column_str & 31) = Range(column_str & 31) - 1
                Range(Chr(selected + 96) & 31).Select
                Selection.Copy
                Exit Do
            Else
            End If
        Loop
        wordAppl.Selection.PasteExcelTable False, False, True
        Application.CutCopyMode = False
        
        small_pic = "图片小" & Range(Chr(selected_product + 96) & 32)
        
        ThisWorkbook.Sheets("Sheet2").Shapes.Range(Array(small_pic)).Select
        Selection.Copy
        wordAppl.Selection.Paste
        
        Range(Chr(selected_product + 96) & 32).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        Range(Chr(selected + 96) & 37).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
        Range(Chr(selected + 96) & 38).Select
        Selection.Copy
        wordAppl.Selection.PasteExcelTable False, False, True
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        wordAppl.Selection.TypeText Text:=" "
        Application.CutCopyMode = False
        
'            If iRnd(1, 2) = 1 Then
'            big_pic = Range("a" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("a" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("a" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        Else
'            big_pic = Range("b" & 36)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(big_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            wordAppl.Selection.TypeParagraph
'
'            Range("a" & 30).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 31).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("b" & 1).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            Range("a" & 32).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'
'            wordAppl.Selection.TypeBackspace
'
'            small_pic = Range("b" & 37)
'            ThisWorkbook.Sheets("Sheet1").Shapes.Range(Array(small_pic)).Select
'            Selection.Copy
'            wordAppl.Selection.Paste
'            Range("b" & 35).Copy
'            wordAppl.Selection.PasteExcelTable False, False, true
'        End If
'
        
        
        'operations
        
    
        
        
        objWord.Save
        objWord.Close
    Next
    
'    For i = 0 To (num - 1)
'
'        If (hour_top + i * interval) Mod 60 = 0 Then
'            name = (hour_top + i * interval) / 60 & "点"
'        Else
'            name = (hour_top + i * interval) \ 60 & "点" & (hour_top + i * interval) Mod 60 & "分"
'        End If
'
'        .Documents.Open (path & name & ".doc")
'
'
'        Operation
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 2)
'            With wordAppl.Selection.Find
'                .Text = "mtwos"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 4)
'            With wordAppl.Selection.Find
'                .Text = "www.taobao.com"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 6)
'            With wordAppl.Selection.Find
'                .Text = "马赛克"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'
'        If iRnd(1, 2) = 1 Then
'            paste_string = Range("b" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        Else
'            paste_string = Range("c" & 7)
'            With wordAppl.Selection.Find
'                .Text = "货比三家（2-3分钟一家）"
'                .Replacement.Text = paste_string
'                .Forward = True
'                .Wrap = wdFindContinue
'                .Format = False
'                .MatchCase = False
'                .MatchWholeWord = False
'                .MatchByte = True
'                .MatchWildcards = False
'                .MatchSoundsLike = False
'                .MatchAllWordForms = False
'            End With
'            wordAppl.Selection.Find.Execute Replace:=wdReplaceAll
'        End If
'        Operation
'
'
'        .objWord.Save
'        .objWord.Close
'    Next

    '.objWord.Close False
    .Quit False
End With
    Set wordAppl = Nothing '释放存储空间
    Application.ScreenUpdating = False '关闭屏幕刷新
    Range("BA2").Select
    ActiveCell.FormulaR1C1 = "=RC[-26]*R1C8"
    Range("BA2").Select
    Selection.AutoFill Destination:=Range("BA2:BA32"), Type:=xlFillDefault
    Range("BA2:BA32").Select
    Selection.AutoFill Destination:=Range("BA2:BP32"), Type:=xlFillDefault
    Range("BA2:BP32").Select
    Range("a" & 1).Select
    
End
Exit Sub
erro3:
    Selection.Copy
    Resume
End Sub