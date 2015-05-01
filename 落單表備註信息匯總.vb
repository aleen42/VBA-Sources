Sub 宏i()   '提取信息并拷贝公共信息

With Sheets("出货记录")

    
    .Range("B1").Select
    Selection.TextToColumns Destination:=Range("V1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, TAB:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=True, OtherChar:= _
        "，", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 2), Array(4, 1), Array(5, 1), _
        Array(6, 1), Array(7, 1), Array(8, 1)), TrailingMinusNumbers:=True
    
    
    
    .Cells(1, .Columns.Count).End(xlToLeft).Select  '选择第一行最后一个
    
    Selection.TextToColumns Destination:=Range("V2"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, TAB:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="/", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1)), _
        TrailingMinusNumbers:=True
        
    .Cells(1, .Columns.Count).End(xlToLeft).Select
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
    Call 宏j
    
    Range("b3").Select
End Sub
Sub 宏j()
    Dim i, num As Integer 'i计数,num代表总箱数
    Dim signal As Boolean '判断是否送货上门
    signal = False
With Sheets("出货记录")
    '拷贝公共信息
    .Cells(2, .Columns.Count).End(xlToLeft).Select
     If Right(Selection, 1) = "门" Then
        signal = True
        Selection.ClearContents
     End If
    num = .Cells(2, .Columns.Count).End(xlToLeft)
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
    
    .Cells(2, .Columns.Count).End(xlToLeft).Select '箱數
    Selection.Copy
    .Range(Cells(3, "N"), Cells(3 + num - 1, "N")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.Count).End(xlToLeft).Select
    Selection.ClearContents
    
    
    .Cells(2, .Columns.Count).End(xlToLeft).Select '物流費用
    Selection.Copy
    .Range(Cells(3, "Q"), Cells(3 + num - 1, "Q")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.Count).End(xlToLeft).Select
    Selection.ClearContents
    
    .Cells(2, .Columns.Count).End(xlToLeft).Select '付款方式
    Selection.Copy
    .Range(Cells(3, "P"), Cells(3 + num - 1, "P")).Select
    ActiveSheet.Paste
    .Cells(2, .Columns.Count).End(xlToLeft).Select
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
    .Cells(1, .Columns.Count).End(xlToLeft).Select '網點
    If Right(Selection, 1) <> "区" And Right(Selection, 1) <> "县" And Right(Selection, 1) <> "市" And Right(Selection, 1) <> "省" Then
        Selection.Copy
        .Range(Cells(3, "H"), Cells(3 + num - 1, "H")).Select
        ActiveSheet.Paste
        .Cells(1, .Columns.Count).End(xlToLeft).Select
        Selection.ClearContents
    End If
    
    '拷贝不同的信息
    .Cells(1, .Columns.Count).End(xlToLeft).Select '縣/區/市
    If Right(Selection, 1) = "县" Or Right(Selection, 1) = "区" Or Right(Selection, 1) = "市" Then    '判断最右一个字符是否含区或县
        Selection.Copy
        .Range(Cells(3, "D"), Cells(3 + num - 1, "D")).Select
        ActiveSheet.Paste
    End If
    
    .Cells(2, .Columns.Count).End(xlToLeft).Select  '快车/慢车
    If Right(Selection, 1) = "车" Then    '判断最右一个字符是否含车,德邦
        Selection.Copy
        .Range(Cells(3, "O"), Cells(3 + num - 1, "O")).Select
        ActiveSheet.Paste
        .Range("G3") = "德邦"
        .Range("G3").Select
        Selection.Copy
        .Range(Cells(3, "G"), Cells(3 + num - 1, "G")).Select
        ActiveSheet.Paste
        .Cells(2, .Columns.Count).End(xlToLeft).Select
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
End With
End Sub

    


