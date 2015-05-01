Sub 宏1()
Dim num As String
num = Sheets("Sheet1").Range("g" & 60)
    With ActiveSheet.QueryTables.Add(Connection:= _
        "URL;http://trade.taobao.com/trade/memo/update_sell_memo.htm?spm=a1z09.1.11.16.K3i2el&seller_id=768112611&biz_order_id=" & num & "&user_type=1&page_num=null&auction_title=&biz_order_time_begin=&biz_order_time_end=&comment_status=&buye" _
        , Destination:=Range("$A$1"))
        .Name = _
        "update_sell_memo.htm?spm=a1z09.1.11.16.K3i2el&seller_id=768112611&biz_order_id=" & num & "&user_type=1&page_num=null&auction_title=&biz_order_time_begin=&biz_order_time_end=&comment_status=&buye_1"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
End Sub


