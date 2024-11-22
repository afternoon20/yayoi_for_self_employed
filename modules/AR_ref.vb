Option Explicit

Sub reflection()
'    NOTE: init
    Dim keyName As String
'    NOTE: リストを昇順
    Dim userTable As ListObject
    Set userTable = ActiveWorkbook.Worksheets("預金").ListObjects("預金")
    userTable.Range.AutoFilter Field:=3
    userTable.Sort.SortFields.Clear
    userTable.Sort.SortFields.Add2 key:= _
        Range("預金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With userTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'NOTE:     テーブル内の売掛金行を取得
    Dim accountsRecievalbleList As Object: Set accountsRecievalbleList = New Dictionary
    Dim row As Excel.ListRow

    For Each row In userTable.ListRows
        If row.Range.Cells(3) = "" Then
            Exit For
        End If
        If row.Range.Cells(4) <> "売掛金" Then
            GoTo Continue
        End If

'        NOTE: ARクラスの配列に設定
        Dim AR As AR: Set AR = New AR
        AR.columns("date") = row.Range.Cells(3).Value()
        AR.columns("account") = row.Range.Cells(4).Value()
        AR.columns("taxType") = row.Range.Cells(5).Value()
        AR.columns("amount") = row.Range.Cells(6).Value()
        AR.columns("withholdingTax") = row.Range.Cells(7).Value()
        AR.columns("totalAmount") = row.Range.Cells(8).Value()
        AR.columns("customer") = row.Range.Cells(9).Value()
        AR.columns("content") = row.Range.Cells(10).Value()
        AR.columns("invoiceNumber") = row.Range.Cells(12).Value()
'        NOTE: 配列に設定
        keyName = AR.columns("date") & "_預金_" & AR.columns("customer") & "_" & AR.columns("content")
'TODO:         同じキーで例外投げる
        Call accountsRecievalbleList.Add(keyName, AR)
Continue:
    Next

'    NOTE: 売掛金シートに反映
'    NOTE: リストを昇順
    Dim arTable As ListObject
    Set arTable = ActiveWorkbook.Worksheets("売掛金").ListObjects("売掛金")
    arTable.Range.AutoFilter Field:=3
    arTable.Sort.SortFields.Clear
    arTable.Sort.SortFields.Add2 key:= _
        Range("売掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With arTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For Each row In arTable.ListRows
        keyName = row.Range.Cells(3).Value() & "_" & row.Range.Cells(4).Value() & "_" & row.Range.Cells(9).Value() & "_" & row.Range.Cells(10).Value()
        If (accountsRecievalbleList.Exists(keyName)) Then
            accountsRecievalbleList.Remove (keyName)
        End If
        If row.Range.Cells(3) = "" Then
            Dim blankRange As Range: Set blankRange = row.Range
            Dim r As Long: r = blankRange.row
            Dim accountsRecievalbleListKey As Variant
            Sheets("売掛金").Activate
            For Each accountsRecievalbleListKey In accountsRecievalbleList.Keys
                Set AR = accountsRecievalbleList(accountsRecievalbleListKey)
                Cells(r, 3).Value = AR.columns("date")
                Cells(r, 4).Value = "預金"
                Cells(r, 5).Value = AR.columns("taxType")
                Cells(r, 6).Value = AR.columns("amount") - (AR.columns("amount") * 2)
                Cells(r, 7).Value = AR.columns("withholdingTax")
                Cells(r, 9).Value = AR.columns("customer")
                Cells(r, 10).Value = AR.columns("content")
                Cells(r, 12).Value = AR.columns("invoiceNumber")
                r = r + 1
            Next
            Exit For
        End If
    Next
'  NOTE: 日付の昇順に並び替え
    arTable.Range.AutoFilter Field:=3
    arTable.Sort.SortFields.Clear
    arTable.Sort.SortFields.Add2 key:= _
        Range("売掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With arTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
