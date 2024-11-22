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
'NOTE:     テーブル内の買掛金行を取得
    Dim accountsPayableList As Object: Set accountsPayableList = New Dictionary
    Dim row As Excel.ListRow

    For Each row In userTable.ListRows
        If row.Range.Cells(3) = "" Then
            Exit For
        End If
        If row.Range.Cells(4) <> "買掛金" Then
            GoTo Continue
        End If

'        NOTE: ARクラスの配列に設定
        Dim AP As AP: Set AP = New AP
        AP.columns("date") = row.Range.Cells(3).Value()
        AP.columns("account") = row.Range.Cells(4).Value()
        AP.columns("taxType") = row.Range.Cells(5).Value()
        AP.columns("amount") = row.Range.Cells(6).Value()
        AP.columns("withholdingTax") = row.Range.Cells(7).Value()
        AP.columns("totalAmount") = row.Range.Cells(8).Value()
        AP.columns("customer") = row.Range.Cells(9).Value()
        AP.columns("content") = row.Range.Cells(10).Value()
        AP.columns("invoiceNumber") = row.Range.Cells(12).Value()
'        NOTE: 配列に設定
        keyName = AP.columns("date") & "_預金_" & AP.columns("customer") & "_" & AP.columns("content")
'TODO:         同じキーで例外投げる
        Call accountsPayableList.Add(keyName, AP)
Continue:
    Next

'    NOTE: 買掛金シートに反映
'    NOTE: リストを昇順
    Dim apTable As ListObject
    Set apTable = ActiveWorkbook.Worksheets("買掛金").ListObjects("買掛金")
    apTable.Range.AutoFilter Field:=3
    apTable.Sort.SortFields.Clear
    apTable.Sort.SortFields.Add2 key:= _
        Range("買掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With apTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    For Each row In apTable.ListRows
        keyName = row.Range.Cells(3).Value() & "_" & row.Range.Cells(4).Value() & "_" & row.Range.Cells(9).Value() & "_" & row.Range.Cells(10).Value()
        If (accountsPayableList.Exists(keyName)) Then
            accountsPayableList.Remove (keyName)
        End If
        If row.Range.Cells(3) = "" Then
            Dim blankRange As Range: Set blankRange = row.Range
            Dim r As Long: r = blankRange.row
            Dim accountsPayableListKey As Variant
            Sheets("買掛金").Activate
            For Each accountsPayableListKey In accountsPayableList.Keys
                Set AP = accountsPayableList(accountsPayableListKey)
                Cells(r, 3).Value = AP.columns("date")
                Cells(r, 4).Value = "預金"
                Cells(r, 5).Value = AP.columns("taxType")
                Cells(r, 6).Value = AP.columns("amount")
                Cells(r, 7).Value = AP.columns("withholdingTax")
                Cells(r, 9).Value = AP.columns("customer")
                Cells(r, 10).Value = AP.columns("content")
                Cells(r, 12).Value = AP.columns("invoiceNumber")
            Next
            Exit For
        End If
    Next
'  NOTE: 日付の昇順に並び替え
    apTable.Range.AutoFilter Field:=3
    apTable.Sort.SortFields.Clear
    apTable.Sort.SortFields.Add2 key:= _
        Range("買掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With apTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
