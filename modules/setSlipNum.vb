Option Explicit

Sub setSlipNum()
    Dim allTable As ListObject: Set allTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")

    Call sortTable

   Dim slipNum As Long: slipNum = 0
   Dim slipDate As String
   Dim row As Excel.ListRow
   For Each row In allTable.ListRows
      If row.Range.Cells(4) = "" Then
            Exit For
      End If
      If slipNum = 0 Then
         slipNum = slipNum + 1
         slipDate = row.Range.Cells(4).Value()
         row.Range.Cells(2).Value() = slipNum
         GoTo Continue
      End If
      If row.Range(4).Value() = slipDate Then
         row.Range.Cells(2).Value() = slipNum
      Else
         slipNum = slipNum + 1
         slipDate = row.Range.Cells(4).Value()
         row.Range.Cells(2).Value() = slipNum
      End If
Continue:
   Next
End Sub

Private Sub sortTable()
    Dim bankTable As ListObject
    Dim cashTable As ListObject
    Dim creditTable As ListObject
    Dim receivableTable As ListObject
    Dim payableTable As ListObject
    Set bankTable = ActiveWorkbook.Worksheets("預金").ListObjects("預金")
    Set cashTable = ActiveWorkbook.Worksheets("現金").ListObjects("現金")
    Set creditTable = ActiveWorkbook.Worksheets("クレジットカード").ListObjects("クレジットカード")
    Set receivableTable = ActiveWorkbook.Worksheets("売掛金").ListObjects("売掛金")
    Set payableTable = ActiveWorkbook.Worksheets("買掛金").ListObjects("買掛金")

    bankTable.Range.AutoFilter Field:=3
    bankTable.Sort.SortFields.Clear
    bankTable.Sort.SortFields.Add2 key:= _
       Range("預金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
       DataOption:=xlSortNormal
    With bankTable.Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With

    cashTable.Range.AutoFilter Field:=3
    cashTable.Sort.SortFields.Clear
    cashTable.Sort.SortFields.Add2 key:= _
       Range("現金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
       DataOption:=xlSortNormal
    With cashTable.Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With

    creditTable.Range.AutoFilter Field:=3
    creditTable.Sort.SortFields.Clear
    creditTable.Sort.SortFields.Add2 key:= _
       Range("クレジットカード[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
       DataOption:=xlSortNormal
    With creditTable.Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With

    receivableTable.Range.AutoFilter Field:=3
    receivableTable.Sort.SortFields.Clear
    receivableTable.Sort.SortFields.Add2 key:= _
       Range("売掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
       DataOption:=xlSortNormal
    With receivableTable.Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With

    payableTable.Range.AutoFilter Field:=3
    payableTable.Sort.SortFields.Clear
    payableTable.Sort.SortFields.Add2 key:= _
       Range("買掛金[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
       DataOption:=xlSortNormal
    With payableTable.Sort
       .Header = xlYes
       .MatchCase = False
       .Orientation = xlTopToBottom
       .SortMethod = xlPinYin
       .Apply
    End With
End Sub

