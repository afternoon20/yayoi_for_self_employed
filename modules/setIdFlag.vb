Option Explicit

Sub setIdFlag()
   Dim allTable As ListObject: Set allTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")

   Call sortTable

   Dim row As Long: row = 2
   Dim rowIndex As Long: rowIndex = 0
   Dim slipDate As String
   While Cells(row, 4) <> ""
      If row = 2 Then
         slipDate = Cells(row, 4).Value()
         Cells(row, 1).Value() = "2110"
         rowIndex = rowIndex + 1
         GoTo Continue
      End If
      If Cells(row, 4).Value() = slipDate Then
         Cells(row, 1).Value() = "2100"
         rowIndex = rowIndex + 1
      ElseIf Cells(row, 4).Value() <> slipDate And rowIndex = 1 Then
         slipDate = Cells(row, 4).Value()
         rowIndex = 1
         Cells(row - 1, 1).Value() = "2111"
         Cells(row, 1).Value() = "2110"
      Else
         slipDate = Cells(row, 4).Value()
         rowIndex = 1
         Cells(row - 1, 1).Value() = "2101"
         Cells(row, 1).Value() = "2110"
      End If
      If Cells(row + 1, 4) = "" And rowIndex = 1 Then
         Cells(row, 1).Value() = "2111"
      ElseIf Cells(row + 1, 4).Value() = "" Then
         Cells(row, 1).Value() = "2101"
      End If
Continue:
      row = row + 1
   Wend
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

