Option Explicit

Sub reflection()
    Debug.Print "全取引 start:" & Now
    Dim startDate As Date
    startDate = Now
    Application.ScreenUpdating = False

    Call AP_ref.reflection
    Call AR_ref.reflection
    
    Dim allTable As ListObject
    Set allTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")
    
    Call sortTable
    
    Dim allSlipList As Object: Set allSlipList = getSlips()
    
    Dim allSlipListKey As Variant
    Dim yayoi As yayoi
    Dim x As Long: x = 1
    Dim isFirstLoop As Boolean: isFirstLoop = True
    Dim yayoiCount As Long: yayoiCount = allSlipList.Count
    Debug.Print "弥生行数:" & yayoiCount
    Dim rangeStr As String: rangeStr = "A2:Y" & allSlipList.Count
    ReDim tableArray(1 To yayoiCount, 1 To 25) As Variant
    For Each allSlipListKey In allSlipList.Keys
        Set yayoi = allSlipList(allSlipListKey)
        tableArray(x, 1) = "=IF(D" & x + 1 & "<>D" & x & ", IF(D" & x + 1 & "<>D" & x + 2 & ", 2111, 2110), IF(D" & x + 1 & "<>D" & x + 2 & ", 2101, 2100))"
        tableArray(x, 2) = "=IF(COUNTIF($D$2:D2, D2)=1, MAX($B$" & x & ":B1)+1, INDEX($B$2:B2, MATCH(D2, $D$2:D2, 0)))"
        tableArray(x, 3) = yayoi.columns("financStat")
        tableArray(x, 4) = yayoi.columns("slipDay")
        tableArray(x, 5) = yayoi.columns("debitName")
        tableArray(x, 6) = yayoi.columns("debitSub")
        tableArray(x, 7) = yayoi.columns("debitDep")
        tableArray(x, 8) = yayoi.columns("debitTaxType")
        tableArray(x, 9) = yayoi.columns("debitAmo")
        tableArray(x, 10) = yayoi.columns("debitTax")
        tableArray(x, 11) = yayoi.columns("creditName")
        tableArray(x, 12) = yayoi.columns("creditSub")
        tableArray(x, 13) = yayoi.columns("creditDep")
        tableArray(x, 14) = yayoi.columns("creditTaxType")
        tableArray(x, 15) = yayoi.columns("creditAmo")
        tableArray(x, 16) = yayoi.columns("creditTax")
        tableArray(x, 17) = yayoi.columns("summary")
        tableArray(x, 18) = yayoi.columns("num")
        tableArray(x, 19) = yayoi.columns("settlement")
        tableArray(x, 20) = yayoi.columns("slipType")
        tableArray(x, 21) = yayoi.columns("origin")
        tableArray(x, 22) = yayoi.columns("memo")
        tableArray(x, 23) = yayoi.columns("tag1")
        tableArray(x, 24) = yayoi.columns("tag2")
        tableArray(x, 25) = yayoi.columns("adjustment")
        x = x + 1
    Next

    With allTable
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
        Dim newRange As Range
        Set newRange = .HeaderRowRange.Resize(yayoiCount)
        .Resize newRange
        .DataBodyRange.Value = tableArray
    End With
    Sheets("全取引").Activate

    allTable.Range.AutoFilter Field:=3
    allTable.Sort.SortFields.Clear
    allTable.Sort.SortFields.Add2 key:= _
        Range("All[[#All],[日付]]"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With allTable.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Sheets("全取引").Activate
    Application.ScreenUpdating = True
    
    Dim endDate As Date
    endDate = Now
    Debug.Print "実行時間:" & DateDiff("s", startDate, endDate) & "秒"
    Debug.Print "全取引 end:" & Now
    
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

Function getSlips() As Object
    Set getSlips = New Dictionary
    Dim keyIndex As Long: keyIndex = 1
    Dim yayoi As yayoi
    
    Dim bankTable As ListObject
    Dim cashTable As ListObject
    Dim creditTable As ListObject
    Dim receivableTable As ListObject
    Dim payableTable As ListObject
    Dim allTable As ListObject
    Set bankTable = ActiveWorkbook.Worksheets("預金").ListObjects("預金")
    Set cashTable = ActiveWorkbook.Worksheets("現金").ListObjects("現金")
    Set creditTable = ActiveWorkbook.Worksheets("クレジットカード").ListObjects("クレジットカード")
    Set receivableTable = ActiveWorkbook.Worksheets("売掛金").ListObjects("売掛金")
    Set payableTable = ActiveWorkbook.Worksheets("買掛金").ListObjects("買掛金")
    Set allTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")
    
    Call createYayoiByTable(getSlips, keyIndex, bankTable)
    Call createYayoiByTable(getSlips, keyIndex, cashTable)
    Call createYayoiByTable(getSlips, keyIndex, creditTable)
    Call createYayoiByTable(getSlips, keyIndex, receivableTable, True)
    Call createYayoiByTable(getSlips, keyIndex, payableTable, True, True)
    
End Function

Private Sub createYayoiByTable(ByRef YayoiArray As Object, ByRef keyIndex As Long, table As ListObject, Optional excludeYokin As Boolean = False, Optional isCreditTable As Boolean = False)
    Dim yayoi As yayoi
    Dim row As Excel.ListRow
    Dim absoluteAmount As Double
    Dim taxType As String
    Dim containWithholdingTax As Boolean
    Dim leftName As String: leftName = "debit"
    Dim rightName As String: rightName = "credit"
    
    For Each row In table.ListRows
        If row.Range.Cells(3) = "" Then
            Exit For
        End If
        If excludeYokin And row.Range.Cells(4) = "預金" Then
            GoTo Continue
        End If
        
        Set yayoi = New yayoi
        absoluteAmount = Abs(row.Range.Cells(6).Value())
        containWithholdingTax = False
        leftName = "debit"
        rightName = "credit"
        If (isCreditTable = False And row.Range.Cells(8).Value() < 0) Or (isCreditTable = True And row.Range.Cells(8).Value() > 0) Then
            leftName = "credit"
            rightName = "debit"
        End If
        If row.Range.Cells(7).Value() > 0 Then
            containWithholdingTax = True
        End If
        yayoi.columns("slipDay") = row.Range.Cells(3).Value()
        yayoi.columns("summary") = row.Range.Cells(11).Value()
        yayoi.columns(leftName & "Name") = row.Range.Cells(1).Value()
        ' NOTE:源泉徴収があればそれを差し引いた額を設定
        If containWithholdingTax Then
            yayoi.columns(leftName & "Amo") = absoluteAmount - row.Range.Cells(7).Value()
        Else
            yayoi.columns(leftName & "Amo") = absoluteAmount
        End If
        yayoi.columns(rightName & "Name") = row.Range.Cells(4).Value()
        yayoi.columns(rightName & "TaxType") = setTaxType(row.Range.Cells(4).Value(), row.Range.Cells(5).Value())
        yayoi.columns(rightName & "Tax") = setTaxAmo(absoluteAmount, row.Range.Cells(5).Value())
        yayoi.columns(rightName & "Amo") = absoluteAmount
        Call YayoiArray.Add(keyIndex, yayoi)
        keyIndex = keyIndex + 1

        ' NOTE:源泉徴収の仕訳設定
        If containWithholdingTax Then
            Set yayoi = New yayoi
            yayoi.columns("slipDay") = row.Range.Cells(3).Value()
            yayoi.columns("summary") = row.Range.Cells(11).Value() & "_源泉徴収"
            ' NOTE:買掛金に付随する取引は預り金として処理する
            If isCreditTable = True Or row.Range.Cells(4).Value() = "買掛金" Then
                yayoi.columns(leftName & "Name") = "預り金"
            Else
                yayoi.columns(leftName & "Name") = "仮払源泉税"
            End If
            yayoi.columns(leftName & "Amo") = Abs(row.Range.Cells(7).Value())
            yayoi.columns(rightName & "Amo") = 0
            Call YayoiArray.Add(keyIndex, yayoi)
            keyIndex = keyIndex + 1
        End If
Continue:
    Next
End Sub

Function setTaxType(accountName As String, taxType As String) As String
    'NOTE: 売上高の科目設定
    Dim ARnameList As Object
    Set ARnameList = New Dictionary
    Call ARnameList.Add("売上高", "売上高")
    Call ARnameList.Add("雑収入", "雑収入")
    
    'NOTE: 仕入れか売上高か設定
    Dim shiireUriageName As String: shiireUriageName = "課対仕入内"
    If (ARnameList.Exists(accountName)) Then
        shiireUriageName = "課税売上込"
    End If
    
    'NOTE: パーセント設定
    Dim taxRate As String: taxRate = "10%"
    If taxType Like "*8%" Then
        taxRate = "8%"
    End If
    If taxType Like "*軽減税率*" Then
        taxRate = "軽減8%"
    End If
    
    'NOTE: 経過措置設定
    Dim keikasochi As String: keikasochi = ""
    If taxType Like "*経過*" Then
        If taxType Like "*50*" Then
            keikasochi = "区分50%"
        ElseIf taxType Like "*80*" Then
            keikasochi = "区分80%"
        ElseIf taxType Like "*100*" Then
            keikasochi = "区分100%"
        ElseIf taxType Like "*不可*" Then
            keikasochi = "区分控不"
        End If
    Else
        If (ARnameList.Exists(accountName)) Then
            keikasochi = ""
        Else
            keikasochi = "適格"
        End If
        
    End If
    
    If taxType = "" Then
        setTaxType = "対象外"
    Else
        setTaxType = shiireUriageName & taxRate & keikasochi
    End If
End Function

Function setTaxAmo(totalAmo As Double, taxType As String) As String
    Dim taxRate As Double: taxRate = 0.1
    Dim keikaTaxRate As Double: keikaTaxRate = 1
    'NOTE: パーセント設定
    If taxType Like "*8*" Then
        taxRate = 0.08
    End If
    
    'NOTE: 経過措置設定
    If taxType Like "*経過*" Then
        If taxType Like "*50*" Then
            keikaTaxRate = 0.5
        ElseIf taxType Like "*80*" Then
            keikaTaxRate = 0.8
        End If
    End If
    
    If taxType = "" Then
        setTaxAmo = ""
    Else
        setTaxAmo = Str(Round((totalAmo / (1 + taxRate) * taxRate) * keikaTaxRate, 0))
    End If
End Function

