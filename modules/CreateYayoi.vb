Option Explicit

Sub output()
    ' NOTE: macだと不具合でるのでコメントブロック
    ' Application.ScreenUpdating = False
    
    Dim allTable As ListObject
    Set allTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")
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

    Sheets("全取引").Copy
    Dim newTable As ListObject: Set newTable = ActiveWorkbook.Worksheets("全取引").ListObjects("All")
    newTable.Unlist
        Cells.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    columns("A:B").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Rows("1:1").Select
    Selection.Delete Shift:=xlUp

    Range("A1").Select
    Selection.End(xlDown).Select
    Dim lastBlankRow As String
    lastBlankRow = Str(Selection.row() + 1) & ":1048537"
    Rows(lastBlankRow).Select
    Selection.Delete Shift:=xlUp

    ActiveWindow.Zoom = 100
    Range("A1").Select
    ' NOTE: macだと不具合でるのでコメントブロック
    ' Application.ScreenUpdating = True
End Sub

