Option Explicit

Public columns As Object

Private Sub Class_Initialize()
    Set columns = New Dictionary
    Call Me.columns.Add("idFlag", "") '識別フラグ
    Call Me.columns.Add("slipNum", "") '伝票No(管理用)
    Call Me.columns.Add("financStat", "") '決算
    Call Me.columns.Add("slipDay", "") '日付
    Call Me.columns.Add("debitName", "") '借方勘定科目
    Call Me.columns.Add("debitSub", "") '借方補助科目
    Call Me.columns.Add("debitDep", "") '借方部門
    Call Me.columns.Add("debitTaxType", "対象外") '借方税区分
    Call Me.columns.Add("debitAmo", "") '借方金額
    Call Me.columns.Add("debitTax", "") '借方税金額
    Call Me.columns.Add("creditName", "") '貸方勘定科目
    Call Me.columns.Add("creditSub", "") '貸方補助科目
    Call Me.columns.Add("creditDep", "") '貸方部門
    Call Me.columns.Add("creditTaxType", "対象外") '貸方税区分
    Call Me.columns.Add("creditAmo", "") '貸方金額
    Call Me.columns.Add("creditTax", "") '貸方税金額
    Call Me.columns.Add("summary", "") '摘要
    Call Me.columns.Add("num", "") '番号
    Call Me.columns.Add("setlement", "") '期日
    Call Me.columns.Add("slipType", 3) 'タイプ（仕訳データの場合は「0」、振伝は「3」）
    Call Me.columns.Add("origin", "") '生成元
    Call Me.columns.Add("memo", "") '仕訳メモ
    Call Me.columns.Add("tag1", "") '付箋1
    Call Me.columns.Add("tag2", "") '付箋2
    Call Me.columns.Add("adjustment", "no") '調整（noと記入
End Sub

Private Sub Class_Terminate()
End Sub

Public Sub setData(params As Object)
    Dim key As Variant
    For Each key In params
        If Me.columns.Exists(key) Then
            Me.columns(key) = params(key)
        End If
    Next
End Sub

