Option Explicit

Public columns As Object

Private Sub Class_Initialize()
    Set columns = New Dictionary
    Call Me.columns.Add("date", "") '日付
    Call Me.columns.Add("account", "") '勘定科目
    Call Me.columns.Add("taxType", "") '消費税区分
    Call Me.columns.Add("amount", "") '税込金額
    Call Me.columns.Add("withholdingTax", "") '源泉徴収税額
    Call Me.columns.Add("totalAmount", "") '合計金額
    Call Me.columns.Add("customer", "") '取引先
    Call Me.columns.Add("content", "") '取引内容
    Call Me.columns.Add("invoiceNumber", "") 'インボイス登録番号
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
