Attribute VB_Name = "Module3"
Sub 空欄と確認中のみ()
Attribute 空欄と確認中のみ.VB_Description = "全注残から状態欄が空欄と確認中のみを表示"
Attribute 空欄と確認中のみ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 空欄と確認中のみ Macro
' 全注残から状態欄が空欄と確認中のみを表示
'

'
    ActiveSheet.Range("$A$2:$W$387").AutoFilter Field:=5, Criteria1:="=確認中", _
        Operator:=xlOr, Criteria2:="="
End Sub
