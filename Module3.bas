Attribute VB_Name = "Module3"
Sub �󗓂Ɗm�F���̂�()
Attribute �󗓂Ɗm�F���̂�.VB_Description = "�S���c�����ԗ����󗓂Ɗm�F���݂̂�\��"
Attribute �󗓂Ɗm�F���̂�.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �󗓂Ɗm�F���̂� Macro
' �S���c�����ԗ����󗓂Ɗm�F���݂̂�\��
'

'
    ActiveSheet.Range("$A$2:$W$387").AutoFilter Field:=5, Criteria1:="=�m�F��", _
        Operator:=xlOr, Criteria2:="="
End Sub
