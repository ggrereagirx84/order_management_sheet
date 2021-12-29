Attribute VB_Name = "Module2"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Sub 当日確定分()
'
' 当日確定分 Macro
'
Call ボタン13_Click
'
    ActiveSheet.Range("$A$2:$Y$301").AutoFilter Field:=2, Criteria1:= _
        xlFilterToday, Operator:=xlFilterDynamic
End Sub
Sub 全注残()
'
' 全注残 Macro
'

'
End Sub
Sub ボタン13_Click()
    
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FW As Boolean: FW = False
    
    If ORDERED_SHEET.FilterMode Then ORDERED_SHEET.ShowAllData
    Do While FW = False
        If Not ORDERED_SHEET.FilterMode Then
            FW = True
        End If
    Loop
End Sub


Sub test25()
    Dim WshNetworkObject As Object
    
    Set WshNetworkObject = CreateObject("WScript.Network")
      
    With WshNetworkObject
        MsgBox "ユーザー名： " & .UserName & vbCrLf _
             & "コンピュータ名： " & .ComputerName
    End With
    
    Set WshNetworkObject = Nothing
End Sub

Sub メール作成()


Dim objOutlook As Object
Dim mail_format As Worksheet
Dim body_text_1 As String
Dim body_text_2 As String
Dim target_sheet As Worksheet
Dim t_end_row As Integer
Dim Dic, i As Long, buf As String
Set Dic = CreateObject("Scripting.Dictionary")
Dim to_names As String
Dim name
 
Set objOutlook = CreateObject("Outlook.Application")
Set mail_format = ThisWorkbook.Sheets("メール作成")
Set target_sheet = ThisWorkbook.Sheets("修正後")

mail_format.Cells(6, 20) = "=SUBSTITUTE(RC[-17],CHAR(10),""<br>"")"
body_text_1 = mail_format.Cells(6, 20)
mail_format.Cells(6, 20) = ""

mail_format.Cells(14, 20) = "=SUBSTITUTE(RC[-17],CHAR(10),""<br>"")"
body_text_2 = mail_format.Cells(14, 20)
mail_format.Cells(14, 20) = ""

t_end_row = target_sheet.Cells(60000, 2).End(xlUp).row

target_sheet.Range(target_sheet.Cells(1, 1), target_sheet.Cells(t_end_row, 6)).Copy

For i = 2 To t_end_row
    buf = target_sheet.Cells(i, 3).Value
    If Not Dic.Exists(buf) Then
        Dic.Add buf, buf
    End If
Next i

For Each name In Dic
    to_names = to_names & name & ";"
Next
Set Dic = Nothing

With objOutlook
    With .CreateItem(0)
        .To = to_names
        .CC = mail_format.Cells(3, 3).Value
        .BCC = mail_format.Cells(4, 3).Value
        .Subject = mail_format.Cells(5, 3).Value
        .Display
        .HTMLBody = "<body>" & body_text_2 & "<br>" & "<br>" & "<br>" & "</body>" & .HTMLBody
        .GetInspector.WordEditor.Windows(1).Selection.Paste
        .HTMLBody = "<body>" & body_text_1 & "<br>" & "<br>" & "<br>" & "</body>" & .HTMLBody
        .Application.ActiveWindow.Activate
    End With
End With


End Sub

Sub 地図整理()
    Dim buf As String, msg As String
    Dim fso As Object, fl, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim test, i As Integer, st As String, k As Integer, isdone_folder_name As String, move_file_name As String
    Dim data, key, target_folder_name As String
    
    Dim Temps As Object: Set Temps = CreateObject("Scripting.Dictionary")
    
    target_folder_name = FRONT_SHEET.Cells(3, 5).Value & "\"
    isdone_folder_name = target_folder_name & "▼済み\"
    
    Set fl = fso.GetFolder(target_folder_name)
    
    Dim count As Integer
    count = 0
    For Each f In fl.Files
        If f.name Like "*pdf" Then
            Temps.Add count, f.name
            count = count + 1
        End If
    Next
    
    data = ORDERED_SHEET.Cells(1, 1).CurrentRegion
    
    For i = 2 To UBound(data, 1)
        st = "*" & data(i, 4) & " " & Left(data(i, 8), 6) & "*"
        For Each key In Temps
            If Temps(key) Like st Then
                Temps.Remove key
                'Debug.Print key
            End If
        Next
    Next
    
    'Application.DisplayAlerts = False
    'Call fso.DeleteFolder(folderPath, True)
    
    If Dir(isdone_folder_name, vbDirectory) = "" Then
        MkDir isdone_folder_name
    End If
    

    For Each key In Temps
        move_file_name = target_folder_name & Temps(key)
        fso.MoveFile move_file_name, isdone_folder_name
        'Debug.Print move_file_name
    Next

    Set fso = Nothing
    
    MsgBox "OK"

End Sub
Sub 見積整理()
    Dim buf As String, msg As String
    Dim fso As Object, fl, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim test, i As Integer, st As String, k As Integer, isdone_folder_name As String, move_file_name As String
    Dim data, key, target_folder_name As String
    
    Dim Temps As Object: Set Temps = CreateObject("Scripting.Dictionary")
    
    target_folder_name = FRONT_SHEET.Cells(3, 5).Value & "\見積\"
    isdone_folder_name = target_folder_name & "▼済み\"

    Set fl = fso.GetFolder(target_folder_name)
    
    Dim count As Integer
    count = 0
    For Each f In fl.Files
        If f.name Like "*pdf" Then
            Temps.Add count, f.name
            count = count + 1
        End If
    Next
    
    data = ORDERED_SHEET.Cells(1, 1).CurrentRegion
    Dim conversion_st As String
    
    For i = 2 To UBound(data, 1)
        st = data(i, 4) & " " & data(i, 9) & ".pdf"
        For Each key In Temps
            With WorksheetFunction
                conversion_st = .Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(Temps(key), "ｰ", "-"), "ｧ", "ｱ"), "ｨ", "ｲ"), "ｩ", "ｳ"), "ｪ", "ｴ"), "ｫ", "ｵ"), "ｬ", "ﾔ"), "ｭ", "ﾕ"), "ｮ", "ﾖ"), "ｯ", "ﾂ")
            End With
            If conversion_st Like "*pdf" And conversion_st Like st Then
                Temps.Remove key
                'Debug.Print key
            End If
        Next
    Next
    
    'Application.DisplayAlerts = False
    'Call fso.DeleteFolder(folderPath, True)
    
    If Dir(isdone_folder_name, vbDirectory) = "" Then
        MkDir isdone_folder_name
    End If
    

    For Each key In Temps
        move_file_name = target_folder_name & Temps(key)
        fso.MoveFile move_file_name, isdone_folder_name
        'Debug.Print move_file_name
    Next

    Set fso = Nothing
    
    MsgBox "OK"

End Sub


Sub 単価タム()
    Dim buf As String, msg As String
    Dim fso As Object, Target_Folder, f, Target_File
    dim custormerCode as string
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim test, i As Integer, st As String, k As Integer, isdone_folder_name As String, move_file_name As String
    Dim data, key, target_folder_name As String, ship_day_st As String, stay_flag As Boolean, clone_st As String
    
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")

    Dim F_end_row As Long: F_end_row = FRONT_SHEET.Cells(Rows.count, 2).End(xlUp).row
    
    Dim Temps As Object: Set Temps = CreateObject("Scripting.Dictionary")
    Dim filepath_for_person As Object: Set filepath_for_person = CreateObject("Scripting.Dictionary")

    For i = 3 To F_end_row
        key = FRONT_SHEET.Cells(i, 2).Value
        If key <> "" Then
            If FRONT_SHEET.Cells(i, 5).Value <> "" Then: filepath_for_person.Add key, FRONT_SHEET.Cells(i, 5).Value: Else filepath_for_person.Add key, ""
        End If
    Next
    
    target_folder_name = FRONT_SHEET.Cells(3, 5).Value & "\見積\"
    isdone_folder_name = target_folder_name & "☆単価タム\"
    
    Set Target_Folder = fso.GetFolder(target_folder_name)
    'Set Target_File = fso.GetFile(target_folder_name)

    If Dir(isdone_folder_name, vbDirectory) = "" Then
        MkDir isdone_folder_name
    End If
    
    Dim count As Integer
    count = 0
    ' For Each f In Target_Folder.Files
    '     If f.name Like "*pdf" Then
    '         Temps.Add count, f.name
    '         count = count + 1
    '     End If
    ' Next
    
    data = ORDERED_SHEET.Cells(1, 1).CurrentRegion
    Dim conversion_st As String
    
    For i = 2 To UBound(data, 1)
        If data(i, 2) = Date Then
        'If data(i, 2) = "2021/08/16" Then
            If data(i, 5) = "確定" Or data(i, 5) = "B10" Then
                custormerCode = data(i, 4)
                st = custormerCode & " " & data(i, 9) & ".pdf"
                target_folder_name = filepath_for_person.item(custormerCode) & "\見積\"
                Set Target_Folder = fso.GetFolder(target_folder_name)

                If ship_day_st = "" Then: ship_day_st = "≪" & WorksheetFunction.Substitute(Right(data(i, 17), 5), "/", "-") & "≫"
                For Each f In Target_Folder.Files
                    With WorksheetFunction
                        conversion_st = .Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(f.name, "ｰ", "-"), "ｧ", "ｱ"), "ｨ", "ｲ"), "ｩ", "ｳ"), "ｪ", "ｴ"), "ｫ", "ｵ"), "ｬ", "ﾔ"), "ｭ", "ﾕ"), "ｮ", "ﾖ"), "ｯ", "ﾂ")
                    End With
                    If conversion_st Like "*pdf" And conversion_st Like st Then
                    'If f.name Like "*pdf" And f.name Like st Then
                        clone_st = f.name
                        move_file_name = target_folder_name & f.name
                        fso.MoveFile move_file_name, isdone_folder_name
                        stay_flag = True
                        Do While stay_flag = True
                            If Not Dir(isdone_folder_name & clone_st) = "" Then
                                stay_flag = False
                            End If
                        Loop
                        Set Target_File = fso.GetFile(isdone_folder_name & clone_st)
                        Target_File.name = ship_day_st & clone_st
                        'Debug.Print f.name
                    End If
                Next
            End If
        End If
    Next
    
    'Application.DisplayAlerts = False
    'Call fso.DeleteFolder(folderPath, True)
    

    ' For Each key In Temps
    '     move_file_name = target_folder_name & Temps(key)
    '     fso.MoveFile move_file_name, isdone_folder_name
    '     'Debug.Print move_file_name
    ' Next

    Set fso = Nothing
    
    MsgBox "OK"

End Sub

Sub test()
    Dim buf As String, msg As String
    Dim fso As Object, Target_Folder, f, Target_File
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim i As Integer, st As String, k As Integer, isdone_folder_name As String, move_file_name As String
    Dim data, key, target_folder_name As String, ship_day_st As String, stay_flag As Boolean, clone_st As String
    
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    
    Dim Temps As Object: Set Temps = CreateObject("Scripting.Dictionary")
    
    target_folder_name = FRONT_SHEET.Cells(3, 5).Value & "\見積\"
    isdone_folder_name = target_folder_name & "☆単価タム\"
    
    Set Target_Folder = fso.GetFolder(target_folder_name)
    'Set Target_File = fso.GetFile(target_folder_name)

    If Dir(isdone_folder_name, vbDirectory) = "" Then
        MkDir isdone_folder_name
    End If
    
    Dim count As Integer
    count = 0
    ' For Each f In Target_Folder.Files
    '     If f.name Like "*pdf" Then
    '         Temps.Add count, f.name
    '         count = count + 1
    '     End If
    ' Next
    
    data = ORDERED_SHEET.Cells(1, 1).CurrentRegion
    
    Dim test As String
    
    
    For Each f In Target_Folder.Files
        clone_st = f.name
        Set Target_File = fso.GetFile(target_folder_name & clone_st)
        MsgBox clone_st
        test = WorksheetFunction.Substitute(clone_st, "ｰ", "-")
        Target_File.name = test
        Exit For
    Next
    
    'Application.DisplayAlerts = False
    'Call fso.DeleteFolder(folderPath, True)
    

    ' For Each key In Temps
    '     move_file_name = target_folder_name & Temps(key)
    '     fso.MoveFile move_file_name, isdone_folder_name
    '     'Debug.Print move_file_name
    ' Next

    Set fso = Nothing
    
    MsgBox "OK"


End Sub

Sub 仮()
Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
Dim F_end_row As Integer: F_end_row = FRONT_SHEET.Cells(Rows.count, 2).End(xlUp).row
Dim test As String, customerCode As String

customerCode = "43297"

With WorksheetFunction

    If .Asc(.index(FRONT_SHEET.Range(FRONT_SHEET.Cells(3, 6), FRONT_SHEET.Cells(F_end_row, 6)), .Match(customerCode, FRONT_SHEET.Range(FRONT_SHEET.Cells(3, 2), FRONT_SHEET.Cells(F_end_row, 2)), 0))) Like "*B10*" Then
        Debug.Print "OK"
    End If
    
End With

End Sub

Sub testPath()

    Dim buf As String, msg As String
    Dim fso As Object, fl, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim filePath As String
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim test, i As Integer, st As String, k As Integer, isdone_folder_name As String, move_file_name As String
    Dim data, key, target_folder_name As String
    
    Dim Temps As Object: Set Temps = CreateObject("Scripting.Dictionary")
    
    target_folder_name = "C:\Users\11047261\LIXIL\LWT西日本受注センター - 受信メール保存_ＰＵ\Ｌ_清水"
    isdone_folder_name = target_folder_name & "▼済み\"
    
    Set fl = fso.GetFolder(target_folder_name)
    
    Dim count As Integer
    count = 0
    For Each f In fl.Files
        If f.name Like "*pdf" Then
            MsgBox f.name
        End If
    Next

    

    Set fso = Nothing
    
    MsgBox "OK"




End Sub





