Attribute VB_Name = "Module1"
Option Explicit
#If VBA7 Then
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

Sub 担当店抽出()
    Application.ScreenUpdating = False
    Dim email_subject As String: email_subject = "*【見張る君】「西日本受注C 注残現場(20時時点)*"
    Dim sub_name As String: sub_name = "担当店抽出"
    Const data_max_col As Integer = 20

    Call fetch_data(email_subject, sub_name, data_max_col)

    Application.ScreenUpdating = True

    MsgBox "OK"
End Sub
Sub 確定作業確認()
    Application.ScreenUpdating = False
    Dim email_subject As String: email_subject = "*【見張る君】「注残確認用*"
    Dim sub_name As String: sub_name = "確定作業確認"
    Const data_max_col As Integer = 21

    Call fetch_data(email_subject, sub_name, data_max_col)

    Application.ScreenUpdating = True

    MsgBox "OK"
End Sub

Function fetch_data(email_subject As String, sub_name As String, data_max_col As Integer)
    Dim customerCode
    Dim InboxFolder, wsh As Object, fso As Object, path1 As String
    Dim myNameSpace, objmailItem, propertyName As String, consecutiveNum As String, con_key As String
    Dim outlookObj As Outlook.Application
    Dim file As String, folderPath As String, fileName As String, name As String, slipNum As String, floorNum As String, Folder As String
    Dim x As Workbook: Set x = ActiveWorkbook
    Dim data()
    Dim i As Long, j As Long, st As String
    Dim count As Long
    
    '========================================================
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
'========================================================
    
    'エラーを無視することで添付のない転送メールで止まることを回避
    On Error Resume Next
    
    Dim HOLIDAY_SHEET As Worksheet: Set HOLIDAY_SHEET = ActiveWorkbook.Sheets("2021 3月期 休日")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim CUSTOMER_SHEET As Worksheet: Set CUSTOMER_SHEET = ActiveWorkbook.Sheets("店用出力")
    
    Dim F_end_row As Long: F_end_row = FRONT_SHEET.Cells(Rows.count, 2).End(xlUp).row
    
    '状態を格納するオブジェクト
    Dim Status As Object: Set Status = CreateObject("Scripting.Dictionary")
    Dim Vehicle_type As Object: Set Vehicle_type = CreateObject("Scripting.Dictionary")
    Dim DCS_message As Object: Set DCS_message = CreateObject("Scripting.Dictionary")
    Dim Remarks As Object: Set Remarks = CreateObject("Scripting.Dictionary")
    Dim Construction As Object: Set Construction = CreateObject("Scripting.Dictionary")
    Dim CM As Object: Set CM = CreateObject("Scripting.Dictionary")
    Dim RM As Object: Set RM = CreateObject("Scripting.Dictionary")


    Vehicle_type.Add "A", "トレーラ": Vehicle_type.Add "B", "11t": Vehicle_type.Add "C", "4t": Vehicle_type.Add "D", "指定なし": Vehicle_type.Add "E", "2t": Vehicle_type.Add "F", "4tｼｮｰﾄ": Vehicle_type.Add "G", "4tﾕﾆｯｸ": Vehicle_type.Add "H", "2tﾕﾆｯｸ": Vehicle_type.Add "I", "2tｼｮｰﾄ": Vehicle_type.Add "J", "11tﾕﾆｯｸ": Vehicle_type.Add "K", "2t単独車": Vehicle_type.Add "L", "2tｼｮｰﾄ単独車": Vehicle_type.Add "M", "軽トラ": Vehicle_type.Add "N", "11t平ﾎﾞﾃﾞｨ": Vehicle_type.Add "P", "4t平ﾎﾞﾃﾞｨ": Vehicle_type.Add "Z", "その他":
    
    'ソートのクリア
    Dim FW As Boolean: FW = False
    If ORDERED_SHEET.FilterMode Then ORDERED_SHEET.ShowAllData
    Do While FW = False
        If Not ORDERED_SHEET.FilterMode Then
            FW = True
        End If
    Loop
    
    '注残現場一覧の最上伝票位置
    Dim O_start_row As Integer: O_start_row = 3
    '注残現場一覧の最下伝票位置
    Dim O_end_row As Long: O_end_row = ORDERED_SHEET.Cells(Rows.count, 7).End(xlUp).row
    '注残現場一覧の最右伝票位置
    Dim O_end_col As Integer: O_end_col = 25

    Dim obj_max_count As Integer: obj_max_count = O_end_col
    
    '特別対応の最上伝票位置
    Dim S_start_row As Integer: S_start_row = 7
    '特別対応の最下伝票位置
    Dim S_end_row As Long: S_end_row = FRONT_SHEET.Cells(Rows.count, 10).End(xlUp).row
    

    
    '確定状態を連番と共に格納(連番:伝№＆管理№＆フロア)
    Dim key As String
    For i = O_start_row To O_end_row
        If ORDERED_SHEET.Cells(i, 7).Value <> 0 Then
            key = ORDERED_SHEET.Cells(i, 7).Value & ORDERED_SHEET.Cells(i, 8).Value & ORDERED_SHEET.Cells(i, 11).Value
            Status.Add key, ORDERED_SHEET.Cells(i, 5).Value
            DCS_message.Add key, ORDERED_SHEET.Cells(i, 22).Value
            Remarks.Add key, ORDERED_SHEET.Cells(i, 23).Value
            Construction.Add key, WorksheetFunction.Substitute(ORDERED_SHEET.Cells(i, 14).Value, "(変更済み)", "")
        End If
    Next

    For i = 3 To F_end_row
        key = FRONT_SHEET.Cells(i, 2).Value
        If key <> "" Then
            If FRONT_SHEET.Cells(i, 3).Value <> "" Then: CM.Add key, FRONT_SHEET.Cells(i, 3).Value
            If FRONT_SHEET.Cells(i, 4).Value <> "" Then: RM.Add key, FRONT_SHEET.Cells(i, 4).Value: Else RM.Add key, ""
        End If
    Next

    
    Sleep 20
    
    '========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "現状保存まで" & processTime
'========================================================
    
        
'*****************************************************************
'デスクトップのアドレスをを取得
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    path1 = wsh.specialFolders("MyDocuments")
    Folder = "見張る君"
    fileName = Folder & ".csv"
    file = path1 & "\" & Folder & "\" & fileName
    Set outlookObj = CreateObject("Outlook.Application")
    Set myNameSpace = outlookObj.GetNamespace("MAPI")
    Set InboxFolder = myNameSpace.GetDefaultFolder(6)
'******************************************************************
'メールを検索してデスクトップへ保存
    Sleep 20
    i = 0
    For Each objmailItem In InboxFolder.Items
    i = i + 1
    'For i = 1 To i
        'Set objmailItem = InboxFolder.Items(i)
        If objmailItem.Subject Like email_subject Then
            folderPath = path1 & "\" & Folder
            If Dir(folderPath, vbDirectory) = "" Then
                MkDir folderPath
            Else
                If Not Dir(file) = "" Then
                    Exit For
                End If
            End If

            objmailItem.Attachments.Item(1).SaveAsFile file
          
        End If
    Next
    
    
    Sleep 10
    
    If Dir(file) = "" Then
        MsgBox "受信BOXに【見張る君】が見つかりませんでした。", Buttons:=vbCritical
        Exit Function
    End If
'*******************************************************************

'========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "CSV保存まで" & processTime
'========================================================


'ファイルを開く
    i = 0
    
    Dim buf As String, flag As Boolean, A As Variant, B As Variant
    Open file For Input As #1
        Do Until EOF(1)
            ReDim Preserve data(obj_max_count - 1, i)
            Line Input #1, buf
            B = Split(buf, ",")
            customerCode = Left(Replace(B(0), """", ""), 5)
            'If 0 < WorksheetFunction.CountIf(FRONT_SHEET.Range(FRONT_SHEET.Cells(3, 2), FRONT_SHEET.Cells(F_end_row, 2)), customerCode) Or customerCode = "店コード" Then
            If CM.Exists(customerCode) = True Then
                If UBound(B) > data_max_col Then
                    j = 0
                    count = 0
                    Do
                        If Right(B(j), 1) <> """" Then
                            flag = False
                            Do While flag = False
                                st = B(j - count) & "," & B(j + 1)
                                count = count + 1
                                j = j + 1
                                If Right(st, 1) = """" Then
                                    flag = True
                                End If
                            Loop
                            'A(j - count) = st
                            'data(j - count, i) = Replace(st, """", "")
                            Call conversion_string(HOLIDAY_SHEET, data, j - count, st, i, Vehicle_type, Status, DCS_message, RM, CM, sub_name, Remarks, Construction)
                        Else
                            'A(j - count) = B(j)
                            'data(j - count, i) = Replace(B(j), """", "")
                            Call conversion_string(HOLIDAY_SHEET, data, j - count, B(j), i, Vehicle_type, Status, DCS_message, RM, CM, sub_name, Remarks, Construction)
                        End If
                        j = j + 1
                    Loop While j <= UBound(B)
                Else
                    'A = B
                    For j = 0 To UBound(B)
                        'data(j, i) = Replace(B(j), """", "")
                        Call conversion_string(HOLIDAY_SHEET, data, j, B(j), i, Vehicle_type, Status, DCS_message, RM, CM, sub_name, Remarks, Construction)
                    Next j
                End If
                ' For j = 0 To UBound(A)
                '     data(j, i) = Replace(A(j), """", "")
                ' Next j
                i = i + 1
            End If
        Loop
    Close #1


'****************************************************************

'========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "ファイル開くまで" & processTime
'========================================================
    
    Sleep 10
    
    
    Set outlookObj = Nothing
    Set wsh = Nothing
    
    
    Sleep 10
    
    ORDERED_SHEET.Range(ORDERED_SHEET.Cells(O_start_row, 1), ORDERED_SHEET.Cells(O_end_row, O_end_col)).Borders.LineStyle = False
    ORDERED_SHEET.Range(ORDERED_SHEET.Cells(O_start_row, 1), ORDERED_SHEET.Cells(O_end_row, O_end_col)) = ""

    

    ORDERED_SHEET.Range(ORDERED_SHEET.Cells(O_start_row, 1), ORDERED_SHEET.Cells(O_start_row + i - 1, obj_max_count)) = WorksheetFunction.Transpose(data)
    
    Erase data


    ORDERED_SHEET.Range(ORDERED_SHEET.Cells(O_start_row, 1), ORDERED_SHEET.Cells(O_end_row, O_end_col)).Borders.LineStyle = True
    
    
    
    Sleep 10
    
    
    'Dim P_end_row As Integer: P_end_row = data(Rows.count, 2).End(xlUp).row
    'Dim P_end_col As Integer: P_end_col = data(1, 100).End(xlToLeft).Column
    

    '*************************************************************************
'抽出

    
    Application.DisplayAlerts = False
    
    Call fso.DeleteFolder(folderPath, True)
    Set fso = Nothing

    
    
    If S_end_row > 6 Then
        Call test003(x)
    End If
    
test:


'========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "終了まで" & processTime
'========================================================
    



End Function


Function conversion_string(HOLIDAY_SHEET As Worksheet, data, index As Long, ByVal str As String, i As Long, VT As Object, DC As Object, MM As Object, RM As Object, CM As Object, sub_name As String, Remarks, Construction)
    str = Replace(str, """", "")
    With WorksheetFunction
        If index = 0 Then
            Dim customerCode As String: customerCode = Left(str, 5)
            If RM.Exists(customerCode) = True And Not RM.Item(customerCode) = "" Then
                data(0, i) = RM.Item(customerCode)
            End If
            data(2, i) = CM.Item(customerCode)
            data(3, i) = customerCode
        ElseIf index = 3 Then
            Dim propertyName As String: propertyName = .Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(.Substitute(str, "\", "_"), "/", "_"), ":", "_"), "*", "_"), "?", "_"), "<", "_"), ">", "_"), "|", "_")
            data(8, i) = propertyName
        ElseIf index = 4 Then
            Dim slipNum As String: slipNum = Right(.Trim(str), 7)
            data(6, i) = slipNum
        ElseIf index = 5 Then
            Dim floorNum As String: floorNum = .Trim(str)
            data(10, i) = floorNum
        ElseIf index = 6 Then
            Dim consecutiveNum As String: consecutiveNum = .Trim(str)
            data(7, i) = consecutiveNum
        ElseIf index = 7 Then
            data(11, i) = str
        ElseIf index = 8 Then
            data(12, i) = str
        ElseIf index = 9 Then
            data(1, i) = _
            .Text( _
                .WorkDay( _
                    DateValue( _
                        "20" & Left(str, 2) & "/" & Mid(str, 3, 2) & "/" & Right(str, 2) _
                    ), _
                -4, HOLIDAY_SHEET.Range(HOLIDAY_SHEET.Cells(3, 2), HOLIDAY_SHEET.Cells(186, 2))), _
            "yyyy/mm/dd")
            data(16, i) = _
            .Text( _
                DateValue( _
                    "20" & Left(str, 2) & "/" & Mid(str, 3, 2) & "/" & Right(str, 2) _
                ), _
            "yyyy/mm/dd")
        ElseIf index = 10 Then
            data(17, i) = _
            .Text( _
                DateValue( _
                    "20" & Left(str, 2) & "/" & Mid(str, 3, 2) & "/" & Right(str, 2) _
                ), _
            "yyyy/mm/dd")
        ElseIf index = 11 Then
            If str <> "" Then
                data(15, i) = .Replace(str, Len(str) - 1, 0, ":")
            End If
        ElseIf index = 12 Then
            If str <> "" Then
                If VT.Exists(str) = True Then
                    data(14, i) = VT.Item(str)
                End If
            End If
        ElseIf index = 13 Then
            If str <> "" Then: data(13, i) = "LTS": Else: data(13, i) = "店"
        ElseIf index = 14 Then
            data(19, i) = str
        ElseIf index = 15 Then
            data(20, i) = str
        ElseIf index = 16 Then
            If str Like "*ｼﾞﾄﾞｳｼﾖﾘ*" Then: data(18, i) = "": Else: data(18, i) = "非素通り"
        ElseIf index = 18 Then
            data(24, i) = str
        ElseIf index = 19 Then
            data(23, i) = str
        ElseIf index = 20 Then
            data(5, i) = data(3, i) & " " & data(7, i) & " " & data(8, i)
            data(9, i) = data(3, i) & " " & data(8, i)
            Dim con_key As String: con_key = data(6, i) & data(7, i) & data(10, i)
            If DC.Exists(con_key) = True Then
                If sub_name = "担当店抽出" Then
                    If DC.Item(con_key) = "確定" Or DC.Item(con_key) = "B10" Then
                        data(4, i) = DC.Item(con_key)
                    Else
                        data(4, i) = ""
                    End If
                ElseIf sub_name = "確定作業確認" Then
                    data(4, i) = DC.Item(con_key)
                End If
                data(21, i) = MM.Item(con_key)
                data(22, i) = Remarks.Item(con_key)
                If data(13, i) <> Construction.Item(con_key) Then
                    data(13, i) = Construction.Item(con_key) & "(変更済み)"
                End If
            End If
        End If
    End With
End Function



Sub DCS自動送付()

    'Dim driver As New ChromeDriver
    Dim driver As New Selenium.PhantomJSDriver
    Dim myBy As New By, elm As WebElement
    Dim slipNum As String, directionsURL As String, mapCheck1 As String, mapCheck2 As String
    Dim file As String, propertyName As String, customerCode As String, rc As VbMsgBoxResult
    Dim data1 As String, data2 As String, data3 As String, data4 As String, memo As String
    Dim filePath As String, fileName As String

    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")

    rc = MsgBox("本日確定分、自動送付しますか？", vbYesNo)
    If rc = vbYes Then
        Dim O_day_col As Integer: O_day_col = 2
        Dim O_custormer_col As Integer: O_custormer_col = 4
        Dim O_status_col As Integer: O_status_col = 5
        Dim O_slip_col As Integer: O_slip_col = 7
        Dim O_consecutive_col As Integer: O_consecutive_col = 8
        Dim O_file_col As Integer: O_file_col = 9
        Dim O_memo_col As Integer: O_memo_col = 22
        Dim O_start_row As Integer: O_start_row = 3
        Dim row As Long: row = O_start_row
        
        Dim F_start_row As Integer: F_start_row = 3
        Dim F_end_row As Integer: F_end_row = FRONT_SHEET.Cells(60000, 2).End(xlUp).row
        Dim F_path_col As Integer: F_path_col = 5
        Dim F_search_col As Integer: F_search_col = 2
        Dim F_mandatory_col As Integer: F_mandatory_col = 7
        Dim F_irregular_col As Integer: F_irregular_col = 6
        
        Dim userId As String: userId = FRONT_SHEET.Cells(2, 11).Value
        Dim password As String: password = FRONT_SHEET.Cells(3, 11).Value

        Dim errorSyntax As String: errorSyntax = ""

        If Not userId <> "" And password <> "" Then
            errorSyntax = "DCSログイン時のID・パスワードが記入されていません！"
            GoTo label1
        End If

        
        'Data1 = myform.d30001.Value
        'Data2 = myform.d20021.Value
        'Data3 = myform.d30077.Value
        'Data4 = myform.status.Value
        'url = "detail_genba.php?d30001=" + data1 + "&d20021=" + data2 + "&d30077=" + data3 + "&status=" + data4;
        
        With driver

            '.Start "chrome"
            .Start
            
            .Get "http://delivery.i2.inax.co.jp/index.php"
            
            .FindElementByName("userid").SendKeys userId
            
            .FindElementByName("password").SendKeys password
            
            .FindElementByXPath("//*[@value=""LOGIN""]").Click
            
            .Get "http://delivery.i2.inax.co.jp/check/list_search_charter_nzi.php"

            If Not .IsElementPresent(myBy.XPath("//*[@id=""main""]/form/table[1]/tbody/tr[1]/td[2]/select")) Then
                errorSyntax = "DCSログインできません！ID・パスワードの確認お願いします。"
                GoTo label1
            End If

            Do While ORDERED_SHEET.Cells(row, O_day_col) <> ""
                If ORDERED_SHEET.Cells(row, O_day_col) = Date Then
                    If ORDERED_SHEET.Cells(row, O_status_col) = "確定" Or ORDERED_SHEET.Cells(row, O_status_col) = "B10" Then
                        slipNum = ORDERED_SHEET.Cells(row, O_slip_col).Value
                        .FindElementByName("syukkabi").AsSelect.SelectByText ("すべて")
                        .FindElementByName("denno").Clear
                        .FindElementByName("denno").SendKeys slipNum
                        
                        .FindElementByName("Submit").Click
                        
                        Sleep 200
                        
                        If .IsElementPresent(myBy.XPath("//*[@title=""黄色地は締切済。赤色地は承認後に修正あり物件""]")) Then
                            directionsURL = "http://delivery.i2.inax.co.jp/check/detail_genba.php?d30001="
                            data1 = .FindElementByName("d30001").Value
                            data2 = .FindElementByName("d20021").Value
                            data3 = .FindElementByName("d30077").Value
                            data4 = .FindElementByXPath("//*[@title=""黄色地は締切済。赤色地は承認後に修正あり物件""]").FindElementByName("status").Value
                            directionsURL = directionsURL + data1 + "&d20021=" + data2 + "&d30077=" + data3 + "&status=" + data4
                            
                            .Get directionsURL
                            
                            Sleep 20

                            customerCode = ORDERED_SHEET.Cells(row, O_custormer_col).Value
                            filePath = WorksheetFunction.index(FRONT_SHEET.Range(FRONT_SHEET.Cells(F_start_row, F_path_col), FRONT_SHEET.Cells(F_end_row, F_path_col)), WorksheetFunction.Match(customerCode, FRONT_SHEET.Range(FRONT_SHEET.Cells(F_start_row, F_search_col), FRONT_SHEET.Cells(F_end_row, F_search_col)), 0))
                            Dim is_irregular As Boolean
                            If WorksheetFunction.index(FRONT_SHEET.Range(FRONT_SHEET.Cells(F_start_row, F_irregular_col), FRONT_SHEET.Cells(F_end_row, F_irregular_col)), WorksheetFunction.Match(customerCode, FRONT_SHEET.Range(FRONT_SHEET.Cells(F_start_row, F_search_col), FRONT_SHEET.Cells(F_end_row, F_search_col)), 0)) = "★" Then
                                fileName = ORDERED_SHEET.Cells(row, O_custormer_col).Value & " " & ORDERED_SHEET.Cells(row, O_file_col).Value
                                is_irregular = True
                            Else
                                fileName = ORDERED_SHEET.Cells(row, O_custormer_col).Value & " " & Left(ORDERED_SHEET.Cells(row, O_consecutive_col).Value, 6)
                                is_irregular = False
                            End If
                            file = ""
                            file = AttachMap(filePath, fileName, is_irregular)
                            If file = "" Then
                                file = SearchSubfolder(filePath, fileName, is_irregular)
                            End If
                            If file <> "" Then
                                .FindElementByName("fname_1").SendKeys file
                            End If
                            Sleep 20

                            memo = ORDERED_SHEET.Cells(row, O_memo_col).Value
                            If memo <> "" Then
                                .FindElementByName("memoin").Clear
                                .FindElementByName("memoin").SendKeys memo
                                Sleep 20
                            End If
                            
                            If .IsElementPresent(myBy.XPath("//*[@value=""OPEN""]")) Then
                                .FindElementByXPath("//*[@value=""OPEN""]").Click
                            End If
                            
                            .FindElementByXPath("//*[@value=""更新""]").Click
                            
                            Sleep 200

                            .Get "http://delivery.i2.inax.co.jp/check/list_search_charter_nzi.php"
                            
                        End If
                    End If
                End If
                
                row = row + 1
            Loop
        
        End With
    
        MsgBox "OK"
    End If

Exit Sub

label1:
MsgBox errorSyntax, Buttons:=vbCritical


End Sub

Function AttachMap(filePath As String, fileName As String, is_irregular As Boolean) As String
    Dim clone As String: clone = ""
    If is_irregular = True Then
        Dim deletWordCount As Integer: deletWordCount = 0
        Dim adjust_num As Integer: adjust_num = 6
        Dim wordLength As Integer: wordLength = Len(fileName) - adjust_num
        For deletWordCount = 0 To ((wordLength / 10) * 3)
            clone = Dir(filePath & "\" & "*" & Left(fileName, wordLength - deletWordCount + adjust_num) & "*" & ".pdf")
            If clone <> "" Then
                AttachMap = filePath & "\" & clone
                Exit For
            End If
        Next
    Else
        clone = Dir(filePath & "\" & "*" & fileName & "*" & ".pdf")
        If clone <> "" Then
            AttachMap = filePath & "\" & clone
            'Exit For
        End If
        ' Next
    End If
End Function

Function SearchSubfolder(filePath As String, fileName As String, is_irregular As Boolean) As String
    Dim f As Object
    Dim clone As String: clone = ""
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(filePath).SubFolders
            clone = AttachMap(f.path, fileName, is_irregular)
            If clone <> "" Then
                SearchSubfolder = clone
                Exit For
            End If
        Next f
    End With
End Function

Function test003(x As Workbook)
    Dim customerCode
    Dim propertyName As String, consecutiveNum As String
    
    Dim CUSTOMER_SHEET As Worksheet: Set CUSTOMER_SHEET = x.Sheets("店用出力")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = x.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = x.Sheets("注残現場一覧")
    
    '注残現場一覧の最上伝票位置
    Dim O_start_row As Integer: O_start_row = 3
    '注残現場一覧の最下伝票位置
    Dim O_end_row As Integer: O_end_row = ORDERED_SHEET.Cells(Rows.count, 7).End(xlUp).row

    '店用の最上伝票位置
    Dim C_start_row As Integer: C_start_row = 5
    '店用の最下伝票位置
    Dim C_end_row As Integer: C_end_row = CUSTOMER_SHEET.Cells(Rows.count, 1).End(xlUp).row
    If C_end_row < C_start_row Then
        C_end_row = C_start_row
    End If

    Dim C_end_col As Integer: C_end_col = 8

    '表紙の最上伝票位置
    Dim F_start_row As Integer: F_start_row = 5
    '表紙の最下伝票位置
    Dim F_end_row As Integer: F_end_row = FRONT_SHEET.Cells(Rows.count, 2).End(xlUp).row
    '特別対応の最上伝票位置
    Dim S_start_row As Integer: S_start_row = 7
    '特別対応の最下伝票位置
    Dim S_end_row As Integer: S_end_row = FRONT_SHEET.Cells(Rows.count, 10).End(xlUp).row

    CUSTOMER_SHEET.Range(CUSTOMER_SHEET.Cells(C_start_row, 1), CUSTOMER_SHEET.Cells(C_end_row, 8)) = ""
    
    Dim orders()
    
    Dim row As Integer
    Dim col As Integer
    Dim count As Integer
    
    count = 0
    With WorksheetFunction
        For row = O_start_row To O_end_row
            customerCode = ORDERED_SHEET.Cells(row, 4)
            If 0 < .CountIf(FRONT_SHEET.Range(FRONT_SHEET.Cells(S_start_row, 10), FRONT_SHEET.Cells(S_end_row, 10)), customerCode) Then
                ReDim Preserve orders(7, count)
                orders(0, count) = customerCode
                orders(1, count) = ORDERED_SHEET.Cells(row, 9)
                orders(2, count) = ORDERED_SHEET.Cells(row, 7)
                orders(3, count) = ORDERED_SHEET.Cells(row, 12)
                orders(4, count) = ORDERED_SHEET.Cells(row, 11)
                orders(5, count) = ORDERED_SHEET.Cells(row, 18)
                If ORDERED_SHEET.Cells(row, 14) = "店" Then
                    If ORDERED_SHEET.Cells(row, 5) = "確定" Then
                        orders(6, count) = "確定済"
                    Else
                        orders(6, count) = ""
                    End If
                End If
                orders(7, count) = ORDERED_SHEET.Cells(row, 2)
                count = count + 1
            End If
        Next
    CUSTOMER_SHEET.Range(CUSTOMER_SHEET.Cells(C_start_row, 1), CUSTOMER_SHEET.Cells(C_start_row + count - 1, C_end_col)) = .Transpose(orders)
    End With

End Function



Function toggle_library(macro_name)
    Dim refObj As Variant
    Dim WshNetworkObject As Object
    Dim user_id  As String
    Dim selenium_path As String
    Dim bResult As Boolean
    
    With ThisWorkbook.VBProject
        For Each refObj In ThisWorkbook.VBProject.References
            If refObj.Description = "Selenium Type Library" Then
                .References.Remove refObj
            End If
        Next refObj
    End With
    
    Set WshNetworkObject = CreateObject("WScript.Network")

    user_id = WshNetworkObject.UserName
    selenium_path = "C:\Users\" & user_id & "\AppData\Local\SeleniumBasic\Selenium32.tlb"
    
    bResult = IsExistDirB(selenium_path)
    
    If bResult = False Then: GoTo exit_sub

    ActiveWorkbook.VBProject.References.AddFromFile selenium_path
    
    
    Application.Run macro_name
    
    
    With ThisWorkbook.VBProject
        For Each refObj In ThisWorkbook.VBProject.References
            If refObj.Description = "Selenium Type Library" Then
                .References.Remove refObj
            End If
        Next refObj
    End With
    
    Set WshNetworkObject = Nothing
    
    Exit Function
    
exit_sub:
    
    Set WshNetworkObject = Nothing
    
    MsgBox "DCS自動送信の設定が完了していません。", vbCritical

End Function

Sub Button_自動送信()

Call toggle_library("DCS自動送付")

End Sub

Sub Button_送信チェック()

Call toggle_library("送信チェック")

End Sub

Function IsExistDirB(a_sFolder As String) As Boolean
    Dim result
    
    result = Dir(a_sFolder, vbDirectory)
    
    If result = "" Then
        '// フォルダが存在しない
        IsExistDirB = False
    Else
        '// フォルダが存在する
        IsExistDirB = True
    End If
End Function


Sub IsExistDirA利用例()
    Dim A As String
    Dim bResult As Boolean
    Dim WshNetworkObject As Object
    Dim user_id  As String
    
    bResult = False
    
    Set WshNetworkObject = CreateObject("WScript.Network")
    user_id = WshNetworkObject.UserName
    
    A = "C:\Users\" & user_id & "\AppData\Local\SeleniumBasic"
    
    bResult = IsExistDirB(A)
    
    MsgBox bResult
    
    Set WshNetworkObject = Nothing
End Sub

Function creaeStringAlphaNum(ByVal lLength As Long) As String

    Dim iBeginCodeN     As Integer
    Dim iEndCodeN       As Integer
    Dim iBeginCodeAU    As Integer
    Dim iEndCodeAU      As Integer
    Dim iBeginCodeAL    As Integer
    Dim iEndCodeAL      As Integer
    Dim iCode   As Integer
    Dim sResult As String
    Dim i       As Long

    iBeginCodeN = Asc("0")
    iEndCodeN = Asc("9")
    iBeginCodeAU = Asc("A")
    iEndCodeAU = Asc("Z")
    iBeginCodeAL = Asc("a")
    iEndCodeAL = Asc("z")

    Randomize

    For i = 1 To lLength
        Do While True
            iCode = Int((iEndCodeAL - iBeginCodeN + 1) * Rnd) + iBeginCodeN

            Select Case iCode
            Case iBeginCodeN To iEndCodeN
                Exit Do
            Case iBeginCodeAU To iEndCodeAU
                Exit Do
            Case iBeginCodeAL To iEndCodeAL
                Exit Do
            End Select
        Loop

        sResult = sResult & Chr(iCode)
    Next i

    creaeStringAlphaNum = sResult

End Function


Sub DCS改良()
    Dim driver As New ChromeDriver
    'Dim driver As New Selenium.PhantomJSDriver
    Dim myBy As New By, elm As WebElement
    Dim slipNum As String, directionsURL As String, mapCheck1 As String, mapCheck2 As String
    Dim file As String, propertyName As String, customerCode As String, rc As VbMsgBoxResult
    Dim data1 As String, data2 As String, data3 As String, data4 As String, memo As String
    Dim filePath As String, fileName As String, today As String

    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")

    rc = MsgBox("本日確定分、自動送付しますか？", vbYesNo)
    If rc = vbYes Then
        Dim O_day_col As Integer: O_day_col = 2
        Dim O_custormer_col As Integer: O_custormer_col = 4
        Dim O_status_col As Integer: O_status_col = 5
        Dim O_slip_col As Integer: O_slip_col = 7
        Dim O_consecutive_col As Integer: O_consecutive_col = 8
        Dim O_file_col As Integer: O_file_col = 9
        Dim O_memo_col As Integer: O_memo_col = 22
        Dim O_start_row As Integer: O_start_row = 3
        Dim row As Long: row = O_start_row
        
        Dim F_start_row As Integer: F_start_row = 3
        Dim F_end_row As Integer: F_end_row = FRONT_SHEET.Cells(60000, 2).End(xlUp).row
        Dim F_path_col As Integer: F_path_col = 5
        Dim F_search_col As Integer: F_search_col = 2
        Dim F_mandatory_col As Integer: F_mandatory_col = 7
        
        Dim userId As String: userId = FRONT_SHEET.Cells(2, 11).Value
        Dim password As String: password = FRONT_SHEET.Cells(3, 11).Value
        
        'Data1 = myform.d30001.Value
        'Data2 = myform.d20021.Value
        'Data3 = myform.d30077.Value
        'Data4 = myform.status.Value
        'url = "detail_genba.php?d30001=" + data1 + "&d20021=" + data2 + "&d30077=" + data3 + "&status=" + data4;
        
        With driver

            .Start "chrome"
            '.Start
            
            .Get "http://delivery.i2.inax.co.jp/index.php"
            
            .FindElementByName("userid").SendKeys userId
            
            .FindElementByName("password").SendKeys password
            
            .FindElementByXPath("//*[@value=""LOGIN""]").Click
            
            .Get "http://delivery.i2.inax.co.jp/check/den_search_charter_check.php"
            
            row = 88

            'Do While ORDERED_SHEET.Cells(row, O_day_col) <> ""
                If ORDERED_SHEET.Cells(row, O_day_col) = Date Then
                    If ORDERED_SHEET.Cells(row, O_status_col) = "確定" Or ORDERED_SHEET.Cells(row, O_status_col) = "B10" Then
                        slipNum = ORDERED_SHEET.Cells(row, O_slip_col).Value
                        .FindElementByName("denno").Clear
                        .FindElementByName("denno").SendKeys slipNum
                        
                        .FindElementByName("Submit").Click
                        
                        Sleep 200
                        
                        If .IsElementPresent(myBy.XPath("//*[@title=""黄色地は締切済。赤色地は承認後に修正あり物件""]")) Then
                            directionsURL = "http://delivery.i2.inax.co.jp/check/detail_genba.php?d30001="
                            data1 = .FindElementByName("d30001").Value
                            data2 = .FindElementByName("d20021").Value
                            data3 = .FindElementByName("d30077").Value
                            data4 = .FindElementByXPath("//*[@title=""黄色地は締切済。赤色地は承認後に修正あり物件""]").FindElementByName("status").Value
                            directionsURL = directionsURL + data1 + "&d20021=" + data2 + "&d30077=" + data3 + "&status=" + data4
                            
                            .Get directionsURL
                            
                        End If
                        
                        
                    End If
                End If
                
                'row = row + 1
            'Loop
        
        End With
    
        MsgBox "OK"
    End If
  

End Sub


Sub Button_DCS改良()

Call toggle_library("DCS改良")

End Sub



Sub testOpen1()
    Dim customerCode
    Dim InboxFolder, wsh As Object, fso As Object, path1 As String
    Dim myNameSpace, objmailItem As Object, propertyName As String, consecutiveNum As String, test1 As Integer, con_key As String
    Dim outlookObj As Outlook.Application
    Dim file As String, folderPath As String, fileName As String, name As String, slipNum As String, floorNum As String, Folder As String
    Dim x As Workbook: Set x = ActiveWorkbook
    Dim data()
    
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
    
    'Const email_subject As String = "*【見張る君】「西日本受注C 注残現場(20時時点)*"
    Const email_subject As String = "*【見張る君】「注残確認用*"
    
    'エラーを無視することで添付のない転送メールで止まることを回避
    'On Error Resume Next
    
    Dim HOLIDAY_SHEET As Worksheet: Set HOLIDAY_SHEET = ActiveWorkbook.Sheets("2021 3月期 休日")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim CUSTOMER_SHEET As Worksheet: Set CUSTOMER_SHEET = ActiveWorkbook.Sheets("店用出力")
    
    '状態を格納するオブジェクト
    Dim DC As Object: Set DC = CreateObject("Scripting.Dictionary")
    Dim VT As Object: Set VT = CreateObject("Scripting.Dictionary")
    Dim MM As Object: Set MM = CreateObject("Scripting.Dictionary")

    VT.Add "A", "トレーラ": VT.Add "B", "11t": VT.Add "C", "4t": VT.Add "D", "指定なし": VT.Add "E", "2t": VT.Add "F", "4tｼｮｰﾄ": VT.Add "G", "4tﾕﾆｯｸ": VT.Add "H", "2tﾕﾆｯｸ": VT.Add "I", "2tｼｮｰﾄ": VT.Add "J", "11tﾕﾆｯｸ": VT.Add "K", "2t単独車": VT.Add "L", "2tｼｮｰﾄ単独車": VT.Add "M", "軽トラ": VT.Add "N", "11t平ﾎﾞﾃﾞｨ": VT.Add "P", "4t平ﾎﾞﾃﾞｨ": VT.Add "Z", "その他":
    
    'ソートのクリア
    Dim FW As Boolean: FW = False
    If ORDERED_SHEET.FilterMode Then ORDERED_SHEET.ShowAllData
    Do While FW = False
        If Not ORDERED_SHEET.FilterMode Then
            FW = True
        End If
    Loop
    
    '注残現場一覧の最上伝票位置
    Dim O_start_row As Integer: O_start_row = 3
    '注残現場一覧の最下伝票位置
    Dim O_end_row As Integer: O_end_row = ORDERED_SHEET.Cells(Rows.count, 7).End(xlUp).row
    '注残現場一覧の最右伝票位置
    Dim O_end_col As Integer: O_end_col = 23
    
    '特別対応の最上伝票位置
    Dim S_start_row As Integer: S_start_row = 7
    '特別対応の最下伝票位置
    Dim S_end_row As Integer: S_end_row = FRONT_SHEET.Cells(Rows.count, 10).End(xlUp).row
    

    
    '確定状態を連番と共に格納(連番:伝№＆管理№＆フロア)
    Dim i As Integer
    Dim key As String
    For i = O_start_row To O_end_row
        If ORDERED_SHEET.Cells(i, 7).Value <> 0 Then
            key = ORDERED_SHEET.Cells(i, 7).Value & ORDERED_SHEET.Cells(i, 8).Value & ORDERED_SHEET.Cells(i, 11).Value
            DC.Add key, ORDERED_SHEET.Cells(i, 5).Value
            MM.Add key, ORDERED_SHEET.Cells(i, 22).Value
        End If
    Next
    
    Sleep 20
        
'*****************************************************************
'デスクトップのアドレスをを取得
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    path1 = wsh.specialFolders("MyDocuments")
    Folder = creaeStringAlphaNum(20)
    fileName = Folder & ".csv"
    file = path1 & "\" & Folder & "\" & fileName
    Set outlookObj = CreateObject("Outlook.Application")
    Set myNameSpace = outlookObj.GetNamespace("MAPI")
    Set InboxFolder = myNameSpace.GetDefaultFolder(6)
'******************************************************************
'メールを検索してデスクトップへ保存
    Sleep 20
    For i = 1 To 100
        Set objmailItem = InboxFolder.Items(i)
        If objmailItem.Subject Like email_subject Then
            folderPath = path1 & "\" & Folder
            If Dir(folderPath, vbDirectory) = "" Then
                MkDir folderPath
            Else
                If Not Dir(file) = "" Then
                    Exit For
                End If
            End If
            
            objmailItem.Attachments.Item(1).SaveAsFile file
          
        End If
    Next
    
    Sleep 10
    
    If Dir(file) = "" Then
        MsgBox "受信BOXに【見張る君】が見つかりませんでした。", Buttons:=vbCritical
        Exit Sub
    End If
'*******************************************************************

'ファイルを開く

    Dim count As Long, j As Long
    Dim st As String

    i = 0
    
    Dim buf As String, flag As Boolean, A As Variant, B As Variant
    Open file For Input As #1
        Do Until EOF(1)
            ReDim Preserve data(28, i)
            Line Input #1, buf
            B = Split(buf, ",")
            If UBound(B) > 21 Then
                j = 0
                count = 0
                Do
                    If Right(B(j), 1) <> """" Then
                        flag = False
                        Do While flag = False
                            st = B(j - count) & "," & B(j + 1)
                            count = count + 1
                            j = j + 1
                            If Right(st, 1) = """" Then
                                flag = True
                            End If
                        Loop
                        A(j - count) = st
                    Else
                        A(j - count) = B(j)
                    End If
        
                    j = j + 1
                Loop While j <= UBound(B)
            Else
                A = B
            End If
            
            If UBound(A) > 21 Then
                Debug.Print i
            End If
            For j = 0 To UBound(A)
                data(j, i) = Replace(A(j), """", "")
                'data(j, i) = A(j)
            Next j
            i = i + 1
        Loop
    Close #1



'****************************************************************
    
    Application.DisplayAlerts = False
    
    Call fso.DeleteFolder(folderPath, True)

    Set fso = Nothing
    Set outlookObj = Nothing
    Set wsh = Nothing


'終了時間取得
endTime = Timer

'処理時間計算
processTime = endTime - startTime

Debug.Print processTime

For i = 0 To 28
    Debug.Print data(i, 3413)
    'Debug.Print data(1, 0)
Next


End Sub


Sub testOpen2()
    Dim customerCode
    Dim InboxFolder, wsh As Object, fso As Object, path1 As String
    Dim myNameSpace, objmailItem As Object, propertyName As String, consecutiveNum As String, test1 As Integer, con_key As String
    Dim outlookObj As Outlook.Application
    Dim file As String, folderPath As String, fileName As String, name As String, slipNum As String, floorNum As String, Folder As String
    Dim x As Workbook: Set x = ActiveWorkbook
    Dim data()
    
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
    
    Const email_subject As String = "*【見張る君】「西日本受注C 注残現場(20時時点)*"
    'Const email_subject As String = "*【見張る君】「注残確認用*"
    
    'エラーを無視することで添付のない転送メールで止まることを回避
    'On Error Resume Next
    
    Dim HOLIDAY_SHEET As Worksheet: Set HOLIDAY_SHEET = ActiveWorkbook.Sheets("2021 3月期 休日")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim CUSTOMER_SHEET As Worksheet: Set CUSTOMER_SHEET = ActiveWorkbook.Sheets("店用出力")
    
    '状態を格納するオブジェクト
    Dim DC As Object: Set DC = CreateObject("Scripting.Dictionary")
    Dim VT As Object: Set VT = CreateObject("Scripting.Dictionary")
    Dim MM As Object: Set MM = CreateObject("Scripting.Dictionary")

    VT.Add "A", "トレーラ": VT.Add "B", "11t": VT.Add "C", "4t": VT.Add "D", "指定なし": VT.Add "E", "2t": VT.Add "F", "4tｼｮｰﾄ": VT.Add "G", "4tﾕﾆｯｸ": VT.Add "H", "2tﾕﾆｯｸ": VT.Add "I", "2tｼｮｰﾄ": VT.Add "J", "11tﾕﾆｯｸ": VT.Add "K", "2t単独車": VT.Add "L", "2tｼｮｰﾄ単独車": VT.Add "M", "軽トラ": VT.Add "N", "11t平ﾎﾞﾃﾞｨ": VT.Add "P", "4t平ﾎﾞﾃﾞｨ": VT.Add "Z", "その他":
    
    'ソートのクリア
    Dim FW As Boolean: FW = False
    If ORDERED_SHEET.FilterMode Then ORDERED_SHEET.ShowAllData
    Do While FW = False
        If Not ORDERED_SHEET.FilterMode Then
            FW = True
        End If
    Loop
    
    '注残現場一覧の最上伝票位置
    Dim O_start_row As Integer: O_start_row = 3
    '注残現場一覧の最下伝票位置
    Dim O_end_row As Integer: O_end_row = ORDERED_SHEET.Cells(Rows.count, 7).End(xlUp).row
    '注残現場一覧の最右伝票位置
    Dim O_end_col As Integer: O_end_col = 23
    
    '特別対応の最上伝票位置
    Dim S_start_row As Integer: S_start_row = 7
    '特別対応の最下伝票位置
    Dim S_end_row As Integer: S_end_row = FRONT_SHEET.Cells(Rows.count, 10).End(xlUp).row
    

    
    '確定状態を連番と共に格納(連番:伝№＆管理№＆フロア)
    Dim i As Integer
    Dim key As String
    For i = O_start_row To O_end_row
        If ORDERED_SHEET.Cells(i, 7).Value <> 0 Then
            key = ORDERED_SHEET.Cells(i, 7).Value & ORDERED_SHEET.Cells(i, 8).Value & ORDERED_SHEET.Cells(i, 11).Value
            DC.Add key, ORDERED_SHEET.Cells(i, 5).Value
            MM.Add key, ORDERED_SHEET.Cells(i, 22).Value
        End If
    Next
    
    Sleep 20
        
'*****************************************************************
'デスクトップのアドレスをを取得
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    path1 = wsh.specialFolders("MyDocuments")
    Folder = creaeStringAlphaNum(20)
    fileName = Folder & ".csv"
    file = path1 & "\" & Folder & "\" & fileName
    Set outlookObj = CreateObject("Outlook.Application")
    Set myNameSpace = outlookObj.GetNamespace("MAPI")
    Set InboxFolder = myNameSpace.GetDefaultFolder(6)
'******************************************************************
'メールを検索してデスクトップへ保存
    Sleep 20
    For i = 1 To 100
        Set objmailItem = InboxFolder.Items(i)
        If objmailItem.Subject Like email_subject Then
            folderPath = path1 & "\" & Folder
            If Dir(folderPath, vbDirectory) = "" Then
                MkDir folderPath
            Else
                If Not Dir(file) = "" Then
                    Exit For
                End If
            End If
            
            objmailItem.Attachments.Item(1).SaveAsFile file
          
        End If
    Next
    
    Sleep 10
    
    If Dir(file) = "" Then
        MsgBox "受信BOXに【見張る君】が見つかりませんでした。", Buttons:=vbCritical
        Exit Sub
    End If
'*******************************************************************

'ファイルを開く


    
    With ActiveSheet.QueryTables.Add("TEXT;" & file, Range("A1"))
        .TextFileCommaDelimiter = True
        .Refresh
    End With

    data = Range("A1").CurrentRegion



'****************************************************************
    
    Application.DisplayAlerts = False
    
    Call fso.DeleteFolder(folderPath, True)

    Set fso = Nothing
    Set outlookObj = Nothing
    Set wsh = Nothing


'終了時間取得
endTime = Timer

'処理時間計算
processTime = endTime - startTime

Debug.Print processTime

For i = 0 To 28
    'Debug.Print data(i, 2114)
    'Debug.Print data(1, 0)
Next


End Sub


Sub testOpen3()
    Dim customerCode
    Dim InboxFolder, wsh As Object, fso As Object, path1 As String
    Dim myNameSpace, objmailItem As Object, propertyName As String, consecutiveNum As String, test1 As Integer, con_key As String
    Dim outlookObj As Outlook.Application
    Dim file As String, folderPath As String, fileName As String, name As String, slipNum As String, floorNum As String, Folder As String
    Dim x As Workbook: Set x = ActiveWorkbook
    Dim data()
    
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
    
    Const email_subject As String = "*【見張る君】「西日本受注C 注残現場(20時時点)*"
    'Const email_subject As String = "*【見張る君】「注残確認用*"
    
    'エラーを無視することで添付のない転送メールで止まることを回避
    'On Error Resume Next
    
    Dim HOLIDAY_SHEET As Worksheet: Set HOLIDAY_SHEET = ActiveWorkbook.Sheets("2021 3月期 休日")
    Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")
    Dim ORDERED_SHEET As Worksheet: Set ORDERED_SHEET = ActiveWorkbook.Sheets("注残現場一覧")
    Dim CUSTOMER_SHEET As Worksheet: Set CUSTOMER_SHEET = ActiveWorkbook.Sheets("店用出力")
    
    '状態を格納するオブジェクト
    Dim DC As Object: Set DC = CreateObject("Scripting.Dictionary")
    Dim VT As Object: Set VT = CreateObject("Scripting.Dictionary")
    Dim MM As Object: Set MM = CreateObject("Scripting.Dictionary")

    VT.Add "A", "トレーラ": VT.Add "B", "11t": VT.Add "C", "4t": VT.Add "D", "指定なし": VT.Add "E", "2t": VT.Add "F", "4tｼｮｰﾄ": VT.Add "G", "4tﾕﾆｯｸ": VT.Add "H", "2tﾕﾆｯｸ": VT.Add "I", "2tｼｮｰﾄ": VT.Add "J", "11tﾕﾆｯｸ": VT.Add "K", "2t単独車": VT.Add "L", "2tｼｮｰﾄ単独車": VT.Add "M", "軽トラ": VT.Add "N", "11t平ﾎﾞﾃﾞｨ": VT.Add "P", "4t平ﾎﾞﾃﾞｨ": VT.Add "Z", "その他":
    
    'ソートのクリア
    Dim FW As Boolean: FW = False
    If ORDERED_SHEET.FilterMode Then ORDERED_SHEET.ShowAllData
    Do While FW = False
        If Not ORDERED_SHEET.FilterMode Then
            FW = True
        End If
    Loop
    
    '注残現場一覧の最上伝票位置
    Dim O_start_row As Integer: O_start_row = 3
    '注残現場一覧の最下伝票位置
    Dim O_end_row As Integer: O_end_row = ORDERED_SHEET.Cells(Rows.count, 7).End(xlUp).row
    '注残現場一覧の最右伝票位置
    Dim O_end_col As Integer: O_end_col = 23
    
    '特別対応の最上伝票位置
    Dim S_start_row As Integer: S_start_row = 7
    '特別対応の最下伝票位置
    Dim S_end_row As Integer: S_end_row = FRONT_SHEET.Cells(Rows.count, 10).End(xlUp).row
    

    
    '確定状態を連番と共に格納(連番:伝№＆管理№＆フロア)
    Dim i As Integer
    Dim key As String
    For i = O_start_row To O_end_row
        If ORDERED_SHEET.Cells(i, 7).Value <> 0 Then
            key = ORDERED_SHEET.Cells(i, 7).Value & ORDERED_SHEET.Cells(i, 8).Value & ORDERED_SHEET.Cells(i, 11).Value
            DC.Add key, ORDERED_SHEET.Cells(i, 5).Value
            MM.Add key, ORDERED_SHEET.Cells(i, 22).Value
        End If
    Next
    
    Sleep 20
        
'*****************************************************************
'デスクトップのアドレスをを取得
    Set wsh = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    path1 = wsh.specialFolders("MyDocuments")
    Folder = creaeStringAlphaNum(20)
    fileName = Folder & ".csv"
    file = path1 & "\" & Folder & "\" & fileName
    Set outlookObj = CreateObject("Outlook.Application")
    Set myNameSpace = outlookObj.GetNamespace("MAPI")
    Set InboxFolder = myNameSpace.GetDefaultFolder(6)
'******************************************************************
'メールを検索してデスクトップへ保存
    Sleep 20
    For i = 1 To 100
        Set objmailItem = InboxFolder.Items(i)
        If objmailItem.Subject Like email_subject Then
            folderPath = path1 & "\" & Folder
            If Dir(folderPath, vbDirectory) = "" Then
                MkDir folderPath
            Else
                If Not Dir(file) = "" Then
                    Exit For
                End If
            End If
            
            objmailItem.Attachments.Item(1).SaveAsFile file
          
        End If
    Next
    
    Sleep 10
    
    If Dir(file) = "" Then
        MsgBox "受信BOXに【見張る君】が見つかりませんでした。", Buttons:=vbCritical
        Exit Sub
    End If
'*******************************************************************

'ファイルを開く


    
    With ActiveSheet.QueryTables.Add("TEXT;" & file, Range("A1"))
        .TextFileCommaDelimiter = True
        .Refresh
    End With

    data = Range("A1").CurrentRegion



'****************************************************************
    
    Application.DisplayAlerts = False
    
    Call fso.DeleteFolder(folderPath, True)

    Set fso = Nothing
    Set outlookObj = Nothing
    Set wsh = Nothing


'終了時間取得
endTime = Timer

'処理時間計算
processTime = endTime - startTime

Debug.Print processTime

For i = 0 To 28
    'Debug.Print data(i, 2114)
    'Debug.Print data(1, 0)
Next


End Sub

Sub test5()

'========================================================
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
'========================================================

Dim FRONT_SHEET As Worksheet: Set FRONT_SHEET = ActiveWorkbook.Sheets("表紙")

Dim i As Long
Dim startrow As Long: startrow = 3
Dim endrow As Long: endrow = FRONT_SHEET.Cells(Rows.count, 2).End(xlUp).row
Dim data()
Dim key As String
Dim test(5, 5)


Call testtest2(FRONT_SHEET, i, startrow, endrow, data, key, endrow - 1000)

For i = 0 To 5
    'Debug.Print data(0, i)
Next

test(5, 5) = data(0, 5)

Debug.Print test(5, 5)

'========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "終了まで" & processTime
'========================================================

End Sub

Function testtest1(FRONT_SHEET As Worksheet, i As Long, startrow As Long, endrow As Long, data, key As String)
Dim count As Long: count = 0
For i = startrow To endrow
    key = FRONT_SHEET.Cells(i, 2).Value
    If 0 < WorksheetFunction.CountIf(FRONT_SHEET.Range(FRONT_SHEET.Cells(startrow, 2), FRONT_SHEET.Cells(endrow, 2)), key) Then
        ReDim Preserve data(0, count)
        data(0, count) = "OK"
        count = count + 1
    End If
Next



End Function

Function testtest2(FRONT_SHEET As Worksheet, i As Long, startrow As Long, endrow As Long, data, key As String, test As Long)
Dim DC As Object: Set DC = CreateObject("Scripting.Dictionary")
Dim count As Long: count = 0
For i = startrow To endrow
    key = FRONT_SHEET.Cells(i, 2).Value
    DC.Add key, FRONT_SHEET.Cells(i, 3).Value
Next

For i = startrow To endrow
    key = FRONT_SHEET.Cells(i, 2).Value
    If DC.Exists(key) = True Or key = "店コード" Then
        ReDim Preserve data(0, count)
        data(0, count) = DC.Item(key)
        count = count + 1
    End If

Next

Debug.Print test


End Function

Sub test33()

'========================================================
    Dim startTime As Double
    Dim endTime As Double
    Dim processTime As Double
     
    '開始時間取得
    startTime = Timer
'========================================================
    Dim i As Long
    Dim data()
    
    For i = 0 To 200000
        ReDim Preserve data(0, i)
        data(0, i) = testfunc1(i)
    Next
    
    'Debug.Print i

'========================================================
    
    '終了時間取得
    endTime = Timer
    
    '処理時間計算
    processTime = endTime - startTime
    
    Debug.Print "終了まで" & processTime
'========================================================
    
    
End Sub

Function testfunc3(ByVal i As String)
    Dim str As String: str = Right(i, 1)
    If str = "0" Then
        testfunc1 = "0"
    ElseIf str = "1" Then
        testfunc1 = "1"
    ElseIf str = "2" Then
        testfunc1 = "2"
    ElseIf str = "3" Then
        testfunc1 = "3"
    ElseIf str = "4" Then
        testfunc1 = "4"
    ElseIf str = "5" Then
        testfunc1 = "5"
    ElseIf str = "6" Then
        testfunc1 = "6"
    ElseIf str = "7" Then
        testfunc1 = "7"
    ElseIf str = "8" Then
        testfunc1 = "8"
    ElseIf str = "9" Then
        testfunc1 = "9"
    End If
End Function

Function testfunc2(ByVal i As String)
    Dim str As String: str = Right(i, 1)
    Select Case str
        Case "0"
            testfunc2 = "0"
        Case "1"
            testfunc2 = "1"
        Case "2"
            testfunc2 = "2"
        Case "3"
            testfunc2 = "3"
        Case "4"
            testfunc2 = "4"
        Case "5"
            testfunc2 = "5"
        Case "6"
            testfunc2 = "6"
        Case "7"
            testfunc2 = "7"
        Case "8"
            testfunc2 = "8"
        Case "9"
            testfunc2 = "9"
    End Select
End Function

Sub ts()
Dim driver As Object: Set driver = CreateObject("Selenium.ChromeDriver")

With driver
    .Start "chrome"
    '.Start
    .Get "http://groups.intra.lixil.lan/sites/JP/LJC/Pages/Index.aspx"
    MsgBox "OK"
End With

End Sub




