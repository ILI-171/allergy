Sub search_menu()
'既製品の検索
    Call search_menu_f("業者１")
    Call search_menu_f("業者２")
    Call search_menu_f("業者３")
    Call search_menu_f("業者４")
    Call search_menu_f("業者５")
    Call search_menu_f("業者６")
    Call search_menu_f("業者７")
    Call search_menu_f("業者８")

End Sub
Sub checkbox()
'チェックボックスの内容を文字に直す
    Dim tw As Worksheet
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    Count = 0
    ale = ""
    
    For i = 4 To 11
        If tw.Cells(22, i).Value = True Then
            ale = ale & " " & tw.Cells(21, i).Text
        Else
            Count = Count + 1
        End If
    Next i
    
    For i = 4 To 13
        If tw.Cells(25, i).Value = True Then
            ale = ale & " " & tw.Cells(24, i).Text
        Else
            Count = Count + 1
        End If
    Next i
    
    For i = 4 To 13
        If tw.Cells(28, i).Value = True Then
            ale = ale & " " & tw.Cells(27, i).Text
        Else
            Count = Count + 1
        End If
    Next i
    If Count = 28 Then
        ale = "/"
    End If
    
    tw.Cells(18, 4) = ale

End Sub
Sub all()
'チェックボックスのチェックをすべて外す
    Application.ScreenUpdating = False

    Dim tw As Worksheet
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    
    ale = ""
    
    For i = 4 To 11
        tw.Cells(22, i) = False
    Next i
    
    For i = 4 To 13
        tw.Cells(25, i) = False
    Next i
    
    For i = 4 To 13
        tw.Cells(28, i) = False
    Next i
    
    Call checkbox
    
    tw.Cells(18, 4) = "/"

    Application.ScreenUpdating = True
    
End Sub
Sub update()
'既製品アレルギー内容の更新

    Dim tw As Worksheet
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    tw.Cells(10, 14) = tw.Cells(18, 4).Text
    
    Call update_f("業者１")
    Call update_f("業者２")
    Call update_f("業者３")
    Call update_f("業者４")
    Call update_f("業者５")
    Call update_f("業者６")
    Call update_f("業者７")
    Call update_f("業者８")
    
    For i = 34 To 68 Step 2
        If tw.Cells(i, "D").Value = "" Then
            tw.Cells(i, "D") = tw.Cells(10, 7).Value
            Exit For
        End If
        If i = 68 Then
            For j = 34 To 68 Step 2
                If tw.Cells(j, "J").Value = "" Then
                    tw.Cells(j, "J") = tw.Cells(10, 7).Value
                    Exit For
                End If
            Next j
        End If
    Next i
    
    MsgBox ("更新が完了しました")

End Sub
Sub update_all()
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    UserForm1.Show vbModeless
    Dim tw As Worksheet
    Dim ws As Worksheet
    Dim mb As Workbook
    Dim num_rng As Range
    Dim ipoter_rng As Range
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    Set ws_new = ThisWorkbook.Sheets("レシピ  (新)")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set xlApp = CreateObject("Excel.Application")
    Dim path()
    Dim j()
    Floor = tw.Cells(4, 14).Value
    ReDim path(1 To 100, 1 To 10)
    ReDim j(1 To 10)
    Dim alergy(1 To 28)
    Dim ale_c(1 To 28)
    Dim sheet(1 To 9)
    alergy(1) = "えび"
    alergy(2) = "かに"
    alergy(3) = "くるみ"
    alergy(4) = "小麦"
    alergy(5) = "そば"
    alergy(6) = "卵"
    alergy(7) = "乳"
    alergy(8) = "落花生"
    alergy(9) = "アーモンド"
    alergy(10) = "あわび"
    alergy(11) = "いか"
    alergy(12) = "いくら"
    alergy(13) = "オレンジ"
    alergy(14) = "カシューナッツ"
    alergy(15) = "キウイフルーツ"
    alergy(16) = "牛肉"
    alergy(17) = "ごま"
    alergy(18) = "さけ"
    alergy(19) = "さば"
    alergy(20) = "大豆"
    alergy(21) = "鶏肉"
    alergy(22) = "バナナ"
    alergy(23) = "豚肉"
    alergy(24) = "まつたけ"
    alergy(25) = "もも"
    alergy(26) = "やまいも"
    alergy(27) = "りんご"
    alergy(28) = "ゼラチン"
    sheet(1) = "業者１"
    sheet(2) = "業者２"
    sheet(3) = "業者３"
    sheet(4) = "業者４"
    sheet(5) = "業者５"
    sheet(6) = "業者６"
    sheet(7) = "業者７"
    sheet(8) = "業者８"
    sheet(9) = "パートレシピ"
    x = 11
    folder_path = tw.Cells(4, 8).Text
    
    '指定フォルダ内のフォルダパスをすべて取得
    For i = 1 To 10
        If i = 1 Then
            Set objFolder = objFSO.GetFolder(folder_path)
            j(i) = 1
            
            For Each objSubFolder In objFolder.SubFolders
                    path(j(i), i) = objSubFolder.path
                    j(i) = j(i) + 1
            Next objSubFolder
            j(i) = j(i) - 1
        Else
            j(i) = 1
            For k = 1 To j(i - 1)
                Set objFolder = objFSO.GetFolder(path(k, i - 1))
                For Each objSubFolder In objFolder.SubFolders
                    path(j(i), i) = objSubFolder.path
                    j(i) = j(i) + 1
                Next objSubFolder
            Next k
            j(i) = j(i) - 1
        End If
    Next i
    
    'メニューエクセルを一つずつ検索し、各々更新していく(パートレシピの更新に対応するため、計2回同じ処理を行う)
    For h = 1 To 2
        file_name = Dir(folder_path & "\*.xlsx")
        Do While file_name <> ""
            Set mb = xlApp.Workbooks.Open(Filename:=folder_path & "\" & file_name, UpdateLinks:=0)
            '処理内容
            For x = 1 To 28
                ale_c(x) = 0
            Next x
            num = mb.Sheets(1).Cells(5, 10).Value
            y = 1
            For x = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 1 To mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 20
                ipoter = Replace(mb.Sheets(1).Cells(x, 1).Value, Chr(160), " ")
                ipoter = Trim(ipoter)
                For y = 1 To 9
                    Set ipoter_rng = ThisWorkbook.Sheets(sheet(y)).Columns(1).Find(ipoter)
                    If Not ipoter_rng Is Nothing Then
                        ipoter_row = ipoter_rng.Row
                        ale_in = ThisWorkbook.Sheets(sheet(y)).Cells(ipoter_row, 12).Text
                        mb.Sheets(1).Cells(x, 12) = ale_in
                        For Z = 1 To 28
                            ale_c(Z) = ale_c(Z) + InStr(ale_in, alergy(Z))
                        Next Z
                        Exit For
                    End If
                Next y
            Next x
            
            'For x = 1 To 20
                'For y = 1 To 28
                    'ale_c(y) = ale_c(y) + InStr(tw.Cells(x, "BB").Text, alergy(y))
                'Next y
            'Next x
            
            For x = 1 To 28
                If ale_c(x) <> 0 Then
                    ale = ale & " " & alergy(x)
                End If
            Next x
            
            For x = 11 To 18
                Set num_rng = ThisWorkbook.Sheets(x).Columns(2).Find(num)
                If Not num_rng Is Nothing Then
                    ThisWorkbook.Sheets(x).Cells(num_rng.Row, 6) = ale
                End If
                
                Set num_rng = ThisWorkbook.Sheets(sheet(9)).Columns(1).Find(num)
                If Not num_rng Is Nothing Then
                    ThisWorkbook.Sheets(sheet(9)).Cells(num_rng.Row, 12) = ale
                    Exit For
                End If
            Next x
            
            y = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 24
            'For x = 1 To 20
                'mb.Sheets(1).Cells(y, 12) = tw.Cells(x, "BB").Text
                'y = y + 1
            'Next x
            mb.Sheets(1).Cells(y, 1) = ale
            mb.Save
            mb.Close
            Set mb = Nothing
            ale = ""
            '処理ここまで
            file_name = Dir
        Loop
            
        For i = 1 To 10
            For k = 1 To j(i)
                
                file_name = Dir(path(k, i) & "\*.xlsx")
                Do While file_name <> ""
                    Set mb = xlApp.Workbooks.Open(Filename:=path(k, i) & "\" & file_name, UpdateLinks:=0)
                    '処理内容
                    For x = 1 To 28
                        ale_c(x) = 0
                    Next x
                    num = mb.Sheets(1).Cells(5, 10).Value
                    y = 1
                    For x = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 1 To mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 20
                        ipoter = Replace(mb.Sheets(1).Cells(x, 1).Value, Chr(160), " ")
                        ipoter = Trim(ipoter)
                        For y = 1 To 9
                            Set ipoter_rng = ThisWorkbook.Sheets(sheet(y)).Columns(1).Find(ipoter)
                            If Not ipoter_rng Is Nothing Then
                                ipoter_row = ipoter_rng.Row
                                ale_in = ThisWorkbook.Sheets(sheet(y)).Cells(ipoter_row, 12).Text
                                mb.Sheets(1).Cells(x, 12) = ale_in
                                For Z = 1 To 28
                                    ale_c(Z) = ale_c(Z) + InStr(ale_in, alergy(Z))
                                Next Z
                            Exit For
                        End If
                        Next y
                    Next x
                    
                    'For x = 1 To 20
                        'For y = 1 To 28
                            'ale_c(y) = ale_c(y) + InStr(tw.Cells(x, "BB").Text, alergy(y))
                        'Next y
                    'Next x
                    
                    For x = 1 To 28
                        If ale_c(x) <> 0 Then
                            ale = ale & " " & alergy(x)
                        End If
                    Next x
                    
                    For x = 11 To 18
                        Set num_rng = ThisWorkbook.Sheets(x).Columns(2).Find(num)
                        If Not num_rng Is Nothing Then
                            ThisWorkbook.Sheets(x).Cells(num_rng.Row, 6) = ale
                        End If
                        
                        Set num_rng = ThisWorkbook.Sheets(sheet(9)).Columns(1).Find(num)
                        If Not num_rng Is Nothing Then
                            ThisWorkbook.Sheets(sheet(9)).Cells(num_rng.Row, 12) = ale
                            Exit For
                        End If
                    Next x
                    
                    y = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 24
                    'For x = 1 To 20
                        'mb.Sheets(1).Cells(y, 12) = tw.Cells(x, "BB").Text
                        'y = y + 1
                    'Next x
                    mb.Sheets(1).Cells(y, 1) = ale
                    mb.Save
                    mb.Close
                    Set mb = Nothing
                    ale = ""
                    '処理ここまで
                    file_name = Dir
                Loop
            Next k
        Next i
    Next h
    
    Set mb = Nothing
    Set xlApp = Nothing
    
    For i = 34 To 68 Step 2
        Set cell = ThisWorkbook.Sheets("レシピ更新").Cells(i, 4)
        If cell.MergeCells Then
            Set mergedArea = cell.MergeArea
            mergedArea.ClearContents
        End If
        Set cell = ThisWorkbook.Sheets("レシピ更新").Cells(i, 10)
        If cell.MergeCells Then
            Set mergedArea = cell.MergeArea
            mergedArea.ClearContents
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Unload UserForm1
    
End Sub
Function search_menu_f(sheet)
    
    Dim ws As Worksheet
    Dim tw As Worksheet
    Dim ipoter_rng As Range
    Dim ale_rng As Range
    Dim alergy(1 To 28)
    Set ws = ThisWorkbook.Sheets(sheet)
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    alergy(1) = "えび"
    alergy(2) = "かに"
    alergy(3) = "くるみ"
    alergy(4) = "小麦"
    alergy(5) = "そば"
    alergy(6) = "卵"
    alergy(7) = "乳"
    alergy(8) = "落花生"
    alergy(9) = "アーモンド"
    alergy(10) = "あわび"
    alergy(11) = "いか"
    alergy(12) = "いくら"
    alergy(13) = "オレンジ"
    alergy(14) = "カシューナッツ"
    alergy(15) = "キウイフルーツ"
    alergy(16) = "牛肉"
    alergy(17) = "ごま"
    alergy(18) = "さけ"
    alergy(19) = "さば"
    alergy(20) = "大豆"
    alergy(21) = "鶏肉"
    alergy(22) = "バナナ"
    alergy(23) = "豚肉"
    alergy(24) = "まつたけ"
    alergy(25) = "もも"
    alergy(26) = "やまいも"
    alergy(27) = "りんご"
    alergy(28) = "ゼラチン"
    ipoter = tw.Cells(4, 3).Text
    menu = tw.Cells(4, 6).Text
    
    Set ipoter_rng = ws.Columns(1).Find(ipoter)
    
    If Not ipoter_rng Is Nothing Then
        ipoter_row = ipoter_rng.Row
        tw.Cells(10, 4) = ws.Cells(ipoter_row, 1).Text
        tw.Cells(10, 7) = ws.Cells(ipoter_row, 2).Text
        tw.Cells(10, 12) = ws.Cells(ipoter_row, 6).Text
        ale = ws.Cells(ipoter_row, 12).Text
        tw.Cells(10, 14) = ale
        tw.Cells(18, 4) = ale
        tw.Cells(4, 14) = ws.Cells(ipoter_row, 7)
        
        For i = 1 To 28
            If InStr(ale, alergy(i)) <> 0 Then
                Set ale_rng = tw.Range("D21:M28").Find(alergy(i))
                ale_r = ale_rng.Row
                ale_c = ale_rng.Column
                tw.Cells(ale_r + 1, ale_c) = "TRUE"
            End If
        Next i
        
    End If

End Function
Function update_f(sheet)

    Dim ws As Worksheet
    Dim tw As Worksheet
    Dim ipoter_rng As Range
    Set ws = ThisWorkbook.Sheets(sheet)
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    
    ipoter = tw.Cells(10, 4).Text
    Set ipoter_rng = ws.Columns(1).Find(ipoter)
    
    If Not ipoter_rng Is Nothing Then
        ipoter_row = ipoter_rng.Row
        ws.Cells(ipoter_row, 12) = tw.Cells(18, 4).Text
        ws.Cells(ipoter_row, "AS") = tw.Cells(4, "P").Text
        ws.Cells(ipoter_row, "G") = tw.Cells(4, "N").Text
    End If

End Function
Sub ShowUserFormNonModal()
    UserForm1.Caption = "処理中"
    UserForm1.Show vbModeless
End Sub
