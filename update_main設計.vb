Sub search_menu()
'既製品の検索
    Dim sheet_names As Variant
    Dim i As Long

    sheet_names = get_vendor_sheets(False)

    For i = LBound(sheet_names) To UBound(sheet_names)
        Call search_menu_f(CStr(sheet_names(i)))
    Next

End Sub
Sub checkbox()
'チェックボックスの内容を文字に直す
    Dim tw As Worksheet
    Dim check_rows As Variant
    Dim label_rows As Variant
    Dim end_columns As Variant
    Dim idx As Long

    Set tw = ThisWorkbook.Sheets("レシピ更新")
    Count = 0
    ale = ""

    check_rows = Array(22, 25, 28)
    label_rows = Array(21, 24, 27)
    end_columns = Array(11, 13, 13)

    For idx = LBound(check_rows) To UBound(check_rows)
        Call collect_checked_allergens(tw, check_rows(idx), label_rows(idx), 4, end_columns(idx), ale, Count)
    Next

    If Count = 28 Then
        ale = "/"
    End If
    
    tw.Cells(18, 4) = ale

End Sub
Sub all()
'チェックボックスのチェックをすべて外す
    Application.ScreenUpdating = False

    Dim tw As Worksheet
    Dim clear_rows As Variant
    Dim end_columns As Variant
    Dim idx As Long

    Set tw = ThisWorkbook.Sheets("レシピ更新")
    
    ale = ""

    clear_rows = Array(22, 25, 28)
    end_columns = Array(11, 13, 13)

    For idx = LBound(clear_rows) To UBound(clear_rows)
        Call clear_checkbox_row(tw, clear_rows(idx), 4, end_columns(idx))
    Next
    
    Call checkbox
    
    tw.Cells(18, 4) = "/"

    Application.ScreenUpdating = True
    
End Sub
Sub update()
'既製品アレルギー内容の更新

    Dim tw As Worksheet
    Dim sheet_names As Variant
    Dim i As Long

    Set tw = ThisWorkbook.Sheets("レシピ更新")
    tw.Cells(10, 14) = tw.Cells(18, 4).Text

    sheet_names = get_vendor_sheets(False)
    For i = LBound(sheet_names) To UBound(sheet_names)
        Call update_f(CStr(sheet_names(i)))
    Next
    
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
    Dim alergy As Variant
    Dim ale_c(1 To 28)
    Dim sheet_names As Variant

    alergy = get_allergen_list()
    sheet_names = get_vendor_sheets(True)
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
            For x = LBound(alergy) To UBound(alergy)
                ale_c(x) = 0
            Next x
            num = mb.Sheets(1).Cells(5, 10).Value
            y = 1
            For x = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 1 To mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 20
                ipoter = Replace(mb.Sheets(1).Cells(x, 1).Value, Chr(160), " ")
                ipoter = Trim(ipoter)
                For y = LBound(sheet_names) To UBound(sheet_names)
                    Set ipoter_rng = ThisWorkbook.Sheets(sheet_names(y)).Columns(1).Find(ipoter)
                    If Not ipoter_rng Is Nothing Then
                        ipoter_row = ipoter_rng.Row
                        ale_in = ThisWorkbook.Sheets(sheet_names(y)).Cells(ipoter_row, 12).Text
                        mb.Sheets(1).Cells(x, 12) = ale_in
                        For Z = LBound(alergy) To UBound(alergy)
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
            
            For x = LBound(alergy) To UBound(alergy)
                If ale_c(x) <> 0 Then
                    ale = ale & " " & alergy(x)
                End If
            Next x
            
            For x = 11 To 18
                Set num_rng = ThisWorkbook.Sheets(x).Columns(2).Find(num)
                If Not num_rng Is Nothing Then
                    ThisWorkbook.Sheets(x).Cells(num_rng.Row, 6) = ale
                End If
                
                Set num_rng = ThisWorkbook.Sheets(sheet_names(UBound(sheet_names))).Columns(1).Find(num)
                If Not num_rng Is Nothing Then
                    ThisWorkbook.Sheets(sheet_names(UBound(sheet_names))).Cells(num_rng.Row, 12) = ale
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
                    For x = LBound(alergy) To UBound(alergy)
                        ale_c(x) = 0
                    Next x
                    num = mb.Sheets(1).Cells(5, 10).Value
                    y = 1
                    For x = mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 1 To mb.Sheets(1).Columns(1).Find("IPOTER番号").Row + 20
                        ipoter = Replace(mb.Sheets(1).Cells(x, 1).Value, Chr(160), " ")
                        ipoter = Trim(ipoter)
                        For y = LBound(sheet_names) To UBound(sheet_names)
                            Set ipoter_rng = ThisWorkbook.Sheets(sheet_names(y)).Columns(1).Find(ipoter)
                            If Not ipoter_rng Is Nothing Then
                                ipoter_row = ipoter_rng.Row
                                ale_in = ThisWorkbook.Sheets(sheet_names(y)).Cells(ipoter_row, 12).Text
                                mb.Sheets(1).Cells(x, 12) = ale_in
                                For Z = LBound(alergy) To UBound(alergy)
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
                    
                    For x = LBound(alergy) To UBound(alergy)
                        If ale_c(x) <> 0 Then
                            ale = ale & " " & alergy(x)
                        End If
                    Next x
                    
                    For x = 11 To 18
                        Set num_rng = ThisWorkbook.Sheets(x).Columns(2).Find(num)
                        If Not num_rng Is Nothing Then
                            ThisWorkbook.Sheets(x).Cells(num_rng.Row, 6) = ale
                        End If
                        
                        Set num_rng = ThisWorkbook.Sheets(sheet_names(UBound(sheet_names))).Columns(1).Find(num)
                        If Not num_rng Is Nothing Then
                            ThisWorkbook.Sheets(sheet_names(UBound(sheet_names))).Cells(num_rng.Row, 12) = ale
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
    Dim alergy As Variant
    Set ws = ThisWorkbook.Sheets(sheet)
    Set tw = ThisWorkbook.Sheets("レシピ更新")
    alergy = get_allergen_list()
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
        
        For i = LBound(alergy) To UBound(alergy)
            If InStr(ale, alergy(i)) <> 0 Then
                Set ale_rng = tw.Range("D21:M28").Find(alergy(i))
                ale_r = ale_rng.Row
                ale_c = ale_rng.Column
                tw.Cells(ale_r + 1, ale_c) = "TRUE"
            End If
        Next i
        
    End If

End Function

Private Sub collect_checked_allergens(tw As Worksheet, check_row As Long, label_row As Long, _
    start_column As Long, end_column As Long, ByRef ale As String, ByRef item_count As Long)

    Dim i As Long

    For i = start_column To end_column
        If tw.Cells(check_row, i).Value = True Then
            ale = ale & " " & tw.Cells(label_row, i).Text
        Else
            item_count = item_count + 1
        End If
    Next

End Sub

Private Sub clear_checkbox_row(tw As Worksheet, row_num As Long, start_column As Long, end_column As Long)

    Dim i As Long

    For i = start_column To end_column
        tw.Cells(row_num, i) = False
    Next

End Sub

Private Function get_vendor_sheets(include_part_recipe As Boolean) As Variant

    If include_part_recipe Then
        get_vendor_sheets = Array("業者１", "業者２", "業者３", "業者４", "業者５", "業者６", "業者７", "業者８", "パートレシピ")
    Else
        get_vendor_sheets = Array("業者１", "業者２", "業者３", "業者４", "業者５", "業者６", "業者７", "業者８")
    End If

End Function

Private Function get_allergen_list() As Variant

    get_allergen_list = Array("えび", "かに", "くるみ", "小麦", "そば", "卵", "乳", "落花生", "アーモンド", "あわび", "いか", "いくら", "オレンジ", "カシューナッツ", "キウイフルーツ", "牛肉", "ごま", "さけ", "さば", "大豆", "鶏肉", "バナナ", "豚肉", "まつたけ", "もも", "やまいも", "りんご", "ゼラチン")

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
