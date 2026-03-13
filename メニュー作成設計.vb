Private Sub Worksheet_Change(ByVal Target As Range)
    
    Application.ScreenUpdating = False
    
    Dim rng As Range
    Dim code_rng As Range
    Dim menu_sheets As Variant
    Dim alergy(1 To 8)

    menu_sheets = settings_menu_category_sheets()
    
    Set rng = Me.Range("B11")
    
    For i = 14 To 608 Step 3
        Set rng = Union(rng, Me.Range("B" & i))
    Next i
    
    If Not Intersect(Target, rng) Is Nothing Then
        r = Target.Row
        code = Target.Value
        
        If code = "タイトル" Then
            ThisWorkbook.Sheets("サービス用").Cells(r - 1, 2) = Me.Cells(r - 1, 3).Value
            ThisWorkbook.Sheets("セール用").Cells(r - 2, 1) = Me.Cells(r - 1, 3).Value
        Else
            Me.Range("C" & r & ":AS" & r - 1).ClearContents
            For i = LBound(menu_sheets) To UBound(menu_sheets)
                Set code_rng = ThisWorkbook.Sheets(menu_sheets(i)).Columns(1).Find(code)
                If Not code_rng Is Nothing Then
                    code_row = code_rng.Row
                    Me.Cells(r - 1, 3) = ThisWorkbook.Sheets(menu_sheets(i)).Cells(code_row, 3).Value
                    Me.Cells(r, 3) = ThisWorkbook.Sheets(menu_sheets(i)).Cells(code_row, 4).Value
                    For j = 13 To 21
                        Me.Cells(r, j) = ThisWorkbook.Sheets(menu_sheets(i)).Cells(code_row, j - 6).Value
                    Next j
                    For j = 25 To 45
                        Me.Cells(r, j) = ThisWorkbook.Sheets(menu_sheets(i)).Cells(code_row, j - 6).Value
                    Next j
                    Exit For
                End If
            Next i
            
            ThisWorkbook.Sheets("サービス用").Cells(r - 1, 2) = Me.Cells(r - 1, 3).Value
            ThisWorkbook.Sheets("サービス用").Cells(r, 2) = Me.Cells(r, 3).Value
            
            For i = 11 To 43
                ThisWorkbook.Sheets("サービス用").Cells(r, i) = Me.Cells(r, i + 2).Value
            Next i
            
            ThisWorkbook.Sheets("セール用").Cells(r - 2, 1) = Me.Cells(r - 1, 3).Value
            ThisWorkbook.Sheets("セール用").Cells(r - 1, 1) = Me.Cells(r, 3).Value
            
            menu_num = Me.Cells(r - 2, 1).Value
            If menu_num = "" Then
                menu_num = 1
            End If
            
            Dim menu_display_sheets As Variant
            menu_display_sheets = settings_menu_display_sheets()
            sheetName = CStr(menu_display_sheets((menu_num - 1) \ 50))
            
            ThisWorkbook.Sheets(sheetName).Unprotect Password:="0385"
            
            If Not ThisWorkbook.Sheets(sheetName).Columns("A").Find(What:=menu_num, LookAt:=xlWhole) Is Nothing Then
                rr = ThisWorkbook.Sheets(sheetName).Columns("A").Find(What:=menu_num, LookAt:=xlWhole).Row
                c = 1
                ThisWorkbook.Sheets(sheetName).Cells(rr + 1, 1) = Me.Cells(r - 1, 3).Value
                ThisWorkbook.Sheets(sheetName).Cells(rr + 2, 1) = Me.Cells(r, 3).Value
            ElseIf Not ThisWorkbook.Sheets(sheetName).Columns("L").Find(What:=menu_num, LookAt:=xlWhole) Is Nothing Then
                rr = ThisWorkbook.Sheets(sheetName).Columns("L").Find(What:=menu_num, LookAt:=xlWhole).Row
                c = 12
                ThisWorkbook.Sheets(sheetName).Cells(rr + 1, "L") = Me.Cells(r - 1, 3).Value
                ThisWorkbook.Sheets(sheetName).Cells(rr + 2, "L") = Me.Cells(r, 3).Value
            End If
        
        End If
        
        Erase alergy
        
        x = 0
        For i = 14 To 21
            If Me.Cells(r, i).Value = "●" Then
                x = x + 1
                alergy(x) = Me.Cells(9, i).Text
            End If
        Next i
        
        
        ThisWorkbook.Sheets(sheetName).Range(ThisWorkbook.Sheets(sheetName).Cells(rr + 3, c), ThisWorkbook.Sheets(sheetName).Cells(rr + 7, c + 10)).ClearContents
        
        ThisWorkbook.Sheets(sheetName).Range(merge_range_str(rr + 5, c)).UnMerge
        
        If x = 0 Then
            ThisWorkbook.Sheets(sheetName).Range(merge_range_str(rr + 5, c)).Merge
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c) = "特定原材料8品目は含みません" & vbCrLf & "This menu does not contain the 8 allergenic ingredients"
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).WrapText = True
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).Font.Size = 11
            
            If code = "" Then
                ThisWorkbook.Sheets(sheetName).Range(merge_range_str(rr + 5, c)).UnMerge
                ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).ClearContents
            End If
            
            'ThisWorkbook.Sheets(sheetName).Cells(rr + 3, c + 5) = "This menu does not contain the 8 allergenic ingredients"
        Else
            j = 6 - ((x + 1) \ 2)
            Call paste_allergen_icons(alergy, sheetName, rr, c, j)
        End If
        
        ThisWorkbook.Sheets(sheetName).Protect Password:="0385"
        
    End If
    
    Application.ScreenUpdating = True

End Sub

Private Sub paste_allergen_icons(alergy As Variant, target_sheet As String, target_row As Long, target_column As Long, start_offset As Long)

    Dim i As Long
    Dim j As Long
    Dim source_row As Long

    j = start_offset

    For i = LBound(alergy) To UBound(alergy)
        If alergy(i) <> "" Then
            source_row = ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row
            ThisWorkbook.Sheets("特定原材料8").Cells(source_row, 3).Copy
            ThisWorkbook.Sheets(target_sheet).Cells(target_row + 5, target_column + j).PasteSpecial Paste:=xlPasteAll
            j = j + 1
        End If
    Next i

End Sub

Private Function merge_range_str(row_num As Long, c As Long) As String

    Dim col_letters As Variant
    Dim parts As Variant

    col_letters = settings_merge_col_ranges()
    parts = Split(CStr(col_letters(IIf(c = 1, 0, 1))), ":")
    merge_range_str = parts(0) & row_num & ":" & parts(1) & row_num

End Function
