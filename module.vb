Sub alergy_list()

    Dim j As Long
    Dim idx As Long
    Dim sheet_name As String
    Dim alergy_range As Range
    Dim alergy_range_1 As Range
    Dim alergy_range_2 As Range
    Dim menu_sheets As Variant
    Dim row_colors As Variant
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    sheet_name = "利用可能リスト"

    ThisWorkbook.Sheets(sheet_name).Range("B6:BE600").ClearContents
    ThisWorkbook.Sheets(sheet_name).Range("B5:BE700").Interior.color = vbWhite

    alergy_column = 3

    Set alergy_range = ThisWorkbook.Sheets(sheet_name).Range("B3:AC3").Find("TRUE")
    
    If alergy_range Is Nothing Then
    
        MsgBox ("アレルギーがないため利用不要です")
        Exit Sub
    
    Else
    
        alergy_num = WorksheetFunction.CountIf(Rows(3), "TRUE")

        menu_sheets = settings_menu_category_sheets()
        row_colors = settings_menu_category_colors()

        For idx = LBound(menu_sheets) To UBound(menu_sheets)
            Call alergy_check(CStr(menu_sheets(idx)), alergy_num, row_colors(idx))
        Next
        
    End If
    
    c = 1
    
    For i = 1 To alergy_num
    
        c = ThisWorkbook.Sheets(sheet_name).Range(ThisWorkbook.Sheets(sheet_name).Cells(3, c), ThisWorkbook.Sheets(sheet_name).Cells(3, 29)).Find("TRUE").Column
        ThisWorkbook.Sheets(sheet_name).Range(ThisWorkbook.Sheets(sheet_name).Cells(5, c + 28), ThisWorkbook.Sheets(sheet_name).Cells(600, c + 28)).Interior.color = vbYellow
        
    Next
    
    ThisWorkbook.Sheets(sheet_name).Range("B6:BE600").Font.color = vbBlack
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
    MsgBox ("リストアップ完了")
    
End Sub

Function alergy_check(sheet, num, color)

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    Dim k As Long

    sheet_name = "利用可能リスト"
    
    alergy_columns_prim = ThisWorkbook.Sheets(sheet_name).Rows(3).Find("TRUE").Column
    alergy_kind_prim = ThisWorkbook.Sheets(sheet_name).Cells(2, alergy_columns_prim).Text
    
    k = ThisWorkbook.Sheets(sheet_name).Cells(Rows.Count, 11).End(xlUp).Row + 1
    dead = ThisWorkbook.Sheets(sheet).Cells(Rows.Count, 3).End(xlUp).Row
    
    acp = ThisWorkbook.Sheets(sheet).Rows(1).Find(alergy_kind_prim).Column
    num_1 = num - 1
    
    If dead = 1 Then
    
        Exit Function
    
    End If
    
    If num = 1 Then
        
        For i = 2 To dead
            
            If ThisWorkbook.Sheets(sheet).Cells(i, 3).Text = "" Then
            
            ElseIf InStr(ThisWorkbook.Sheets(sheet).Cells(i, 3), "【") > 0 Then
                Call copy_base_columns(sheet, i, sheet_name, k)
                ThisWorkbook.Sheets(sheet_name).Range("B" & k & ":AA" & k).Interior.color = vbWhite
                
                k = k + 1
            
            Else
            
                If ThisWorkbook.Sheets(sheet).Cells(i, acp).Value = "" Then
                    Call copy_base_columns(sheet, i, sheet_name, k)
                    Call copy_detail_columns(sheet, i, sheet_name, k)
                    ThisWorkbook.Sheets(sheet_name).Range("B" & k & ":AA" & k).Interior.color = color
                
                    k = k + 1
             
                End If
                
            End If
            
        Next
            
    Else
    
        For i = 2 To dead
        
            alergy_column = alergy_columns_prim
    
            If ThisWorkbook.Sheets(sheet).Cells(i, 3).Text = "" Then
            
            ElseIf InStr(ThisWorkbook.Sheets(sheet).Cells(i, 3), "【") > 0 Then
                Call copy_base_columns(sheet, i, sheet_name, k)
                ThisWorkbook.Sheets(sheet_name).Range("B" & k & ":AA" & k).Interior.color = vbWhite
                
                k = k + 1
            
            Else
            
                ac = acp
                
                For j = 1 To num_1
                
                    If ThisWorkbook.Sheets(sheet).Cells(i, ac).Text = "" Then
                
                        alergy_column = ThisWorkbook.Sheets(sheet_name).Range(ThisWorkbook.Sheets(sheet_name).Cells(3, alergy_column), ThisWorkbook.Sheets(sheet_name).Cells(3, 29)).Find("TRUE").Column
                        alergy_kind = ThisWorkbook.Sheets(sheet_name).Cells(2, alergy_column).Text
                        
                        ac = ThisWorkbook.Sheets(sheet).Rows(1).Find(alergy_kind).Column
                    
                    Else
                
                        Exit For
                    
                    End If
                        
                    If ThisWorkbook.Sheets(sheet).Cells(i, ac).Text = "" Then
                        
                        If j = num - 1 Then
                            Call copy_base_columns(sheet, i, sheet_name, k)
                            Call copy_detail_columns(sheet, i, sheet_name, k)
                            ThisWorkbook.Sheets(sheet_name).Range("B" & k & ":AA" & k).Interior.color = color
                            
                            k = k + 1
                        
                        End If
                            
                    End If
                    
                Next
            
            End If
        
        Next
        
    End If
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Function

Private Sub copy_base_columns(source_sheet As String, source_row As Long, target_sheet As String, target_row As Long)

    Dim source_columns As Variant
    Dim target_columns As Variant
    Dim idx As Long

    source_columns = settings_alergy_copy_base_source_columns()
    target_columns = settings_alergy_copy_base_target_columns()

    For idx = LBound(source_columns) To UBound(source_columns)
        ThisWorkbook.Sheets(target_sheet).Range(target_columns(idx) & target_row) = _
            ThisWorkbook.Sheets(source_sheet).Range(source_columns(idx) & source_row).Text
    Next

End Sub

Private Sub copy_detail_columns(source_sheet As String, source_row As Long, target_sheet As String, target_row As Long)

    Dim source_start_columns As Variant
    Dim source_end_columns As Variant
    Dim target_start_columns As Variant
    Dim target_end_columns As Variant
    Dim idx As Long

    source_start_columns = settings_alergy_copy_detail_source_start_columns()
    source_end_columns = settings_alergy_copy_detail_source_end_columns()
    target_start_columns = settings_alergy_copy_detail_target_start_columns()
    target_end_columns = settings_alergy_copy_detail_target_end_columns()

    For idx = LBound(source_start_columns) To UBound(source_start_columns)
        ThisWorkbook.Sheets(source_sheet).Range(source_start_columns(idx) & source_row & ":" & source_end_columns(idx) & source_row).Copy
        ThisWorkbook.Sheets(target_sheet).Range(target_start_columns(idx) & target_row & ":" & target_end_columns(idx) & target_row).PasteSpecial Paste:=xlPasteValues
    Next

End Sub
