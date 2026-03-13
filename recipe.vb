Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    Dim cell As Range
    Dim vendor_sheets As Variant
    Dim clear_columns As Variant
    Dim idx As Long

    vendor_sheets = settings_vendor_sheets(True)
    clear_columns = settings_recipe_clear_columns()
    
    For Each cell In Target
        If Not Intersect(cell, Me.Range("A14:A33")) Is Nothing Then
            r = cell.Row
            Me.Cells(37, 1) = ""
            For idx = LBound(clear_columns) To UBound(clear_columns)
                Me.Cells(r, clear_columns(idx)) = ""
            Next

            ipoter = cell.Value

            If ipoter <> "" Then
                For idx = LBound(vendor_sheets) To UBound(vendor_sheets)
                    Call smf(CStr(vendor_sheets(idx)), ipoter, r)
                Next
            Else
                Me.Range("B" & r & ":O" & r).ClearContents
            End If
            Call collect_allergens(Me, 14, 33, 12, 37, 1)
        End If
    Next
    
    Application.ScreenUpdating = True
End Sub

Private Sub collect_allergens(ws As Worksheet, start_row As Long, end_row As Long, source_column As Long, target_row As Long, target_column As Long)

    Dim allergens As Variant
    Dim i As Long
    Dim idx As Long
    Dim allergen_name As String

    allergens = settings_allergens()

    For i = start_row To end_row
        For idx = LBound(allergens) To UBound(allergens)
            allergen_name = CStr(allergens(idx))
            If InStr(ws.Cells(i, source_column).Value, allergen_name) <> 0 Then
                If InStr(ws.Cells(target_row, target_column).Value, allergen_name) = 0 Then
                    ws.Cells(target_row, target_column) = ws.Cells(target_row, target_column).Value & " " & allergen_name
                End If
            End If
        Next
    Next

End Sub

Function smf(sheet, ip, r)
    
    Dim ws As Worksheet
    Dim tw As Worksheet
    Dim ipoter_rng As Range
    Dim menu_rng As Range
    Set ws = ThisWorkbook.Sheets(sheet)
    Set tw = ThisWorkbook.Sheets("レシピ  (新)")
    
    last = ws.Cells(Rows.Count, 3).End(xlUp).Row
    
    If last = 1 Then
        Exit Function
    End If
    
    Set ipoter_rng = ws.Columns(1).Find(ip)
    
    If Not ipoter_rng Is Nothing Then
        ip_row = ipoter_rng.Row
        source_columns = settings_recipe_source_columns()
        target_columns = settings_recipe_target_columns()

        For idx = LBound(source_columns) To UBound(source_columns)
            tw.Cells(r, target_columns(idx)) = ws.Cells(ip_row, source_columns(idx)).Value
        Next
    End If
End Function


