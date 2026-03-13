Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    Dim allergens As Variant
    Dim marker_columns As Variant
    Dim idx As Long
    Dim r As Long
    Dim source_text As String

    allergens = settings_allergens()
    marker_columns = settings_marker_columns()
    
    ' 変更されたセルがA1:A10の範囲内かどうかを確認
    If Not Intersect(Target, Me.Columns(6)) Is Nothing Then
        r = Target.Row
        Me.Range("G" & r & ":AL" & r).ClearContents
        source_text = Target.Text

        For idx = LBound(allergens) To UBound(allergens)
            If InStr(source_text, CStr(allergens(idx))) <> 0 Then
                Me.Cells(r, marker_columns(idx)) = "●"
            End If
        Next

        If InStr(source_text, "/") <> 0 Then
            Me.Cells(r, 7) = "/"
        End If
    End If
    
    Application.ScreenUpdating = True
    
End Sub


