Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    Dim allergens As Variant
    Dim marker_columns As Variant
    Dim idx As Long
    Dim r As Long
    Dim source_text As String

    allergens = Array("えび", "かに", "くるみ", "小麦", "そば", "卵", "乳", "落花生", "アーモンド", "あわび", "いか", "いくら", "オレンジ", "カシューナッツ", "キウイフルーツ", "牛肉", "ごま", "さけ", "さば", "大豆", "鶏肉", "バナナ", "豚肉", "まつたけ", "もも", "やまいも", "りんご", "ゼラチン")
    marker_columns = Array(8, 9, 10, 11, 12, 13, 14, 15, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38)
    
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


