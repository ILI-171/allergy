Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    
    ' 変更されたセルがA1:A10の範囲内かどうかを確認
    If Not Intersect(Target, Me.Columns(6)) Is Nothing Then
        r = Target.Row
        Me.Range("G" & r & ":AL" & r).ClearContents
        If InStr(Target.Text, "えび") <> 0 Then
            Me.Cells(r, 8) = "●"
        End If
        If InStr(Target.Text, "かに") <> 0 Then
            Me.Cells(r, 9) = "●"
        End If
        If InStr(Target.Text, "くるみ") <> 0 Then
            Me.Cells(r, 10) = "●"
        End If
        If InStr(Target.Text, "小麦") <> 0 Then
            Me.Cells(r, 11) = "●"
        End If
        If InStr(Target.Text, "そば") <> 0 Then
            Me.Cells(r, 12) = "●"
        End If
        If InStr(Target.Text, "卵") <> 0 Then
            Me.Cells(r, 13) = "●"
        End If
        If InStr(Target.Text, "乳") <> 0 Then
            Me.Cells(r, 14) = "●"
        End If
        If InStr(Target.Text, "落花生") <> 0 Then
            Me.Cells(r, 15) = "●"
        End If
        If InStr(Target.Text, "アーモンド") <> 0 Then
            Me.Cells(r, 19) = "●"
        End If
        If InStr(Target.Text, "あわび") <> 0 Then
            Me.Cells(r, 20) = "●"
        End If
        If InStr(Target.Text, "いか") <> 0 Then
            Me.Cells(r, 21) = "●"
        End If
        If InStr(Target.Text, "いくら") <> 0 Then
            Me.Cells(r, 22) = "●"
        End If
        If InStr(Target.Text, "オレンジ") <> 0 Then
            Me.Cells(r, 23) = "●"
        End If
        If InStr(Target.Text, "カシューナッツ") <> 0 Then
            Me.Cells(r, 24) = "●"
        End If
        If InStr(Target.Text, "キウイフルーツ") <> 0 Then
            Me.Cells(r, 25) = "●"
        End If
        If InStr(Target.Text, "牛肉") <> 0 Then
            Me.Cells(r, 26) = "●"
        End If
        If InStr(Target.Text, "ごま") <> 0 Then
            Me.Cells(r, 27) = "●"
        End If
        If InStr(Target.Text, "さけ") <> 0 Then
            Me.Cells(r, 28) = "●"
        End If
        If InStr(Target.Text, "さば") <> 0 Then
            Me.Cells(r, 29) = "●"
        End If
        If InStr(Target.Text, "大豆") <> 0 Then
            Me.Cells(r, 30) = "●"
        End If
        If InStr(Target.Text, "鶏肉") <> 0 Then
            Me.Cells(r, 31) = "●"
        End If
        If InStr(Target.Text, "バナナ") <> 0 Then
            Me.Cells(r, 32) = "●"
        End If
        If InStr(Target.Text, "豚肉") <> 0 Then
            Me.Cells(r, 33) = "●"
        End If
        If InStr(Target.Text, "まつたけ") <> 0 Then
            Me.Cells(r, 34) = "●"
        End If
        If InStr(Target.Text, "もも") <> 0 Then
            Me.Cells(r, 35) = "●"
        End If
        If InStr(Target.Text, "やまいも") <> 0 Then
            Me.Cells(r, 36) = "●"
        End If
        If InStr(Target.Text, "りんご") <> 0 Then
            Me.Cells(r, 37) = "●"
        End If
        If InStr(Target.Text, "ゼラチン") <> 0 Then
            Me.Cells(r, 38) = "●"
        End If
        If InStr(Target.Text, "/") <> 0 Then
            Me.Cells(r, 7) = "/"
        End If
    End If
    
    Application.ScreenUpdating = True
    
End Sub


