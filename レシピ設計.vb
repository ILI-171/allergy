Private Sub Worksheet_Change(ByVal Target As Range)
    Application.ScreenUpdating = False
    Dim cell As Range
    
    For Each cell In Target
        If Not Intersect(cell, Me.Range("A14:A33")) Is Nothing Then
            r = cell.Row
            Me.Cells(37, 1) = ""
            Me.Cells(r, 2) = ""
            Me.Cells(r, 7) = ""
            Me.Cells(r, 8) = ""
            Me.Cells(r, 9) = ""
            Me.Cells(r, 10) = ""
            Me.Cells(r, 11) = ""
            Me.Cells(r, 12) = ""
            Me.Cells(r, 15) = ""
            ipoter = cell.Value
            If ipoter <> "" Then
                Call smf("業者１", ipoter, r)
                Call smf("業者２", ipoter, r)
                Call smf("業者３", ipoter, r)
                Call smf("業者４", ipoter, r)
                Call smf("業者５", ipoter, r)
                Call smf("業者６", ipoter, r)
                Call smf("業者７", ipoter, r)
                Call smf("業者８", ipoter, r)
                Call smf("パートレシピ", ipoter, r)
            Else
                Me.Range("B" & r & ":O" & r).ClearContents
            End If
            For i = 14 To 33
                If InStr(Me.Cells(i, 12).Value, "えび") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "えび") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "えび"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "かに") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "かに") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "かに"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "くるみ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "くるみ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "くるみ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "小麦") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "小麦") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "小麦"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "そば") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "そば") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "そば"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "卵") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "卵") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "卵"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "乳") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "乳") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "乳"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "落花生") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "落花生") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "落花生"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "アーモンド") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "アーモンド") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "アーモンド"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "あわび") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "あわび") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "あわび"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "いか") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "いか") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "いか"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "いくら") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "いくら") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "いくら"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "オレンジ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "オレンジ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "オレンジ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "カシューナッツ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "カシューナッツ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "カシューナッツ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "キウイフルーツ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "キウイフルーツ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "キウイフルーツ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "牛肉") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "牛肉") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "牛肉"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "ごま") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "ごま") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "ごま"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "さけ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "さけ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "さけ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "さば") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "さば") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "さば"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "大豆") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "大豆") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "大豆"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "鶏肉") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "鶏肉") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "鶏肉"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "バナナ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "バナナ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "バナナ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "豚肉") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "豚肉") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "豚肉"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "まつたけ") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "まつたけ") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "まつたけ"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "もも") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "もも") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "もも"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "やまいも") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "やまいも") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "やまいも"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "りんご") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "りんご") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "りんご"
                    End If
                End If
                If InStr(Me.Cells(i, 12).Value, "ゼラチン") <> 0 Then
                    If InStr(Me.Cells(37, 1).Value, "ゼラチン") = 0 Then
                        Me.Cells(37, 1) = Me.Cells(37, 1).Value & " " & "ゼラチン"
                    End If
                End If
            Next i
        End If
    Next
    
    Application.ScreenUpdating = True
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
        tw.Cells(r, 2) = ws.Cells(ip_row, 2).Value
        tw.Cells(r, 7) = ws.Cells(ip_row, 6).Value
        tw.Cells(r, 9) = ws.Cells(ip_row, 7).Value
        tw.Cells(r, 10) = ws.Cells(ip_row, 10).Value
        'tw.Cells(r, 11) = ws.Cells(ip_row, 2).Value
        tw.Cells(r, 12) = ws.Cells(ip_row, 12).Value
        tw.Cells(r, 15) = ws.Cells(ip_row, "AS").Value
    End If
End Function


