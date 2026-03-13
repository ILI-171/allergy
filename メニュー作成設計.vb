Private Sub Worksheet_Change(ByVal Target As Range)
    
    Application.ScreenUpdating = False
    
    Dim rng As Range
    Dim code_rng As Range
    Dim sheet(1 To 8)
    Dim alergy(1 To 8)
    
    sheet(1) = "冷菜　魚"
    sheet(2) = "冷菜　肉"
    sheet(3) = "冷菜　その他"
    sheet(4) = "温製　魚"
    sheet(5) = "温製　肉"
    sheet(6) = "温製　その他"
    sheet(7) = "デザート　冷製"
    sheet(8) = "デザート　温製"
    
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
            For i = 1 To 8
                Set code_rng = ThisWorkbook.Sheets(sheet(i)).Columns(1).Find(code)
                If Not code_rng Is Nothing Then
                    code_row = code_rng.Row
                    Me.Cells(r - 1, 3) = ThisWorkbook.Sheets(sheet(i)).Cells(code_row, 3).Value
                    Me.Cells(r, 3) = ThisWorkbook.Sheets(sheet(i)).Cells(code_row, 4).Value
                    For j = 13 To 21
                        Me.Cells(r, j) = ThisWorkbook.Sheets(sheet(i)).Cells(code_row, j - 6).Value
                    Next j
                    For j = 25 To 45
                        Me.Cells(r, j) = ThisWorkbook.Sheets(sheet(i)).Cells(code_row, j - 6).Value
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
            
            If menu_num > 0 And menu_num <= 50 Then
                sheetName = "メニュー表示(1-50)"
            ElseIf menu_num > 50 And menu_num <= 100 Then
                sheetName = "メニュー表示(51-100)"
            ElseIf menu_num > 100 And menu_num <= 150 Then
                sheetName = "メニュー表示(101-150)"
            ElseIf menu_num > 150 And menu_num <= 200 Then
                sheetName = "メニュー表示(151-200)"
            End If
            
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
        
        For i = 1 To 8
            alergy(i) = ""
        Next i
        
        x = 0
        For i = 14 To 21
            If Me.Cells(r, i).Value = "●" Then
                x = x + 1
                alergy(x) = Me.Cells(9, i).Text
            End If
        Next i
        
        
        ThisWorkbook.Sheets(sheetName).Range(ThisWorkbook.Sheets(sheetName).Cells(rr + 3, c), ThisWorkbook.Sheets(sheetName).Cells(rr + 7, c + 10)).ClearContents
        
        If c = 1 Then
            ThisWorkbook.Sheets(sheetName).Range("A" & rr + 5 & ":K" & rr + 5).UnMerge
        ElseIf c = 12 Then
            ThisWorkbook.Sheets(sheetName).Range("L" & rr + 5 & ":V" & rr + 5).UnMerge
        End If
        
        If x = 0 Then
            If c = 1 Then
                ThisWorkbook.Sheets(sheetName).Range("A" & rr + 5 & ":K" & rr + 5).Merge
            ElseIf c = 12 Then
                ThisWorkbook.Sheets(sheetName).Range("L" & rr + 5 & ":V" & rr + 5).Merge
            End If
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c) = "特定原材料8品目は含みません" & vbCrLf & "This menu does not contain the 8 allergenic ingredients"
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).WrapText = True
            ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).Font.Size = 11
            
            If code = "" Then
                If c = 1 Then
                    ThisWorkbook.Sheets(sheetName).Range("A" & rr + 5 & ":K" & rr + 5).UnMerge
                ElseIf c = 12 Then
                    ThisWorkbook.Sheets(sheetName).Range("L" & rr + 5 & ":V" & rr + 5).UnMerge
                End If
                ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c).ClearContents
            End If
            
            'ThisWorkbook.Sheets(sheetName).Cells(rr + 3, c + 5) = "This menu does not contain the 8 allergenic ingredients"
        ElseIf x = 1 Then
            For i = 1 To 8
                If alergy(i) <> "" Then
                    j = 5
                    ThisWorkbook.Sheets("特定原材料8").Cells(ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row, 3).Copy
                    ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c + j).PasteSpecial Paste:=xlPasteAll
                End If
            Next i
        ElseIf x = 2 Or x = 3 Then
            j = 4
            For i = 1 To 8
                If alergy(i) <> "" Then
                    ThisWorkbook.Sheets("特定原材料8").Cells(ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row, 3).Copy
                    ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c + j).PasteSpecial Paste:=xlPasteAll
                    j = j + 1
                End If
            Next i
        ElseIf x = 4 Or x = 5 Then
            j = 3
            For i = 1 To 8
                If alergy(i) <> "" Then
                    ThisWorkbook.Sheets("特定原材料8").Cells(ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row, 3).Copy
                    ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c + j).PasteSpecial Paste:=xlPasteAll
                    j = j + 1
                End If
            Next i
        ElseIf x = 6 Or x = 7 Then
            j = 2
            For i = 1 To 8
                If alergy(i) <> "" Then
                    ThisWorkbook.Sheets("特定原材料8").Cells(ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row, 3).Copy
                    ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c + j).PasteSpecial Paste:=xlPasteAll
                    j = j + 1
                End If
            Next i
        ElseIf x = 8 Then
            j = 1
            For i = 1 To 8
                If alergy(i) <> "" Then
                    ThisWorkbook.Sheets("特定原材料8").Cells(ThisWorkbook.Sheets("特定原材料8").Columns("B").Find(alergy(i)).Row, 3).Copy
                    ThisWorkbook.Sheets(sheetName).Cells(rr + 5, c + j).PasteSpecial Paste:=xlPasteAll
                    j = j + 1
                End If
            Next i
        End If
        
        ThisWorkbook.Sheets(sheetName).Protect Password:="0385"
        
    End If
    
    Application.ScreenUpdating = True

End Sub
