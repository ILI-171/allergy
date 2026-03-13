Sub alergy_list()

    Dim j As Long
    Dim sheet_name As String
    Dim alergy_range As Range
    Dim alergy_range_1 As Range
    Dim alergy_range_2 As Range
    
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
        
        Call alergy_check("冷菜　魚", alergy_num, rgbSkyBlue)
        Call alergy_check("冷菜　肉", alergy_num, rgbSkyBlue)
        Call alergy_check("冷菜　その他", alergy_num, rgbSkyBlue)
        Call alergy_check("温製　魚", alergy_num, rgbSandyBrown)
        Call alergy_check("温製　肉", alergy_num, rgbSandyBrown)
        Call alergy_check("温製　その他", alergy_num, rgbSandyBrown)
        Call alergy_check("デザート　冷製", alergy_num, rgbPaleGreen)
        Call alergy_check("デザート　温製", alergy_num, rgbPaleGreen)
        
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
            
                ThisWorkbook.Sheets(sheet_name).Range("B" & k) = ThisWorkbook.Sheets(sheet).Range("A" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("E" & k) = ThisWorkbook.Sheets(sheet).Range("B" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("K" & k) = ThisWorkbook.Sheets(sheet).Range("C" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("Q" & k) = ThisWorkbook.Sheets(sheet).Range("D" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("AA" & k) = ThisWorkbook.Sheets(sheet).Range("E" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("B" & k & ":AA" & k).Interior.color = vbWhite
                
                k = k + 1
            
            Else
            
                If ThisWorkbook.Sheets(sheet).Cells(i, acp).Value = "" Then
             
                    ThisWorkbook.Sheets(sheet_name).Range("B" & k) = ThisWorkbook.Sheets(sheet).Range("A" & i).Text
                    ThisWorkbook.Sheets(sheet_name).Range("E" & k) = ThisWorkbook.Sheets(sheet).Range("B" & i).Text
                    ThisWorkbook.Sheets(sheet_name).Range("K" & k) = ThisWorkbook.Sheets(sheet).Range("C" & i).Text
                    ThisWorkbook.Sheets(sheet_name).Range("Q" & k) = ThisWorkbook.Sheets(sheet).Range("D" & i).Text
                    ThisWorkbook.Sheets(sheet_name).Range("AA" & k) = ThisWorkbook.Sheets(sheet).Range("E" & i).Text
                    ThisWorkbook.Sheets(sheet).Range("H" & i & ":O" & i).Copy
                    ThisWorkbook.Sheets(sheet_name).Range("AD" & k & ":AK" & k).PasteSpecial Paste:=xlPasteValues
                    ThisWorkbook.Sheets(sheet).Range("S" & i & ":AL" & i).Copy
                    ThisWorkbook.Sheets(sheet_name).Range("AL" & k & ":BE" & k).PasteSpecial Paste:=xlPasteValues
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
            
                ThisWorkbook.Sheets(sheet_name).Range("B" & k) = ThisWorkbook.Sheets(sheet).Range("A" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("E" & k) = ThisWorkbook.Sheets(sheet).Range("B" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("K" & k) = ThisWorkbook.Sheets(sheet).Range("C" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("Q" & k) = ThisWorkbook.Sheets(sheet).Range("D" & i).Text
                ThisWorkbook.Sheets(sheet_name).Range("AA" & k) = ThisWorkbook.Sheets(sheet).Range("E" & i).Text
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
                        
                            ThisWorkbook.Sheets(sheet_name).Range("B" & k) = ThisWorkbook.Sheets(sheet).Range("A" & i).Text
                            ThisWorkbook.Sheets(sheet_name).Range("E" & k) = ThisWorkbook.Sheets(sheet).Range("B" & i).Text
                            ThisWorkbook.Sheets(sheet_name).Range("K" & k) = ThisWorkbook.Sheets(sheet).Range("C" & i).Text
                            ThisWorkbook.Sheets(sheet_name).Range("Q" & k) = ThisWorkbook.Sheets(sheet).Range("D" & i).Text
                            ThisWorkbook.Sheets(sheet_name).Range("AA" & k) = ThisWorkbook.Sheets(sheet).Range("E" & i).Text
                            ThisWorkbook.Sheets(sheet).Range("H" & i & ":O" & i).Copy
                            ThisWorkbook.Sheets(sheet_name).Range("AD" & k & ":AK" & k).PasteSpecial Paste:=xlPasteValues
                            ThisWorkbook.Sheets(sheet).Range("S" & i & ":AL" & i).Copy
                            ThisWorkbook.Sheets(sheet_name).Range("AL" & k & ":BE" & k).PasteSpecial Paste:=xlPasteValues
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
