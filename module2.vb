Sub all_clear()

    Application.ScreenUpdating = False

    sheet_name = "利用可能リスト"

    For i = 2 To 29
    
        ThisWorkbook.Sheets(sheet_name).Cells(3, i) = "FALSE"
    
    Next
    
    Application.ScreenUpdating = True

End Sub
