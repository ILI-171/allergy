Sub row_adjust()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("メニュー作成")
    
    For i = 609 To 3006 Step 3
        ws.Rows(i).RowHeight = 5
    Next i

End Sub
