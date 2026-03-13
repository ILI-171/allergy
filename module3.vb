Sub menu_call()

    Dim file_name As String
    Dim folder_name As String
    Dim menu_name As String
    Dim folder_path As String
    Dim menu_book As Workbook
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objSubFolder As Object
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    sheet_name_1 = "レシピ  (新)"
    
    folder_path = ThisWorkbook.Sheets(sheet_name_1).Range("V4").Text
    menu_name = ThisWorkbook.Sheets(sheet_name_1).Range("Q3").Text
    file_name = Dir(folder_path & "\*" & menu_name & "*.xlsx")
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If file_name <> "" Then
    
        Set menu_book = Workbooks.Open(Filename:=folder_path & "\" & file_name, UpdateLinks:=0)
        Call copy_menu_data(file_name, sheet_name_1, 5, 7, 9, 14, 33, -1)
        
        Workbooks(file_name).Close
        
    Else
    
        folder_path = ThisWorkbook.Sheets(sheet_name_1).Range("V6").Text
        file_name = Dir(folder_path & "\*" & menu_name & "*.xlsx")
    
        If file_name <> "" Then
        
            Set menu_book = Workbooks.Open(Filename:=folder_path & "\" & file_name, UpdateLinks:=0)
            Call copy_menu_data(file_name, sheet_name_1, 6, 8, 10, 14, 33, 0)
            
            Workbooks(file_name).Close
            
        Else
        
            Set objFolder = objFSO.GetFolder(folder_path)
            
            For Each objSubFolder In objFolder.SubFolders
    
                file_name = Dir(objSubFolder.path & "\*" & menu_name & "*.xlsx")
                
                If file_name <> "" Then
        
                    Set menu_book = Workbooks.Open(Filename:=objSubFolder.path & "\" & file_name, UpdateLinks:=0)
                    Call copy_menu_data(file_name, sheet_name_1, 6, 8, 10, 14, 33, 0)
                    
                    Workbooks(file_name).Close
                    Exit For
                    
                End If
            Next
                    
            folder_path = ThisWorkbook.Sheets(sheet_name_1).Range("V4").Text
            Set objFolder = objFSO.GetFolder(folder_path)
                
            For Each objSubFolder In objFolder.SubFolders
            
                file_name = Dir(objSubFolder.path & "\*" & menu_name & "*.xlsx")
                        
                If file_name <> "" Then
                
                    Set menu_book = Workbooks.Open(Filename:=objSubFolder.path & "\" & file_name, UpdateLinks:=0)
                    Call copy_menu_data(file_name, sheet_name_1, 6, 7, 9, 13, 32, -1)
                            
                    Workbooks(file_name).Close
                    Exit For
                End If
            Next objSubFolder
        End If
    End If
    
    If file_name = "" Then
        MsgBox ("そのようなメニューは存在しません")
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Private Sub copy_menu_data(file_name As String, target_sheet As String, j_row_count As Long, _
    f8_source_row As Long, f10_source_row As Long, a_start_row As Long, a_end_row As Long, a_row_offset As Long)

    Dim fixed_targets As Variant
    Dim fixed_source_rows As Variant
    Dim idx As Long
    Dim i As Long

    fixed_targets = Array("F8", "F10")
    fixed_source_rows = Array(f8_source_row, f10_source_row)

    For i = 1 To j_row_count
        ThisWorkbook.Sheets(target_sheet).Range("J" & i) = Workbooks(file_name).Sheets(1).Range("J" & i).Text
    Next

    For idx = LBound(fixed_targets) To UBound(fixed_targets)
        ThisWorkbook.Sheets(target_sheet).Range(fixed_targets(idx)) = _
            Workbooks(file_name).Sheets(1).Range("F" & fixed_source_rows(idx)).Text
    Next

    For i = a_start_row To a_end_row
        ThisWorkbook.Sheets(target_sheet).Range("A" & i) = _
            Workbooks(file_name).Sheets(1).Range("A" & (i + a_row_offset)).Text
    Next

End Sub
