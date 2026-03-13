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
        
        For i = 1 To 5
    
            ThisWorkbook.Sheets(sheet_name_1).Range("J" & i) = Workbooks(file_name).Sheets(1).Range("J" & i).Text
    
        Next
        
        ThisWorkbook.Sheets(sheet_name_1).Range("F8") = Workbooks(file_name).Sheets(1).Range("F7").Text
        ThisWorkbook.Sheets(sheet_name_1).Range("F10") = Workbooks(file_name).Sheets(1).Range("F9").Text
        
        For i = 14 To 33
        
            ThisWorkbook.Sheets(sheet_name_1).Range("A" & i) = Workbooks(file_name).Sheets(1).Range("A" & i - 1).Text
        
        Next
        
        Workbooks(file_name).Close
        
    Else
    
        folder_path = ThisWorkbook.Sheets(sheet_name_1).Range("V6").Text
        file_name = Dir(folder_path & "\*" & menu_name & "*.xlsx")
    
        If file_name <> "" Then
        
            Set menu_book = Workbooks.Open(Filename:=folder_path & "\" & file_name, UpdateLinks:=0)
            
            For i = 1 To 6
    
                ThisWorkbook.Sheets(sheet_name_1).Range("J" & i) = Workbooks(file_name).Sheets(1).Range("J" & i).Text

            Next
            
            ThisWorkbook.Sheets(sheet_name_1).Range("F8") = Workbooks(file_name).Sheets(1).Range("F8").Text
            ThisWorkbook.Sheets(sheet_name_1).Range("F10") = Workbooks(file_name).Sheets(1).Range("F10").Text
            
            For i = 14 To 33
            
                ThisWorkbook.Sheets(sheet_name_1).Range("A" & i) = Workbooks(file_name).Sheets(1).Range("A" & i).Text
            
            Next
            
            Workbooks(file_name).Close
            
        Else
        
            Set objFolder = objFSO.GetFolder(folder_path)
            
            For Each objSubFolder In objFolder.SubFolders
    
                file_name = Dir(objSubFolder.path & "\*" & menu_name & "*.xlsx")
                
                If file_name <> "" Then
        
                    Set menu_book = Workbooks.Open(Filename:=objSubFolder.path & "\" & file_name, UpdateLinks:=0)
                    
                    For i = 1 To 6
            
                        ThisWorkbook.Sheets(sheet_name_1).Range("J" & i) = Workbooks(file_name).Sheets(1).Range("J" & i).Text
        
                    Next
                    
                    ThisWorkbook.Sheets(sheet_name_1).Range("F8") = Workbooks(file_name).Sheets(1).Range("F8").Text
                    ThisWorkbook.Sheets(sheet_name_1).Range("F10") = Workbooks(file_name).Sheets(1).Range("F10").Text
                    
                    For i = 14 To 33
                    
                        ThisWorkbook.Sheets(sheet_name_1).Range("A" & i) = Workbooks(file_name).Sheets(1).Range("A" & i).Text
                    
                    Next
                    
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
                            
                    For i = 1 To 6
                    
                        ThisWorkbook.Sheets(sheet_name_1).Range("J" & i) = Workbooks(file_name).Sheets(1).Range("J" & i).Text
                
                    Next
                            
                    ThisWorkbook.Sheets(sheet_name_1).Range("F8") = Workbooks(file_name).Sheets(1).Range("F7").Text
                    ThisWorkbook.Sheets(sheet_name_1).Range("F10") = Workbooks(file_name).Sheets(1).Range("F9").Text
                            
                    For i = 13 To 32
                            
                        ThisWorkbook.Sheets(sheet_name_1).Range("A" & i) = Workbooks(file_name).Sheets(1).Range("A" & i - 1).Text
                            
                    Next
                            
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
