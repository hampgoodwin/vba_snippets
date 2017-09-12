'' -----------------------------------------------------------------
''  - Function: to create microsoft office file dialog picker.
''  - Note: This function only supports selecting a single file.
''  - Dependencies:
''      - Microsoft Office 15.0 Object Library
'' -----------------------------------------------------------------
Function selectFile()
    Dim fd As FileDialog, fileName As String
     
    On Error GoTo ErrorHandler
     
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
     
    fd.AllowMultiSelect = False
     
    If fd.Show = True Then
        If fd.SelectedItems(1) <> vbNullString Then
            fileName = fd.SelectedItems(1)
        End If
    Else
        'Exit code if no file is selected
        End
    End If

    selectFile = fileName
     
    Set fd = Nothing
    
    Exit Function
     
ErrorHandler:
    Set fd = Nothing
    MsgBox "Error " & Err & ": " & Error(Err)
End Function