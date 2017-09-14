'' -----------------------------------------------------------------
''  - Function: Scaffolding for saving a file saved in an attachment
''      field to a target path with a defined file name.
''  - Note: This function is file type agnostic and you should
''      provide the file type in the fileName prop. ex. "excel.xlsx"
''  - Dependencies:
''      - abelgoodwin1988/vba_snippets/io/createMultiMkdir
''      - abelgoodwin1988/vba_snippers/io/fileFolderExists
'' -----------------------------------------------------------------
Public Sub saveTemplate(sqlQry As String, filePath As String, fileName As String)
    Dim rsAttach As Recordset2
    Dim rs As Recordset2

    ' Gets the template to be saved
    Set rs = CurrentDb.OpenRecordset(sqlQry)
    Set rsAttach = rs.Fields("file").Value

    ' Delete the file if it already exists
    If ifFileFolderExists(filePath & "\" & fileName) = True Then
        Kill filePath & "\" & fileName
    End If

    ' Save template
    rsAttach.MoveFirst
    While Not rsAttach.EOF
        'On Error Resume Next
        If Dir(filePath & "\", vbDirectory) = "" Then
            makeMultiMkdir (filePath)
        End If
        rsAttach.Fields("FileData").SaveToFile filePath & "\" & fileName
        rsAttach.MoveNext
    Wend
End Sub