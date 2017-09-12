'' -----------------------------------------------------------------
''  - Function to determine destination table when we have multiple
''      import files to be handled.
''  - Note: Additional Props can be added to further filter.
''  - Example: Want to determine destination table for not
''      only a fileName, but a sheetName as well.
'' -----------------------------------------------------------------
Function getDestinationTable(fileName As String) As String
    If fileName Like "*" Then
        getDestTable = "data"
        Exit Function
    ElseIf fileName Like "*" Then
        getDestTable = ""
        Exit Function
    Else
        getDestTable = "unknown"
        MsgBox "Unknown Source file." & vbCrLf & "File name not recognized."
        Exit Function
    End If
End Function