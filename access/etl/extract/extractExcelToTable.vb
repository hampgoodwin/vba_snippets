'' -----------------------------------------------------------------
''  - Procedure: Perform extraction of an excel file to an access
''      table.
''  - References:
''      - Microsoft Excel 15.0 Object Library
''  - Dependencies:
''      - abelgoodwin1988/vba_snippets/io/selectFile
''      - abelgoodwin1988/vba_snippets/access...
''          io/getDestinationTable
''      - abelgoodwin1988/vba_snippets/excel/...
''          common_functions/lastRow
''      - abelgoodwin1988/vba_snippets/excel/...
''          common_functions/lastColumn
'' -----------------------------------------------------------------
Sub extractExcelToTable(filePath As String)
'' Source File Paths
    Dim srcFilePath As String, srcFileName As String
'' Excel objects
    Dim xlapp As Excel.Application
    Dim xlwb As Excel.Workbook
    Dim xlws As Excel.Worksheet
    Dim xlrng As Excel.Range
'' Destination Information
    Dim destinationTableName As String

    '' Assign passed string fileName parameter to local var, and get fileName
    srcFilePath = filePath
    srcFileName = Right(srcFilePath, Len(srcFilePath) - InStrRev(srcFilePath, "\"))

    '' Pass srcFileName to helper function to determine name of destination table.
    destinationTableName = getDestinationTable(srcFileName)
    If destinationTableName = "unknown" Then GoTo cleanup

    '' Set excel application, workbook, worksheet, and range objects.
        '' Note: You may want to customize what sheet is selected.
    Set xlapp = New Excel.Application
        'xlapp.Visible = True '' use for debugging
    Set xlwb = xlapp.Workbooks.Open(srcFilePath, False, True, , "password")
    Set xlws = xlwb.Worksheets(1)
        'uncomment xlrng if you intend to import a specific range here and in TransferSpreadsheet
    'Set xlrng = xlws.Range(xlws.Cells({row}, {column}), xlws.Cells(lastRow({sheet object}, {column as string}), lastColumn({sheet object}, {row as long})))

    '' Drop destination table, if it exists
    dropTableIfExists ("import_" & destinationTableName)

    DoCmd.TransferSpreadsheet acImport, , "import_" & destinationTableName, srcFilePath, True ', Replace(xlws.Name & "!" & xlrng.Address, "$", "")

    DoCmd.SetWarnings True
    
'' Cleanup
cleanup:
    On Error Resume Next
    Set xlrng = Nothing

    Set xlws = Nothing

    xlwb.Close False
    Set xlwb = Nothing

    xlapp.Quit
    Set xlapp = Nothing

    MsgBox "Import Completed." & vbCrLf & vbCrLf & "Imported File: " & srcFileName & vbCrLf & "To Table: import_" & destinationTableName, , "Import"
End Sub