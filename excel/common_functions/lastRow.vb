'' -----------------------------------------------------------------
''  - Function: to determine last row in a tabular data set when
''      passed a column value (ex. "A").
''  - Note: The passed prop should never be blank, or should
''      be the best option for getting a non-null value for the last
''      row with in the column.
'' -----------------------------------------------------------------
Function lastRow(sheetName As Excel.Worksheet, column As String)
    lastRow = sheetName.Range(column & sheetName.Rows.Count).End(xlUp).row
End Function