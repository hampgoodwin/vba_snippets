'' -----------------------------------------------------------------
''  - Function: to determine last column in a tabular data set when
''      passed a long-integer row value (ex. 2147483648).
''  - Note: The passed prop should never be blank, or should
''      be the best option for getting a non-null value for the last
''      column with in the row. Generally, you should pass the row
''      representing the headers.
'' -----------------------------------------------------------------
Function lastColumn(sheetName As Excel.Worksheet, row As Long)
    lastColumn = sheetName.Cells(row, sheetName.Columns.Count).End(xlToLeft).column
End Function