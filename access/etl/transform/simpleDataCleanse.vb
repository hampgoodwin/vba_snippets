'' -----------------------------------------------------------------
''  - Procedure: Move data from some type of import table into a
''          staging/transformation table. Perform filtering based
''          on custom where clauses.
''  - Dependencies/References:
'' -----------------------------------------------------------------
Sub simpleDataCleanse(tableNameOld As String, prepend As String, tableNameNew As String, ParamArray whereClause())
    Dim qryStrMain As String
    Dim qryStrWhere As String
    Dim qryStrRun As String
    
    Dim tableNameCreate As String
    
    Dim where As Variant
    
    tableNameCreate = prepend & tableNameNew
    
    dropTableIfExists (tableNameCreate)
    
    qryStrMain = "SELECT * INTO " & tableNameCreate & " FROM " & tableNameOld & " "
    
    If Not IsMissing(whereClause) Then
        qryStrWhere = "WHERE "
        
        For Each where In whereClause
            qryStrWhere = qryStrWhere & where & " "
        Next where
    End If
    
    qryStrWhere = qryStrWhere & ";"
                    
    qryStrRun = qryStrMain & qryStrWhere
    
    DoCmd.SetWarnings False
    DoCmd.RunSQL (qryStrRun)
    DoCmd.SetWarnings True
    
    MsgBox tableNameCreate & " has been cleansed.", , "Cleansed."
End Sub