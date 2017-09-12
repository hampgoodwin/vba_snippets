'' -----------------------------------------------------------------
''  - Procedure: Trim all data within a table.
'' -----------------------------------------------------------------
Sub trimAllColumns(tableName As String)
    Dim tdf    As DAO.TableDef
    Dim dbs    As DAO.Database
    Dim fld    As DAO.Field

    Set dbs = CurrentDb
    Set tdf = dbs.TableDefs(tableName)

    For Each fld In tdf.Fields
        If fld.Type = dbText Then
            dbs.Execute "UPDATE " & tableName & " SET [" & fld.Name & "]=Trim([" & fld.Name & "])"
        End If
    Next fld
End Sub