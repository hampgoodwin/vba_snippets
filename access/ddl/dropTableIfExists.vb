'' -----------------------------------------------------------------
''  - Function: Conditionally drop a table if it already exists.
''  - Note: The passed prop should never be blank, or should
'' -----------------------------------------------------------------
Function dropTableIfExists(tableName As String)
    DoCmd.SetWarnings False
        If DCount("[name]", "MSysObjects", "[Name] = '" & tableName & "'") = 1 Then
            DoCmd.RunSQL ("DROP TABLE " & tableName & ";")
        End If
    DoCmd.SetWarnings True
End Function