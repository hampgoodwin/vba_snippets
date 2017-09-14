'' -----------------------------------------------------------------
''  - Procedure: OOB Mkdir functionality does not provide creation
''      of multi-level folder structures from within VBA. This
''      procedure pased a multi-level path will create all paths to
''      the end folder.
'' -----------------------------------------------------------------
Public Sub makeMultiMkdir(sPath As String)
    Dim iStart As Integer
    Dim aDirs As Variant
    Dim sCurDir As String
    Dim i As Integer

    If sPath <> "" Then
        aDirs = Split(sPath, "\")
        If Left(sPath, 2) = "\\" Then
            iStart = 3
        Else
            iStart = 1
        End If

        sCurDir = Left(sPath, InStr(iStart, sPath, "\"))

        For i = iStart To UBound(aDirs)
            sCurDir = sCurDir & aDirs(i) & "\"
            If Dir(sCurDir, vbDirectory) = vbNullString Then
                MkDir sCurDir
            End If
        Next i
    End If
End Sub