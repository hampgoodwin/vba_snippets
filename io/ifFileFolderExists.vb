'' -----------------------------------------------------------------
''  - Function: Return boolean for if a file or folder exists given
''      a path.
'' -----------------------------------------------------------------
Public Function ifFileFolderExists(strFullPath As String) As Boolean
    If Not Dir(strFullPath, vbDirectory) = vbNullString Then fileFolderExists = True
End Function