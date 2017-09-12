'' -----------------------------------------------------------------
''  - Procedure: Having a subform object and a sourceobject, toggle
''      switch the subform sourceobject.
'' -----------------------------------------------------------------
Sub switchSubformSourceObject(subform As subform, sourceObject As String)
    DoCmd.Hourglass True
    If subform.Visible = False Or subform.sourceObject <> sourceObject Then
        subform.sourceObject = sourceObject
        subform.Visible = True
    Else
        subform.Visible = False
        subform.sourceObject = ""
    End If
    DoCmd.Hourglass False
End Sub