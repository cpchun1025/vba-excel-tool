Function ReadTemplateFromFile(templatePath As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(templatePath, 1)
    ReadTemplateFromFile = ts.ReadAll
    ts.Close
End Function