Function FileExistsAbs(sPath As String)
    'Absolute PATH only
    FileExistsAbs = Dir(sPath) <> ""
End Function

Function FileExistsRel(sPath As String)
    'Relative PATH
    FileExistsRel = Dir(Application.ActiveWorkbook.Path & "\" & sPath) <> ""
End Function