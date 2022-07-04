Function NameCleaner(ByVal nameToClean As String) As String
' cleans up any value passed to it for easy sorting
' if it starts with "THE" it becomes "THE + FIRSTNAMEPART" otherwise just "FIRSTNAMEPART"
    Dim cleanedName As String
        If Left(UCase(nameToClean), 3) = "THE" Then
            If InStr(5, nameToClean, " ") > 0 Then
                cleanedName = UCase(Left(nameToClean, InStr(5, nameToClean, " ") - 1))
            Else
                cleanedName = UCase(nameToClean)
            End If
        ElseIf InStr(1, nameToClean, " ") = 0 Then 'no spaces
            cleanedName = UCase(nameToClean)
        ElseIf nameToClean = "" Then 'empties
            cleanedName = ""
        Else
            cleanedName = UCase(Left(nameToClean, InStr(1, nameToClean, " ") - 1))
        End If

    NameCleaner = cleanedName
    
End Function