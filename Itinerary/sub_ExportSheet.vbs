Public Sub ExportSheet()
    Dim bookPrimary, bookSecondary As Workbook
    Dim sheetPrimary, sheetSecondary As Worksheet
    Dim lastrow, i, j As Integer
    Dim artistName, cleanName, saveLocation As String
    Dim shp As Shape
    Dim nameFound, saveExport As Boolean
        
    'sucks!
    If Left(Application.OperatingSystem, 7) = "Windows" Then
        saveLocation = Application.ActiveWorkbook.Path & "\"
    ElseIf Left(Application.OperatingSystem, 9) = "Macintosh" Then
        saveLocation = Application.ActiveWorkbook.Path & "/"
    Else
        MsgBox ("Unable to determine Operating System. Cannot save file.")
    End If
    
    Set bookPrimary = ThisWorkbook
    Set sheetPrimary = bookPrimary.Worksheets(sheetPrimaryName)
    nameFound = False
    
    'does this need to be saved?
    If sheetPrimary.CheckBoxes("CBExport").Value = 1 Then 'alternatively you could use checkboxes(1) if its the only one, otherwise be certain about this name
        saveExport = True
    Else
        saveExport = False
    End If
    
    'truncate the names in an elaborate way using spaces and THEs
    'the idea here is you get rid of all the A PARTY etc parts for sorting
    'eg "THE BEATLES - A PARTY" and "THE BEATLES - B PARTY" just gets sorted as "THE BEATLES"
    artistName = Range(rngArtistName).Value
    cleanName = NameCleaner(artistName)

    'first find if the artist is even in here yet
    lastrow = sheetPrimary.Cells(Rows.Count, colArtist).End(xlUp).Row
    For i = artistOffset To lastrow
        If NameCleaner(sheetPrimary.Cells(i, colArtist).Value) = cleanName Then
            nameFound = True
        End If
    Next i
    
    If nameFound Then
        Set bookSecondary = Workbooks.Add
        sheetPrimary.Copy after:=bookSecondary.Worksheets(1)
        
        Application.DisplayAlerts = False 'stops the "are you sure you want to delete this" prompt
        bookSecondary.Worksheets(1).Delete
        Application.DisplayAlerts = True
        
        Set sheetSecondary = bookSecondary.Worksheets(sheetPrimaryName)
        'give it a new name
        sheetSecondary.Name = Left("ITINERARY SHEET " & UCase(artistName), 24)
        
        sheetSecondary.Cells(1, 1).Value = projName & " ITINERARY " & UCase(artistName) 'puts a title A1
        sheetSecondary.Cells(2, 1).Value = "AS @ " & Day(Date) & " " & MonthName(Month(Date)) 'puts the current date in A2
        
        'now delete the ones that arent correct
        For i = lastrow To artistOffset Step -1 'lastrow is the same as primary sheet, work backwards now
            If NameCleaner(sheetSecondary.Cells(i, colArtist).Value) <> cleanName Then
                sheetSecondary.Rows(i).Delete
            End If
        Next i
        
        'sort by name then date then time
        lastrow = sheetSecondary.Cells(Rows.Count, colArtist).End(xlUp).Row
        sheetSecondary.Range(Cells(artistOffset, colArtist), Cells(lastrow, colFinal)).Sort _
            key1:=Range(Cells(artistOffset, colDate), Cells(artistOffset, colDate)), _
            key2:=Range(Cells(artistOffset, colPickUp), Cells(artistOffset, colPickUp)), _
            Header:=xlNo
        
        'delete the export info
        sheetSecondary.Range(rngArtistName).Validation.Delete
        sheetSecondary.Range(rngArtistName).Value = artistName
        For Each shp In sheetSecondary.Shapes
            shp.Delete
        Next shp
        
        'save in the same folder as the main workbook
        If saveExport Then
            bookSecondary.SaveAs Filename:=saveLocation & projName & " ITINERARY SHEET - " & UCase(artistName)
        End If
        
    Else
        MsgBox ("No entries for " & artistName)
    End If
    
End Sub