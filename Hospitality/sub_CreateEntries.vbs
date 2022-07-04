Public Sub CreateEntries()
    'creates entries from the template based on the artist list on main page
    Dim i, j As Integer
    Dim lastRow As Integer
    Dim bandName, nameClean, playDay As String
    Dim templateSheet, artistSheet As Worksheet
    

    Set templateSheet = Sheets(sheetTemplateName)
    Set artistSheet = Sheets(sheetArtistName)
    
    artistSheet.Activate
    
    'get last entry row
    lastRow = artistSheet.Cells(Rows.Count, colArtist).End(xlUp).Row
    
     For i = artistOffset To lastRow
        bandName = artists.Cells(i, colArtist).Value
        playDay = artists.Cells(i, colDayPlay).Value
        'remove invalid / char and truncate to 24 characters for tab names
        nameClean = Left(Replace(UCase(bandName), "/", "_"), 24)
    
        templateSheet.Copy after:=Sheets(Sheets.Count)
        'need to remove invalid characters :\/?*[]'
        ActiveSheet.Name = nameClean
        'change title on sheet - assumes this is in Range A2
        Sheets(nameClean).Range("A2").Value = bandName & " " & playDay & " Hospitality Sheet"
        
        'formatting - add days as fit?
        ' Select Case playDay
        ' Case "Friday"
        '     Sheets(nameClean).Range("A2").Font.Color = RGB(0, 176, 80)
        '     Sheets(nameClean).Tab.Color = RGB(0, 176, 80)
        ' Case "Saturday"
        '     Sheets(nameClean).Range("A2").Font.Color = RGB(255, 102, 0)
        '     Sheets(nameClean).Tab.Color = RGB(255, 102, 0)
        ' Case "Sunday"
        '     Sheets(nameClean).Range("A2").Font.Color = RGB(84, 141, 213)
        '     Sheets(nameClean).Tab.Color = RGB(84, 141, 213)
        ' End Select

        'add hyperlink
        artistSheet.Activate
        artistSheet.Hyperlinks.Add Range(Cells(i, colArtist), Cells(i, colArtist)), "", "'" & nameClean & "'!A1"
        
     Next i
     
     'fix formatting of hyperlinks
      With artistSheet.Range(Cells(artistOffset, colArtist), Cells(lastRow, colArtist))
        .Font.Name = Calibri
        .Font.Size = 11
        .Font.Underline = xlUnderlineStyleNone
        .Font.Color = RGB(0, 0, 0)
        '.BorderAround Weight:=xlMedium
        '.Borders(xlInsideHorizontal).Weight = xlThin
        .HorizontalAlignment = xlCenter
        .Interior.ColorIndex = 0
      End With
     
End Sub