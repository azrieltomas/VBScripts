Public Sub SortWorksheetsByDay()
    Dim hospBook As Workbook
    Dim artistSheet As Worksheet
    Dim i, wscount, lastRow As Integer
    Dim nameClean As String
    
    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    wscount = hospBook.Worksheets.Count

    artistSheet.Activate
    
    lastRow = artistSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'sort index into day first
        artistSheet.Range(Cells(rowArtistIndex, colArtist), Cells(lastRow, colFinal)).Sort _
        key1:=Range(Cells(rowArtistIndex, colDayPlay), Cells(rowArtistIndex, colDayPlay)), _
        key2:=Range(Cells(rowArtistIndex, colArtist), Cells(rowArtistIndex, colArtist)), Header:=xlYes
    
    'each sheet needs a specific order we can use to calculate it's place
    'put this index in the sheet somewhere then delete after
    For i = artistOffset To lastRow
        'add key to sheet > STARTS AT 6
        nameClean = Left(Replace(UCase(artistSheet.Cells(i, colArtist)), "/", "_"), 24)
        hospBook.Worksheets(nameClean).Cells(1, 8).Value = (wsOffset + i - artistOffset)
    Next i
    
    'sorting time
    For i = wsOffset To wscount - 1
        For j = i + 1 To wscount
            If hospBook.Worksheets(j).Cells(1, 8).Value < hospBook.Worksheets(i).Cells(1, 8).Value Then
                hospBook.Worksheets(j).Move before:=hospBook.Worksheets(i)
            End If
        Next j
    Next i
    
    'delete the index
    For i = wsOffset To wscount
        hospBook.Worksheets(i).Cells(1, 8).Value = ""
    Next i
    
    artistSheet.Activate

End Sub