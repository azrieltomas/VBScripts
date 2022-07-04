Sub SortWorksheetsByAlpha()
    'self explanatory'
    Dim hospBook As Workbook
    Dim artistSheet As Worksheet
    Dim i, j, wscount, lastRow As Integer
    
    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    wscount = hospBook.Worksheets.Count
    
    artistSheet.Activate
    
    lastRow = artistSheet.Cells(Rows.Count, 2).End(xlUp).Row
    
    'sort index into regular alphabet
    artistSheet.Range(Cells(rowArtistIndex, colArtist), Cells(lastRow, colFinal)).Sort _
            key1:=Range(Cells(rowArtistIndex, colArtist), Cells(rowArtistIndex, colArtist)), _
            Header:=xlYes
    
    'move all the sheets around to match
    For i = wsOffset To wscount - 1
        For j = i + 1 To wscount
            If hospBook.Worksheets(j).Name < hospBook.Worksheets(i).Name Then
                hospBook.Worksheets(j).Move before:=hospBook.Worksheets(i)
            End If
        Next j
    Next i
    
    artistSheet.Activate
    
End Sub