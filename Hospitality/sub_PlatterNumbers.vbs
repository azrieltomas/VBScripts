Public Sub PlatterNumbersToThisBook()
    'so named because of PlatterNumbersToNewBook (removed)
    Dim hospBook As Workbook
    Dim artistSheet, platterSheet As Worksheet
    Dim i, j, wscount, lastRow, useRowPlatter, infoRowStart As Integer
    Dim colOffset, colDayOne, colDayTwo, colDayThree, colPlatLast As Integer

    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    Set platterSheet = hospBook.Worksheets(sheetPlatterName)
    wscount = hospBook.Worksheets.Count

    'columns for each day and where they start
    'TODO: cleanup
    colDayOne = 1
    colDayTwo = 7
    colDayThree = 13
    colPlatLast = 17
    infoRowStart = 4 'where info starts after the header
    
    'clear the existing page - needs to be on this sheet
    platterSheet.Activate
    'just get a big enough number and after the headerrow
    lastRow = platterSheet.Cells(Rows.Count, colDayOne).End(xlUp).Row + platterSheet.Cells(Rows.Count, colDayTwo).End(xlUp).Row + _
        platterSheet.Cells(Rows.Count, colDayThree).End(xlUp).Row + infoRowStart
    'and wipe it out
    platterSheet.Range(Cells(infoRowStart, 1), Cells(lastRow, colPlatLast)).ClearContents 'this is why plattersheet.activate
      
    For i = wsOffset To wscount
        lastRow = hospBook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
        
        'day affects the column
        If artistSheet.Cells(i - wsOffset + artistOffset, colDayPlay).Value = "Friday" Then
            colOffset = colDayOne - 1
        ElseIf artistSheet.Cells(i - wsOffset + artistOffset, colDayPlay).Value = "Saturday" Then
            colOffset = colDayTwo - 1
        ElseIf artistSheet.Cells(i - wsOffset + artistOffset, colDayPlay).Value = "Sunday" Then
            colOffset = colDayThree - 1
        End If
        
        'find me some platters
        For j = 1 To lastRow
            If instr(1, hospBook.Worksheets(i).Cells(j, 2).Value, "Platter") > 0 Then
                useRowPlatter = platterSheet.Cells(Rows.Count, colOffset + 1).End(xlUp).Row + 1 'not great but works
                
                platterSheet.Cells(useRowPlatter, colOffset + 1).Value = artistSheet.Cells(i - wsOffset + artistOffset, colArtist).Value  'band name
                platterSheet.Cells(useRowPlatter, colOffset + 3).Value = hospBook.Worksheets(i).Cells(j, 1).Value  'quantity
                platterSheet.Cells(useRowPlatter, colOffset + 4).Value = artistSheet.Cells(i - wsOffset + artistOffset, colDayPlay).Value 'day
                
                'formatting, just on evens to cheat
                'NO. THIS STINKS.
                'If i Mod 2 Then
                '    platterSheet.Range(Cells(rowCounter, 1), Cells(rowCounter, 5)).Interior.Color = RGB(217, 225, 242)
                'End If
            End If
        Next j
    Next i
    
End Sub