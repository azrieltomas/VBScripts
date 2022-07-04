Public Sub ColourApproved()
    'format the tab colours
    Dim i, lastRow As Integer
    Dim hospBook As Workbook
    Dim artistSheet As Worksheet
    
    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    
    'these are from the artist list page
    lastRow = artistSheet.Cells(Rows.Count, 2).End(xlUp).Row
        
    'if there's an X in the approved column, colour the tab
    For i = artistOffset To lastRow
        If artistSheet.Cells(i, colApproved).Value = "X" Then
            hospBook.Worksheets(i + wsOffset - artistOffset).Tab.ColorIndex = 10
        Else
            hospBook.Worksheets(i + wsOffset - artistOffset).Tab.ColorIndex = xlColorIndexNone
        End If
    Next i
End Sub