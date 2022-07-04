Public Sub PrintSetup()
    Dim i, lastRow, wsCount  As Integer
    Dim hospBook As Workbook
    Dim playDay As String
              
    Set hospBook = ThisWorkbook
    wsCount = hospBook.Worksheets.Count
    
    For i = wsOffset To wsCount
        hospBook.Worksheets(i).Activate
        lastRow = hospBook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row 'good reference!
        
        With hospBook.Worksheets(i).PageSetup
            'change print area to an appropriate column as you see fit
            .PrintArea = hospBook.Worksheets(i).Range("A2:I" & lastRow).Address
            .Zoom = False
            .FitToPagesTall = Application.WorksheetFunction.RoundUp((lastRow / maxPrintLength), 0) 'maxPrintLength rows per sheet
            .FitToPagesWide = 1
            .CenterHorizontally = False 'maybe 1 works better
            .PrintTitleRows = "$2:$2" 'repeat header row
        End With
    Next i
End Sub