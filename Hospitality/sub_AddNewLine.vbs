Public Sub AddNewLine()
    'sometimes you need to make a change to the template and every other sheet after its already made
    'this is a not very elegant method of doing so
    Dim i, foundRow, lastRow, wscount As Integer
    Dim insertAbove, insertText As String
    Dim hospBook As Workbook
        
    'set it up here
    insertAbove = InputBox("What line does it need to be above?")
    insertText = InputBox("What line needs to be inserted?")
    '(this is certainly inelegant but it works in a pinch)'

    Set hospBook = ThisWorkbook
    wscount = hospBook.Worksheets.Count
    
    For i = wsOffset To wscount
        lastRow = hospBook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
        
        For foundRow = 1 To lastRow
            If hospBook.Worksheets(i).Cells(foundRow, 1).Value = insertAbove Then 'its fine there's one column but use excel's FIND FUNCTIONS (no)
                Exit For
            End If
        Next foundRow
        
        hospBook.Worksheets(i).Rows(foundRow).Insert
        hospBook.Worksheets(i).Rows(foundRow).Insert
        hospBook.Worksheets(i).Cells(foundRow, 1).Value = insertText
        
    Next i
    
End Sub