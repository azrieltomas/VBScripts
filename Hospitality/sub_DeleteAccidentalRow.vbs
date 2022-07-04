Public Sub DeleteAccidentalRow()
    'sometimes there a line in the template that propagated to all the othersheets and its gotta go
    Dim i, foundRow, lastRow, wscount As Integer
    Dim searchString As String
    Dim hospBook As Workbook
            
    'set it up here
    searchString = InputBox("Input line to be removed")
    
    Set hospBook = ThisWorkbook
    wscount = hospBook.Worksheets.Count

    For i = wsOffset To wscount
        lastRow = hospBook.Worksheets(i).Cells(Rows.Count, 1).End(xlUp).Row
        
        For foundRow = 1 To lastRow
            If hospBook.Worksheets(i).Cells(foundRow, 1).Value = searchString Then
                hospBook.Worksheets(i).Rows(foundRow).Delete
                
                'surrounded by white space? delete another
                'the row has been deleted now so what was the next row is now foundRow
                If hospBook.Worksheets(i).Cells(foundRow, 1).Value = "" And hospBook.Worksheets(i).Cells(foundRow - 1, 1).Value = "" Then
                    hospBook.Worksheets(i).Rows(foundRow).Delete 'only if there's a spacer
                End If
    
                Exit For 'we're done here
            End If
        Next foundRow
        
    Next i
    
End Sub