Public Sub SortByDate()
    'sorts all the rows in the main worksheet by date (and then time)
    Dim transportBook As Workbook
    Dim transportSheet As Worksheet
    Dim i, lastrow As Integer
    
    Set transportBook = ThisWorkbook
    Set transportSheet = transportBook.Worksheets(sheetPrimaryName)
    
    lastrow = transportSheet.Cells(Rows.Count, colArtist).End(xlUp).Row
    
    transportSheet.Range(Cells(artistOffset, colArtist), Cells(lastrow, colFinal)).Sort _
        key1:=Range(Cells(artistOffset, colDate), Cells(artistOffset, colDate)), _
        key2:=Range(Cells(artistOffset, colPickUp), Cells(artistOffset, colPickUp)), _
        Header:=xlNo
    
    'add a row between dates for easy reading
    For i = lastrow To artistOffset + 1 Step -1
        If transportSheet.Cells(i - 1, colDate) <> transportSheet.Cells(i, colDate) Then
            transportSheet.Rows(i).Insert
        End If
    Next i
        
End Sub