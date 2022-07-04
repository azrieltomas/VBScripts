Public Sub SortByName()
    
    Dim transportBook As Workbook
    Dim transportSheet As Worksheet
    Dim i, lastrow As Integer
    
    Set transportBook = ThisWorkbook
    Set transportSheet = transportBook.Worksheets(sheetPrimaryName)
    
    lastrow = transportSheet.Cells(Rows.Count, colArtist).End(xlUp).Row
      
    'adds a key based on truncated artist name
    For i = artistOffset To lastrow
        transportSheet.Cells(i, colFinal + 1).Value = NameCleaner(Cells(i, colArtist).Value)
    Next i
    
    transportSheet.Range(Cells(artistOffset, colArtist), Cells(lastrow, colFinal + 1)).Sort _
        key1:=Range(Cells(artistOffset, colFinal + 1), Cells(artistOffset, colFinal + 1)), _
        key2:=Range(Cells(artistOffset, colDate), Cells(artistOffset, colDate)), _
        key3:=Range(Cells(artistOffset, colPickUp), Cells(artistOffset, colPickUp)), _
        Header:=xlNo
    
    'insert a space for easy reading
    'always gotta work backwards when inserting or deleting
    'and leave out the first row check or you get a false space
    For i = lastrow To artistOffset + 1 Step -1
        If transportSheet.Cells(i - 1, colFinal + 1).Value <> transportSheet.Cells(i, colFinal + 1).Value Then
            transportSheet.Rows(i).Insert
        End If
    Next i

    transportSheet.Columns(colFinal + 1).Delete
End Sub