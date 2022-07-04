Public Sub PrintApproved()
    'exports anything marked X in the approved column to pdf
    Dim i, lastRow As Integer
    Dim hospBook As Workbook
    Dim artistSheet As Worksheet
    Dim saveLocation As String
    
    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    
    'these are from the artist list page
    lastRow = artistSheet.Cells(Rows.Count, 2).End(xlUp).Row
        
    'apparently needs the full path now
    'you need a folder called APPROVED for this to work
    'TODO: check the folder exists and if not, put it wherever
    If Left(Application.OperatingSystem, 7) = "Windows" Then
        saveLocation = Application.ActiveWorkbook.Path & "\APPROVED\"
    ElseIf Left(Application.OperatingSystem, 9) = "Macintosh" Then
        saveLocation = Application.ActiveWorkbook.Path & "/APPROVED/"
    End If
    
    For i = artistOffset To lastRow
        If artistSheet.Cells(i, colApproved).Value = "X" Then
            hospBook.Worksheets(i + wsOffset - artistOffset).ExportAsFixedFormat Type:=xlTypePDF, _
                Filename:=saveLocation & projName & " - APPROVED " & UCase(artistSheet.Cells(i, colDayPlay).Value) & _
                    " HOSPITALITY SHEET - " & hospBook.Worksheets(i + wsOffset - artistOffset).Name
            hospBook.Worksheets(i + wsOffset - artistOffset).Tab.ColorIndex = 10
        Else
            hospBook.Worksheets(i + wsOffset - artistOffset).Tab.ColorIndex = xlColorIndexNone
        End If
    Next i

End Sub