Public Sub PrintSelection()
    'exports selected cells to pdf
    Dim hospBook As Workbook
    Dim artistSheet As Worksheet
    Dim saveLocation As String
    Dim cel, selectedRange As Range
    Dim sheetName, checkStr As String
    
    Set hospBook = ThisWorkbook
    Set artistSheet = hospBook.Worksheets(sheetArtistName)
    Set selectedRange = Application.Selection
    
    checkStr = "" 'better null this first
    
    If Left(Application.OperatingSystem, 7) = "Windows" Then
        saveLocation = Application.ActiveWorkbook.Path & "\"
    ElseIf Left(Application.OperatingSystem, 9) = "Macintosh" Then
        saveLocation = Application.ActiveWorkbook.Path & "/"
    End If
    
    'sanity checks
    If selectedRange.Column <> 2 Then
        MsgBox ("Select from Artist column only")
    ElseIf selectedRange.Columns.Count > 1 Then
        MsgBox ("Select only from Artist column")
    ElseIf selectedRange.Cells.Count = 1 And selectedRange.Cells(1) = "" Then
        MsgBox ("No Artist selected")
    ElseIf selectedRange.Cells.Count = 1 And selectedRange.Cells(1) = "Artist" Then
        MsgBox ("You selected the table header, not an artist")
    Else
        'first pass to determine if the whole thing is empty
        For Each cel In selectedRange
            checkStr = checkStr & cel
        Next cel
        
        'passed check
        If checkStr <> "" Then
            For Each cel In selectedRange
                If cel <> "" And cel <> "Artist" Then 'maybe you selected some empties or the Artist cell accidentally
                    sheetName = Left(Replace(UCase(cel), "/", "_"), 24)
                    hospBook.Worksheets(sheetName).ExportAsFixedFormat Type:=xlTypePDF, _
                    Filename:=saveLocation & projName " - ARTIST HOSPITALITY SHEET - " & sheetName
                End If
            Next cel
        Else
            MsgBox ("No Artists selected")
        End If
    End If
End Sub