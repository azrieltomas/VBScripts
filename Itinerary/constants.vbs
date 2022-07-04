'declare columns
Public Const colArtist As Integer = 1 'column listing the artist
Public Const colDate As Integer = 6 'column listing date of transfer
Public Const colPickUp As Integer = 11 'pick up time
Public Const colFinal As Integer = 18 'last column in use

'declare rows
Public Const artistOffset As Integer = 10 'first row that contains real data after the headers

'some names and ranges
public const sheetPrimaryName as string = "ITINERARY PRIMARY" 'main itinerary sheet
public const sheetDataName as string = "DATA" 'sheet for storing some values, like all the band names
public const rngArtistName as string = ""I2"" ' range that stores the artist you want to export, put it in a header section
public const projName as string = "BIG FESTIVAL" 'this is the name of your project / festival / whatever

'requirements in addition to above:
'3 buttons: export, sort by name, sort by date
'1 checkbox: named CBExport, determines the exported copy should be saved
'the DATA sheet contains a named range called ARTISTLIST. Self explanatory (?)