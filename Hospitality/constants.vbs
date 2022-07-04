'declare columns
Public Const colArtist As Integer = 2 'column with artist info
Public Const colDayPlay As Integer = 4 'column with artist playing day
Public Const colParty As Integer = 5 'column with number of members in artist parts
Public Const colRooms As Integer = 6 'colum with number of backstage rooms assigned
Public Const colDraftComplete As Integer = 10 'column, mark X when draft complete
Public Const colApproved As Integer = 12 'column, mark X with draft approved
Public Const colFinal As Integer = 16 'last column in use

'declare offsets & rows
Public Const rowArtistIndex As Integer = 2 'header row
Public Const wsOffset As Integer = 6 'first hospo sheet after indexes etc
Public Const artistOffset As Integer = 3 'first row on artist list index containing artist info (isnt this just rowArtistIndex + 1?)
public Const maxPrintLength As Integer = 100 'maximum rows per sheet

'declare some strings for naming things
'assumptions: there is one main sheet called ARTIST LIST
'there is a template hospo sheet called TEMPLATE
'there is a sheet called PLATTERS containing common platter info (eg a dozen bands get sandwich platters)
'there's maybe some other sheets we dont care about here, hence the wsOffset of 6
Public Const sheetTemplateName = "TEMPLATE"
Public Const sheetArtistName = "ARTIST LIST"
Public Const sheetPlatterName = "PLATTERS"
Public Const projName = "FESTIVAL OF VB SCRIPTS" 'name your festival / whatever here