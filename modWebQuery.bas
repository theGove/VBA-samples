Attribute VB_Name = "modWebQuery"
Option Explicit


Public Sub CreateWebQuery(Destination As Range, URL As String, Optional WebSelectionType As XlWebSelectionType = xlEntirePage, Optional SaveQuery As Boolean, Optional PlainText As Boolean = True)

  '*********************************************************************************'
  '         Builds a web-query object to retrieve information from a web server
  '
  '    Parameters:
  '
  '        Destination
  '          a reference to a cell where the query output will begin
  '
  '        URL
  '          The webpage to get. Should start with "http"
  '
  '        WebSelectionType (xlEntirePage or xlAllTables)
  '          what part of the page should be brought back to Excel.
  '
  '        SaveQuery (True or False)
  '          Indicates if the query object remains in the workbook after running
  '
  '        PlainText (True or False)
  '          Indicates if the query results should be plain or include formatting
  '
  '*********************************************************************************'
  
      With Destination.Parent.QueryTables.Add(Connection:="URL;" & URL, Destination:=Destination)
        .Name = "WebQuery"
        .RefreshStyle = xlOverwriteCells
        .WebSelectionType = WebSelectionType
        .PreserveFormatting = PlainText
        .BackgroundQuery = False
        .Refresh
        If Not SaveQuery Then .Delete
    End With
    
End Sub



Public Sub webTablesOnCell()
    ' builds a web query looking in the activecell for the URL and returning the tables
    ' from the page in the cell below the active cell
    
    If LCase(Left(ActiveCell.Value, 4)) = "http" Then
      CreateWebQuery ActiveCell.Offset(1), ActiveCell.Value, xlAllTables, False, True
    End If
    
End Sub

Public Sub webPageOnCell()
    ' builds a web query looking in the activecell for the URL and returning the tables
    ' from the page in the cell below the active cell
    
    If LCase(Left(ActiveCell.Value, 4)) = "http" Then
      CreateWebQuery ActiveCell.Offset(1), ActiveCell.Value, xlEntirePage, False, True
    End If
    
End Sub
