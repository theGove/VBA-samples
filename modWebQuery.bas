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

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False _
) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

