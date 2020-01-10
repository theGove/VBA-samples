Attribute VB_Name = "mod_IE_Tools"
Option Explicit


' pauses the execution of your VBA code while waiting for IE to finish loading
Sub waitFor(ie As InternetExplorer)
    Do
        Do
            Application.Wait Now + TimeValue("00:00:01")
            attach ie
            DoEvents
        Loop Until Not ie.Busy And ie.readystate = 4
        Application.Wait Now + TimeValue("00:00:01")
    Loop Until Not ie.Busy And ie.readystate = 4
End Sub



'Connect IE to the most recently opened Internet Explorer windows
' if urlPart is supplied, will only attach on to an explorer that has that string as a part of the URL
Function attach(ie As Object, Optional urlPart As String) As Boolean
    Dim o As Object
    Dim x As Long
    Dim explorers As Object
    Dim name As String
    Set explorers = CreateObject("Shell.application").Windows
    For x = explorers.Count - 1 To 0 Step -1
       name = "Empty"
       On Error Resume Next
       name = explorers((x)).name
       On Error GoTo 0
       If name = "Internet Explorer" Then
          If InStr(1, explorers((x)).LocationURL, urlPart, vbTextCompare) Then
               Set ie = explorers((x))
               attach = True
               Exit For
          End If
       End If
    Next
    
End Function

' Returns the number of the HTML element specified by tagname and identifying text
Function getTagNumber(ie As InternetExplorer, tagName As String, Optional identifyingText As String, Optional startAtTagNumber As Long = 0) As Long
  Dim x As Long
  Dim t As Object
  For x = startAtTagNumber To ie.document.all.Length - 1
     Set t = ie.document.all(x)
     If UCase(t.tagName) = UCase(tagName) Then
        'we found the right kind of tag, check to see if it has the right text
        'Debug.Print t.outerHTML
        If InStr(1, t.outerhtml, identifyingText) > 0 Then
          'we found the right kind fo tag with the right identifying text, return the number
          getTagNumber = x
          Exit Function
        End If
     End If
  Next
  getTagNumber = -1 ' sentinal value indicating the tag was not found
End Function

'returns a reference to a tag object given its number.  Used in conjunction with GetTagNumber
Function getTag(ie As InternetExplorer, tagNumber) As Object
  Set getTag = ie.document.all(tagNumber)
End Function

Public Sub showpage(ie As InternetExplorer)
    savePage ie
    ThisWorkbook.FollowHyperlink ThisWorkbook.path & "\source.html"
End Sub


Public Sub savePage(ie As InternetExplorer, Optional filePath As String)
  'saves a local copy of the document in Internet Explorer as currently rendered
  Dim x As Long
  Dim len1 As Long
  Dim len2 As Long
  
  Dim ff As Integer
  ff = FreeFile
    If filePath = "" Then
       Open ThisWorkbook.path & "\source.html" For Output As #ff
   Else
       Open filePath For Output As #ff
   End If
   
   For x = 0 To ie.document.all.Length - 1
     Print #ff, ie.document.all(x).outerhtml
     If UCase(ie.document.all(x).tagName) = "HTML" Then Exit For
   Next
   
   Close #ff
End Sub

'Uses the WebQuery Wizard to import data from the current page in IE
Public Sub importPage(ie As InternetExplorer, newSheetName As String, Optional wb As Workbook)
  Dim ff As Integer
  Dim s As Worksheet
  
  If TypeName(wb) = "Nothing" Then
    Set wb = ThisWorkbook
  End If
  
  ff = FreeFile
  
   Open ThisWorkbook.path & "\localWebPageAgentFile.html" For Output As #ff
   Print #ff, "<html><head><title>Saved Page</title></head>"
   Print #ff, ie.document.body.outerhtml
   Print #ff, "</html>"
   Close #ff

  
  
  On Error Resume Next
    Application.DisplayAlerts = False
       wb.Sheets(newSheetName).Delete
    Application.DisplayAlerts = True
  On Error GoTo 0
  
  
  Set s = wb.Worksheets.Add
  s.name = newSheetName
  
      With s.QueryTables.Add(Connection:= _
        "URL;file:///" & Replace(ThisWorkbook.path, "\", "/") & "/localWebPageAgentFile.html", Destination:=s.Range("$A$1"))
        .name = "localWebPageAgentFile"
        .FieldNames = True
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .savePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .WebSelectionType = xlEntirePage
        .WebFormatting = xlWebFormattingNone
        .WebPreFormattedTextToColumns = True
        .WebConsecutiveDelimitersAsOne = True
        .WebSingleBlockTextImport = False
        .WebDisableDateRecognition = False
        .WebDisableRedirections = False
        .Refresh BackgroundQuery:=False
    End With
    
    s.QueryTables(1).Delete
    
  Kill ThisWorkbook.path & "\localWebPageAgentFile.html"
  
End Sub


