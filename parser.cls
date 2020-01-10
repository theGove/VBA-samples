VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "parser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
  Dim theHTML As String
  Dim ucaseHTML As String
  Dim pos As Long
  Dim region As String

Public Property Get text() As String
  text = theHTML
End Property

Public Property Let text(theValue As String)
  theHTML = theValue
  ucaseHTML = ""
End Property

Public Property Get position() As Long
  position = pos
End Property

Public Property Let position(theValue As Long)
  pos = theValue
  If pos < 1 Then pos = 1
End Property


Function multiReplace(ByVal expression As String, find As String, replacewith As String) As String
  'this function replaces multiple instances of a substing in a string with another string.
  ' it is helpful to change a arbitrary number of the same character to a single character
  ' such as converting all sequential spaces to a single space.
  
  Do While InStr(1, expression, find) > 0
    expression = Replace(expression, find, replacewith)
  Loop
  multiReplace = expression
End Function

Function removeWhiteSpace(theString) As String
Dim t As String
Dim x As Long
Dim onechar As String * 1
For x = 1 To Len(theString)
  onechar = Mid(theString, x, 1)
  If Asc(onechar) > 32 Then
    t = t & onechar
  Else
    t = t & " "
  End If
  
  Do While InStr(1, t, "  ") > 0
    t = Replace(t, "  ", " ")
  Loop
  
Next

removeWhiteSpace = t

End Function


Function replaceTags(theString As String, replaceTagsWith As String) As String
  ' this function replaces any HTML tags with the specified string.
  ' to remove tags, send a zero-length string as the second parameter
  
  Dim retval As String, dataon As Boolean, onechar As String, x As Long
  retval = ""
  dataon = True
  For x = 1 To Len(theString)
    onechar = Mid(theString, x, 1)
    If onechar = "<" Then
      retval = retval & replaceTagsWith
      dataon = False
    ElseIf onechar = ">" Then
      dataon = True
    ElseIf dataon Then
        retval = retval & onechar
    End If

  Next

  replaceTags = retval

End Function

Function getText(theString As String, Optional ignoreCase As Boolean) As String
Dim myPos As Long
  
  If ignoreCase Then
    If ucaseHTML = "" Then ucaseHTML = UCase(theHTML)
    myPos = InStr(pos, ucaseHTML, UCase(theString))
  Else
    myPos = InStr(pos, theHTML, theString)
  End If
  
  If myPos = 0 Then
    getText = ""
  Else
    getText = Mid(theHTML, pos, myPos - pos)
    pos = myPos + Len(theString)
  End If


End Function

Function moveTo(theString As String, Optional ignoreCase As Boolean, Optional ByVal theCount As Integer = 1) As Boolean
    Dim x As Integer
    moveTo = True
    For x = 1 To theCount
      If Not singleMoveTo(theString, ignoreCase) Then
         moveTo = False
         Exit Function
      End If
    Next
End Function

Private Function singleMoveTo(theString As String, Optional ignoreCase As Boolean) As Boolean
  Dim myPos As Long
  If ignoreCase Then
    If ucaseHTML = "" Then ucaseHTML = UCase(theHTML)
    myPos = InStr(pos, ucaseHTML, UCase(theString))
  Else
    myPos = InStr(pos, theHTML, theString)
  End If
  
  If myPos = 0 Then
    singleMoveTo = False
  Else
    singleMoveTo = True
    pos = myPos + Len(theString)
  End If

End Function
Function moveBackTo(theString As String) As Boolean
  Dim myPos As Long
  If ignoreCase Then
    If ucaseHTML = "" Then ucaseHTML = UCase(theHTML)
    myPos = InStrRev(ucaseHTML, UCase(theString), pos) + Len(theString)
  Else
    myPos = InStrRev(theHTML, theString, pos) + Len(theString)
  End If
  
  
    If myPos = Len(theString) Then
      moveBackTo = False
    Else
      pos = myPos
      moveBackTo = True
    End If

End Function


Function readFile(path As String) As String
  ' returns the contents of a text file
  Open path For Input As #1
    readFile = Input(LOF(1), #1)
  Close #1
End Function

Private Sub Class_Initialize()
End Sub

Private Sub Class_Terminate()
End Sub
 
 Function HTMLDecode(ByVal html As String) As String
     Dim i As Long
     
     HTMLDecode = html
     
     Do
         ' search the next ampersand, exit if no more
         i = InStr(i + 1, HTMLDecode, "&")
         If i = 0 Then Exit Do
 

         If StrComp(Mid(HTMLDecode, i, 6), "&nbsp;", vbTextCompare) = 0 Then
            HTMLDecode = left(HTMLDecode, i - 1) & " " & Mid(HTMLDecode, i + 6)
         ElseIf StrComp(Mid(HTMLDecode, i, 6), "&quot;", vbTextCompare) = 0 Then
             HTMLDecode = left(HTMLDecode, i - 1) & """" & Mid(HTMLDecode, _
                 i + 6)
         ElseIf StrComp(Mid(HTMLDecode, i, 5), "&amp;", vbTextCompare) = 0 Then
             HTMLDecode = left(HTMLDecode, i - 1) & "&" & Mid(HTMLDecode, _
                 i + 5)
         ElseIf StrComp(Mid(HTMLDecode, i, 4), "&lt;", vbTextCompare) = 0 Then
             HTMLDecode = left(HTMLDecode, i - 1) & "<" & Mid(HTMLDecode, _
                 i + 4)
         ElseIf StrComp(Mid(HTMLDecode, i, 4), "&gt;", vbTextCompare) = 0 Then
             HTMLDecode = left(HTMLDecode, i - 1) & ">" & Mid(HTMLDecode, _
                 i + 4)
         End If
     Loop
 End Function
 
 
 Function HTMLEncode(text As String) As String
     Dim i As Long
     Dim acode As Byte
     Dim repl As String
 
     HTMLEncode = text
 
     For i = Len(HTMLEncode) To 1 Step -1
         acode = Asc(Mid(HTMLEncode, i, 1))
         If acode = 32 Then
                 repl = "&nbsp;"
         ElseIf acode = 34 Then
                 repl = "&quot;"
         ElseIf acode = 38 Then
                 repl = "&amp;"
         ElseIf acode = 60 Then
                 repl = "&lt;"
         ElseIf acode = 62 Then
                 repl = "&gt;"
         ElseIf acode > 32 And acpde < 128 Then
                 ' don't touch alphanumeric chars
         Else
                 repl = "&#" & CStr(acode) & ";"
         End If
         If Len(repl) Then
             HTMLEncode = left(HTMLEncode, i - 1) & repl & Mid(HTMLEncode, i + 1)
             repl = ""
         End If
     Next
 End Function



