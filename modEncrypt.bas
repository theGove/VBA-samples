Attribute VB_Name = "modEncrypt"
Option Explicit
Option Compare Binary
Public Const MIN_ASC As Integer = 32
Public Const MAX_ASC As Integer = 126
Public Const NO_OF_CHARS As Integer = MAX_ASC - MIN_ASC + 1
Function MoveAsc(ByVal a, ByVal mLvl)
    'Move the Asc value so it stays inside interval MIN_ASC and MAX_ASC
    mLvl = mLvl Mod NO_OF_CHARS
    a = a + mLvl
    If a < MIN_ASC Then
        a = a + NO_OF_CHARS
    ElseIf a > MAX_ASC Then
        a = a - NO_OF_CHARS
    End If
    MoveAsc = a
End Function
Function encrypt(ByVal s As String, ByVal key As String)
    Dim p, keyPos, c, e, k, chkSum
    If key = "" Then
        encrypt = s
        Exit Function
    End If
    For p = 1 To Len(s)
        If Asc(Mid(s, p, 1)) < MIN_ASC Or Asc(Mid(s, p, 1)) > MAX_ASC Then
            MsgBox "Char at position " & p & " is invalid!"
            Exit Function
        End If
    Next p
    For keyPos = 1 To Len(key)
        chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s)
        c = Asc(Mid(s, p, 1))
        keyPos = keyPos + 1
        If keyPos > Len(key) Then keyPos = 1
        k = Asc(Mid(key, keyPos, 1))
        c = MoveAsc(c, k)
        c = MoveAsc(c, k * Len(key))
        c = MoveAsc(c, chkSum * k)
        c = MoveAsc(c, p * k)
        c = MoveAsc(c, Len(s) * p) 'This is only for getting new chars for different word lengths
        e = e & Chr(c)
    Next p
    encrypt = e
End Function
Function decrypt(ByVal s As String, ByVal key As String)
    Dim p, keyPos, c, d, k, chkSum
    If key = "" Then
        decrypt = s
        Exit Function
    End If
    For keyPos = 1 To Len(key)
        chkSum = chkSum + Asc(Mid(key, keyPos, 1)) * keyPos
    Next keyPos
    keyPos = 0
    For p = 1 To Len(s)
        c = Asc(Mid(s, p, 1))
        keyPos = keyPos + 1
        If keyPos > Len(key) Then keyPos = 1
        k = Asc(Mid(key, keyPos, 1))
        'Do MoveAsc in reverse order from encrypt, and with a minus sign this time(to unmove)
        c = MoveAsc(c, -(Len(s) * p))
        c = MoveAsc(c, -(p * k))
        c = MoveAsc(c, -(chkSum * k))
        c = MoveAsc(c, -(k * Len(key)))
        c = MoveAsc(c, -k)
        d = d & Chr(c)
    Next p
    decrypt = d
End Function

