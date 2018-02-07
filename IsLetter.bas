Private Declare Function IsCharAlphaW Lib "user32" (ByVal cChar As Integer) As Long

Public Property Get IsLetter(character As String) As Boolean
    IsLetter = IsCharAlphaW(AscW(character))
End Property


'-------------------------------------------------------------------

Function IsLetter(strValue As String) As Boolean
    Dim intPos As Integer
    For intPos = 1 To Len(strValue)
        Select Case Asc(Mid(strValue, intPos, 1))
            Case 65 To 90, 97 To 122
                IsLetter = True
            Case Else
                IsLetter = False
                Exit For
        End Select
    Next
End Function