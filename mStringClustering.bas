Attribute VB_Name = "mStringClustering"
Option Explicit

Sub aglomerativeHierachicalClustering()

    Dim tbl As Range
    Dim dist As Double
    Dim aTkt() As Variant
    Dim aDistMatrix() As Double
    Dim aLabels() As String
    Dim n_row As Long, n_dimension As Long
    Dim i As Long, j As Long
    Dim attrLeafOrder As Variant
    Dim attrSizes As Variant
    Dim attrZ As Variant
    Dim attrZHeight As Variant
    Dim attrParents As Variant
    Dim attrClusteredItems As Variant
    Dim elem As Variant
    
    'Select the data
    Set tbl = ActiveCell.CurrentRegion
    Set tbl = tbl.Offset(1, 0).Resize(tbl.Rows.Count - 1, tbl.Columns.Count)
    aTkt = tbl.Value
    
    'get no of rows and columns
    n_row = UBound(aTkt, 1)
    n_dimension = UBound(aTkt, 2)
    
    'set the arrays to the right size
    ReDim aDistMatrix(1 To n_row, 1 To n_row)
    ReDim aLabels(1 To n_row)
    
    'get the lables
    For i = 1 To n_row
        aLabels(i) = aTkt(i, 1)
    Next i

    'calculate the distance matrix - lower triangle
    For i = 1 To n_row - 1
        For j = i + 1 To n_row
            dist = 1 - MatchPhrase(aTkt(i, 2), aTkt(j, 2), vbBinaryCompare)
            aDistMatrix(i, j) = dist
        Next j
    Next i

    'fill the distance matrix - upper triangle
    For i = 1 To n_row - 1
        For j = i + 1 To n_row
            aDistMatrix(j, i) = aDistMatrix(i, j)
        Next j
    Next i
    
    'run the hierarchical clustering
    Dim HC1 As New cHierarchical
    With HC1
        Call .NNChainLinkage(aDistMatrix, "WARD", aLabels)
        'Call .linkage(aDistMatrix, "AVERAGE", aLabels)
        Call .Optimal_leaf_ordering
    End With

    attrLeafOrder = HC1.leaf_order
    attrSizes = HC1.sizes
    attrZ = HC1.z
    attrZHeight = HC1.Z_height
    attrParents = HC1.parents
    attrClusteredItems = HC1.clustered_items
    
    ActiveSheet.Range("C2").Resize(UBound(attrClusteredItems), UBound(attrClusteredItems, 2)) = attrClusteredItems
    
    Stop
    
    
    Call HC1.Reset
    Set HC1 = Nothing


End Sub



Sub Test()

    Dim ret As Long
    Dim sim As Double
    Dim s1 As String, s2 As String
    Dim maxLen As Long
    
    
    's1 = "This is my 1st Text to test this algo."
    's2 = "This is my test no 2 text to this algo."
        
    
    If Len(s1) > Len(s2) Then
        maxLen = Len(s1)
     Else
        maxLen = Len(s2)
    End If
    
    ret = Levenshtein(s1, s2)
    sim = 1 - ret / (Len(s1) + Len(s2))
    
    Debug.Print "Levenshtein Distance is: " & ret
    Debug.Print "Text Similarity based on Levenshtein is: " & sim
    
    sim = MatchPhrase(s1, s2, vbBinaryCompare)
    Debug.Print "MatchPhrase is: " & sim
        
    
    ret = SumOfCommonStrings(s1, s2, vbTextCompare)
    sim = ret / maxLen
    Debug.Print "SumOfCommonStrings is: " & ret
    Debug.Print "Text Similarity based on SumOfCommonStrings is: " & sim
    

End Sub

Public Function MatchPhrase(ByVal Phrase1 As String, ByVal Phrase2 As String, Optional Compare As VbCompareMethod = vbTextCompare) As Double

' Function to compare two sentences. A version of this will be released to cater for the
' specific needs of matching addresses, where we can make some assumptions about common
' word-substitutions and abbreviations.

' THIS CODE IS IN THE PUBLIC DOMAIN


' This function consists of six processes:

' 1  Break out the phrases into arrays of words using the space character as the delimiter
' 2  Populate a grid of word-matching scores for each word in Phrase 1 against Phrase 2;
' 3  For each word in Phrase 1, identify the 'best match' from the words in Phrase 2
' 4  Resolve 'collisions' - two or more words in phrase 1 matching the same word in phrase 2
' 5  Compare the actual sequence of words in P1 with the positions of the matched words in P2
' 6  Weight this comparison by the degree of matching measured at the level of individual words

' Process 4, resolving collisions, is an iterative loop inside process 3
' Process 1 has an addditional step to check for deleted spaces


Dim arr1() As String            ' Phrase 1, broken out into individual words
Dim arr2() As String

Dim arrScores()    As Double    ' an array of percentage matches of each word in p1 against each word in p2

                                ' These two vectors are redundant in the sense that they hold information which
                                ' can be extracted from arrScores(). However, using them saves a lot of looping:

Dim arrPositions() As Integer   ' For each word in p1, the position of the best-matching word in p2
Dim arrSequence()  As Double    ' For each word in p1, a score for its concordance with a constructed sequence of matching words in P2


Dim n As Double                 ' should be an integer, but it will be used in floating-point
                                ' division and I prefer to avoid casting in VBA
Dim s1 As String
Dim s2 As String


Dim i As Integer
Dim j As Integer
Dim k As Integer

Dim iOffset As Integer
Dim iShift As Integer
Dim iDelete As Integer

Dim iPos As Integer
Dim jPos As Integer
Dim kPos As Integer

Dim iTotalLen As Integer

Dim dScore As Double
Dim dBest As Double
Dim dPenalty As Double

Dim d1 As Double
Dim d2 As Double

If Compare = vbTextCompare Then
    Phrase1 = UCase(Phrase1)
    Phrase2 = UCase(Phrase2)
End If

If Phrase1 = Phrase2 Then
    MatchPhrase = 1
    Exit Function
End If

' The line labels SplitSpace1 and SplitSpace2 are resynchronisation points for
' restarting the process after restoring a deleted space in Phase1 or Phrase2.

Phrase1 = StripChars(Phrase1, " ")
SplitSpace1:
arr1 = Split(Phrase1, " ")

Phrase2 = StripChars(Phrase2, " ")
SplitSpace2:
arr2 = Split(Phrase2, " ")

ReDim arrScores(LBound(arr1) To UBound(arr1), LBound(arr2) To UBound(arr2))
ReDim arrPositions(LBound(arr1) To UBound(arr1))
ReDim arrSequence(LBound(arr1) To UBound(arr1))


' Test for deleted spaces. This is a lot of work, but a missing space is a
' common error and the effects are out of all proportion to the size of the
' error: so much so that I'm prepared to risk the occasional 'false alarm'.
' It may even be worth repeating these two loops using fuzzy-matching with
' Levenshtein scores rather than the simple string-comparisons shown below:

For i = LBound(arr1) To UBound(arr1) - 1

    If arr1(i) <> "" And arr1(i + 1) <> "" Then

        s1 = arr1(i) & arr1(i + 1)

        For j = LBound(arr2) To UBound(arr2)
            If UCase(arr2(j)) = UCase(s1) Then
                Phrase2 = Substitute(Phrase2, arr2(j), arr1(i) & " " & arr1(i + 1), 1, Compare)
                GoTo SplitSpace2
            End If
        Next j

    End If ' arr(i) = "" Or arr(i + 1) = "" Then

Next i

For j = LBound(arr2) To UBound(arr2) - 1

    If arr2(j) <> "" And arr2(j + 1) <> "" Then

        s2 = arr2(j) & arr2(j + 1)

        For i = LBound(arr1) To UBound(arr1)
            If UCase(arr1(i)) = UCase(s2) Then
                Phrase1 = Substitute(Phrase1, arr1(i), arr2(j) & " " & arr2(j + 1), 1, Compare)
                GoTo SplitSpace1
            End If
         Next i

    End If

Next j


' Initialise the positions array with a negative value denoting 'not found'

For i = LBound(arr1) To UBound(arr1)
    arrPositions(i) = -1
    iTotalLen = iTotalLen + Len(arr1(i))
Next i

' For each word in Phrase 1, identify the closest matching in Phrase 2 and record its position.

For i = LBound(arr1) To UBound(arr1)

    s1 = arr1(i)
    dBest = 0
    iPos = -1

    For j = LBound(arr2) To UBound(arr2)

        s2 = arr2(j)
        dScore = 0
        dScore = MatchWord(s1, s2, Compare)

        arrScores(i, j) = dScore
        If dScore > dBest Then
            dBest = dScore
            iPos = j
        End If

    Next j

    If iPos >= 0 Then
        arrPositions(i) = iPos
    End If

Next i

' Resolve collisions - two or more words in P1 that have 'best match' scores on the same word in p2
' In theory this could be done without using the positions vector, as the information is in arrScores
' In practice, arrPositions saves processing steps

For i = LBound(arrPositions) To UBound(arrPositions)

    iPos = arrPositions(i)

    For j = i + 1 To UBound(arrPositions)

        If iPos = arrPositions(j) And iPos >= 0 Then
            ' Collision detected: which word has the best score?
            d1 = arrScores(i, iPos)
            d2 = arrScores(j, iPos)

            If d2 > d1 Then

                 'discard this recorded 'best match' position:
                arrScores(i, iPos) = -1

                'find the second-best score for d1
                dBest = 0
                kPos = -1
                For k = LBound(arrScores, 2) To UBound(arrScores, 2)
                    dScore = 0
                    dScore = arrScores(i, k)
                    If dScore > dBest Then
                        dBest = dScore
                        kPos = k
                    End If
                Next k
                 
                ' reset this conflicting position as word (i)'s match in phrase 2:
                arrPositions(j) = kPos
                 
                ' There is now a possibility that we have caused
                ' a collision with a previous word in Phrase 1:
                If k < i Then
                    For k = LBound(arrPositions) To k - 1
                        If arrPositions(k) = kPos Then
                             'restart the loop at the colliding value
                            i = k
                            j = UBound(arr1) + 1
                            Exit For
                        End If
                    Next k
                End If ' k<1

            Else

                 ' discard this recorded 'best match' position:
                arrScores(j, iPos) = -1

                 'find the second-best score for d2 *after* the current position
                dBest = 0
                kPos = -1
                For k = j + 1 To UBound(arr2)
                    dScore = 0
                    dScore = arrScores(j, k)
                    If dScore > dBest Then
                        dBest = dScore
                        kPos = k
                    End If
                Next k

                arrPositions(j) = kPos

            End If ' d2 > d1

        End If

    Next j

Next i


' Constructing a sequence-matching score:


' If we were scoring jumbled sentences of unaltered words, we'd use an edit distance algorithm;
' several are available, including replicating the Levenshtein distance at the word level. I've
' chosen a crude single-pass algorithm with a forward bias, that 'expects' the word sequence to
' resynchronise after each out-of-sequence word. It's quick, and the bias is valid - word-order
' is not neutral in real-life examples, and the heavy penalty for word transpositions reflects
' my belief that this is a more significant 'edit' than character transpositions in a word. A
' more rigorous treatment would venture into the realms of natural-language processing; that is
' out-of-scope for this application and far too ambitious for a self-contained function in VBA.

' Worked example:

' Compare two Phrases:
'  "ABC DEF GHI JKL MNO PQR STU VWX",  "ABC DEF JKL STU MNO PQR VWX"

' Variable arrPositions records the placement of each word in phrase 1 in phrase 2:

' Phrase 1            "ABC DEF GHI JKL MNO PQR STU VWX"
' Expected positions:   0   1   2   3   4   5   6   7
' Actual position in p2 0   1  -1   2   4   5   3   6

' The variable arrSequence will capture the scores
        
' Run the sequence-scoring loop:

' ABC   expected in position 0      found in 0                      Score 1/8
' DEF   expected in position 1      found in 1                      Score 1/8
' GHI   expected in position 2      DELETION     * frame shift -1 * Score NIL
' JKL   expected in position 3-1    found in 2                      Score 1/8
' MNO   expected in position 4-1    found in 4   * frame shift +1 * Score 1/8 * 7/8
' PQR   expected in position 5      found in 5                      Score 1/8
' STU   expected in position 6      found in 3   * frame shift -3 * Score 1/8 * (7/8)^3
' VWX   expected in position 7-3    found in 6   * frame shift +2 * Score 1/8 * (7/8)^2

' Edit distance is 7: the out-of-sequence penalty of 7/8 will be applied seven times

' However, we do not deal with perfectly-matched words in real life, so we cannot apply
' these penalties at the level of the entire phrase; we apply them at the level of the
' individual word, where we can apply a weighting based on each word's Levenshtein score

' The exception is deleted words; we could consider the 'word match' weighting of zero
' to be sufficient penalty but a more consistent result is obtained by applying a penalty
' to the entire phrase



' Sanity check; run the function in reverse, testing Phrase 2 against phrase 1:

' Phrase 2            "ABC DEF JKL STU MNO PQR VWX"
' Expected positions:   0   1   2   3   4   5   6
' Actual position in p1 0   1   3   6   4   5   7

' ABC   expected in position 0      found in 0                      Score 1/8
' DEF   expected in position 1      found in 1                      Score 1/8
' JKL   expected in position 2      found in 3   * frame shift +1 * Score 1/8 * 7/8
' STU   expected in position 3+1    found in 6   * frame shift +2 * Score 1/8 * (7/8)^2
' MNO   expected in position 4+3    found in 4   * frame shift -3 * Score 1/8 * (7/8)^3
' PQR   expected in position 5      found in 5                      Score 1/8
' VWX   expected in position 6      found in 7   * frame shift -1 * Score 1/8 * 7/8

' Edit distance is 7: the out-of-sequence penalty of 7/8 will be applied seven times

' "But wasn't there an insertion, too? Phrase 1 has an extra word that isn't in Phrase 2!"

' Note that our choice of denominator (8, the longer of the two wordcounts) has the effect of
' imputing a score of zero to the inserted word and applying a penalty of 7/8 to the entire phrase.

' A note on identifying the 'inserted word': actually, it's the word in Phrase 1 which didn't
' score as 'best match' against any word in Phrase 2. It could've come a close second to any or
' all of them.



If UBound(arr1) >= UBound(arr2) Then
    n = UBound(arr1) + 1
Else
    n = UBound(arr2) + 1
End If

dPenalty = 1 - (1 / n)
iShift = 0       ' Sequence distance for out-of-place words
iOffset = 0     ' Running total of this 'shift' variable
iDelete = 0     ' Count the number of deletions

For i = LBound(arrPositions) To UBound(arrPositions)

    s1 = arr1(i)

    iPos = arrPositions(i)
    iShift = iPos - i - iOffset


    Select Case iPos
    Case Is < 0     'DELETION: no matching word was found in S2

        iShift = -1
        arrSequence(i) = 0
        iDelete = iDelete + 1

    Case Is = i + iOffset ' matched word is in the expected position

        iShift = 0
        arrSequence(i) = 1 / n

    Case Else

        arrSequence(i) = (dPenalty ^ Abs(iShift)) / n

    End Select

    iOffset = iOffset + iShift

Next i

MatchPhrase = 0



For i = LBound(arrPositions) To UBound(arrPositions)
    dScore = 0
    If arrPositions(i) > -1 Then
        dScore = arrScores(i, arrPositions(i))
        dScore = dScore * arrSequence(i)
    Else
         'apply a deletion penalty - this isn't as arbitrary as it might seem: it is a equivalent to the
        '                           effect of an insertion, which acts by increasing the denominator
        dScore = -Len(arr1(i)) / iTotalLen / n
    End If
    MatchPhrase = MatchPhrase + dScore
Next i




ExitFunction:

    Erase arrScores
    Erase arrSequence
    Erase arr1
    Erase arr2

End Function



Function MatchWord(ByVal str1 As String, ByVal str2 As String, Optional Compare As VbCompareMethod = vbTextCompare) As Double

' Returns a percentage estimate of how closely word 1 matches word 2
' Edit distances exceeding the length of str1 are discarded, returning a percentage match of zero

' THIS CODE IS IN THE PUBLIC DOMAIN

Dim maxLen As Integer
Dim minLen As Integer

If Compare = vbTextCompare Then
    str1 = UCase(str1)
    str2 = UCase(str2)
End If


    If str1 = str2 Then
        MatchWord = 1
        Exit Function
    End If

    If Len(str1) > Len(str2) Then
        maxLen = Len(str1)
        minLen = Len(str2)
    Else
        maxLen = Len(str2)
        minLen = Len(str1)
    End If

    MatchWord = 0
    MatchWord = Levenshtein(str1, str2)

    If MatchWord >= minLen Then
        MatchWord = 0
    Else
        MatchWord = (maxLen - MatchWord) / maxLen
    End If

End Function



Function SumOfCommonStrings( _
                            ByVal s1 As String, _
                            ByVal s2 As String, _
                            Optional Compare As VBA.VbCompareMethod = vbTextCompare, _
                            Optional iScore As Integer = 0 _
                                ) As Integer


' N.Heffernan 06 June 2006 (somewhere over Newfoundland)
' THIS CODE IS IN THE PUBLIC DOMAIN


' Function to measure how much of String 1 is made up of substrings found in String 2

' This function uses a modified Longest Common String algorithm.
' Simple LCS algorithms are unduly sensitive to single-letter
' deletions/changes near the midpoint of the test words, eg:
' Wednesday is obviously closer to WedXesday on an edit-distance
' basis than it is to WednesXXX. So it would be better to score
' the 'Wed' as well as the 'esday' and add up the total matched

' Watch out for strings of differing lengths:
'
'    SumOfCommonStrings("Wednesday", "WednesXXXday")
'
' This scores the same as:
'
'     SumOfCommonStrings("Wednesday", "Wednesday")
'
' So make sure the calling function uses the length of the longest
' string when calculating the degree of similarity from this score.


' This is coded for clarity, not for performance.

Dim arr() As Integer    ' Scoring matrix
Dim n As Integer        ' length of s1
Dim m As Integer        ' length of s2
Dim i As Integer        ' start position in s1
Dim j As Integer        ' start position in s2
Dim subs1 As String     ' a substring of s1
Dim len1 As Integer     ' length of subs1

Dim sBefore1            ' documented in the code
Dim sBefore2
Dim sAfter1
Dim sAfter2

Dim s3 As String


SumOfCommonStrings = iScore

n = Len(s1)
m = Len(s2)

If s1 = s2 Then
    SumOfCommonStrings = n
    Exit Function
End If

If n = 0 Or m = 0 Then
    Exit Function
End If

's1 should always be the shorter of the two strings:
If n > m Then
    s3 = s2
    s2 = s1
    s1 = s3
    n = Len(s1)
    m = Len(s2)
End If

n = Len(s1)
m = Len(s2)

' Special case: s1 is n exact substring of s2
If InStr(1, s2, s1, Compare) Then
    SumOfCommonStrings = n
    Exit Function
End If

For len1 = n To 1 Step -1

    For i = 1 To n - len1 + 1

        subs1 = Mid(s1, i, len1)
        j = 0
        j = InStr(1, s2, subs1, Compare)
       
        If j > 0 Then
       
            ' We've found a matching substring...
            iScore = iScore + len1

          ' Now clip out this substring from s1 and s2...
          ' And search the fragments before and after this excision:

       
            If i > 1 And j > 1 Then
                sBefore1 = Left(s1, i - 1)
                sBefore2 = Left(s2, j - 1)
                iScore = SumOfCommonStrings(sBefore1, _
                                            sBefore2, _
                                            Compare, _
                                            iScore)
            End If
   
   
            If i + len1 < n And j + len1 < m Then
                sAfter1 = Right(s1, n + 1 - i - len1)
                sAfter2 = Right(s2, m + 1 - j - len1)
                iScore = SumOfCommonStrings(sAfter1, _
                                            sAfter2, _
                                            Compare, _
                                            iScore)
            End If
   
   
            SumOfCommonStrings = iScore
            Exit Function

        End If

    Next


Next


End Function

Function Levenshtein(ByVal string1 As String, ByVal string2 As String) As Long

Dim i As Long, j As Long, bs1() As Byte, bs2() As Byte
Dim string1_length As Long
Dim string2_length As Long
Dim distance() As Long
Dim min1 As Long, min2 As Long, min3 As Long

string1_length = Len(string1)
string2_length = Len(string2)
ReDim distance(string1_length, string2_length)
bs1 = string1
bs2 = string2

For i = 0 To string1_length
    distance(i, 0) = i
Next

For j = 0 To string2_length
    distance(0, j) = j
Next

For i = 1 To string1_length
    For j = 1 To string2_length
        'slow way: If Mid$(string1, i, 1) = Mid$(string2, j, 1) Then
        If bs1((i - 1) * 2) = bs2((j - 1) * 2) Then   ' *2 because Unicode every 2nd byte is 0
            distance(i, j) = distance(i - 1, j - 1)
        Else
            'distance(i, j) = Application.WorksheetFunction.Min _
            (distance(i - 1, j) + 1, _
             distance(i, j - 1) + 1, _
             distance(i - 1, j - 1) + 1)
            ' spell it out, 50 times faster than worksheetfunction.min
            min1 = distance(i - 1, j) + 1
            min2 = distance(i, j - 1) + 1
            min3 = distance(i - 1, j - 1) + 1
            If min1 <= min2 And min1 <= min3 Then
                distance(i, j) = min1
            ElseIf min2 <= min1 And min2 <= min3 Then
                distance(i, j) = min2
            Else
                distance(i, j) = min3
            End If

        End If
    Next
Next

Levenshtein = distance(string1_length, string2_length)

End Function

Function StripChars(myString As String, ParamArray Exceptions()) As String

' Strip out all non-alphanumeric characters from a string in a single pass
' Exceptions parameters allow you to retain specific characters (eg: spaces)

' THIS CODE IS IN THE PUBLIC DOMAIN

Dim i As Integer
Dim iLen As Integer
Dim chrA As String * 1
Dim intA As Integer
Dim j As Integer
Dim iStart As Integer
Dim iEnd As Integer

If Not IsEmpty(Exceptions()) Then
    iStart = LBound(Exceptions)
    iEnd = UBound(Exceptions)
End If

iLen = Len(myString)

For i = 1 To iLen
    chrA = Mid(myString, i, 1)
    intA = Asc(chrA)
    Select Case intA
    Case 48 To 57, 65 To 90, 97 To 122
        StripChars = StripChars & chrA
    Case Else
        If Not IsEmpty(Exceptions()) Then
            For j = iStart To iEnd
                If chrA = Exceptions(j) Then
                    StripChars = StripChars & chrA
                    Exit For ' j
                End If
            Next j
        End If
    End Select
Next i



End Function

Function Substitute(ByVal Text As String, _
                            ByVal Old_Text As String, _
                            ByVal New_Text As String, _
                            Optional Instance As Long = 0, _
                            Optional Compare As VbCompareMethod = vbTextCompare _
                            ) As String


' Replace all instances (or the nth instance ) of 'Old' text with 'New'
' Unlike VB.Mid$ this method is not sensitive to length and can replace ALL instances
' This is not exposed as a Public function because there is an Excel Worksheet function
' called Substitute(). However, Workheet Functions have length constraints.

' THIS CODE IS IN THE PUBLIC DOMAIN

Dim iStart As Long
Dim iEnd As Long
Dim iLen As Long
Dim iInstance As Long
Dim strOut As String

iLen = Len(Old_Text)

If iLen = 0 Then
    Substitute = Text
    Exit Function
End If

iEnd = 0
iStart = 1

iEnd = InStr(iStart, Text, Old_Text, Compare)

If iEnd = 0 Then
    Substitute = Text
    Exit Function
End If


strOut = ""

Do Until iEnd = 0

    strOut = strOut & Mid$(Text, iStart, iEnd - iStart)
    iInstance = iInstance + 1

    If Instance = 0 Or Instance = iInstance Then
        strOut = strOut & New_Text
    Else
        strOut = strOut & Mid$(Text, iEnd, Len(Old_Text))
    End If

    iStart = iEnd + iLen
    iEnd = InStr(iStart, Text, Old_Text, Compare)

Loop

iLen = Len(Text)
strOut = strOut & Mid$(Text, iStart, iLen - iEnd)

Substitute = strOut

End Function

Private Function Minimum(ByVal A As Integer, _
                         ByVal B As Integer, _
                         ByVal c As Integer) As Integer
Dim min As Integer

  min = A

  If B < min Then
        min = B
  End If

  If c < min Then
        min = c
  End If

  Minimum = min

End Function


