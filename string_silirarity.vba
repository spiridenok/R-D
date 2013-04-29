'Implements functions to rate how similar two strings are on
'a scale of 0.0 (completely dissimilar) to 1.0 (exactly similar)
'Source:   http://www.catalysoft.com/articles/StrikeAMatch.html
'Author: Bob Chatham, bob.chatham at gmail.com
'9/12/2010
'DSPI: Modified to suite my needs (mostly the types of my variables, the algo is not change)

Option Explicit

Public Function stringSimilarity(str1 As String, str2 As String) As Variant
'Simple version of the algorithm that computes the similiarity metric
'between two strings.
'NOTE: This verision is not efficient to use if you're comparing one string
'with a range of other values as it will needlessly calculate the pairs for the
'first string over an over again; use the array-optimized version for this case.

    Dim sPairs1 As Collection
    Dim sPairs2 As Collection

    Set sPairs1 = New Collection
    Set sPairs2 = New Collection

    WordLetterPairs str1, sPairs1
    WordLetterPairs str2, sPairs2

    stringSimilarity = SimilarityMetric(sPairs1, sPairs2)

    Set sPairs1 = Nothing
    Set sPairs2 = Nothing

End Function

Public Function strSimA(str1 As Variant, rRng As Range) As Variant
'Return an array of string similarity indexes for str1 vs every string in input range rRng
    Dim sPairs1 As Collection
    Dim sPairs2 As Collection
    Dim arrOut As Variant
    Dim l As Long, j As Long

    Set sPairs1 = New Collection

    WordLetterPairs CStr(str1), sPairs1

    l = rRng.Count
    ReDim arrOut(1 To l)
    For j = 1 To l
        Set sPairs2 = New Collection
        WordLetterPairs CStr(rRng(j)), sPairs2
        arrOut(j) = SimilarityMetric(sPairs1, sPairs2)
        Set sPairs2 = Nothing
    Next j

    strSimA = Application.Transpose(arrOut)

End Function

'Public Function strSimLookup(str1 As Variant, rRng As Range, Optional returnType) As Variant
Public Function strSimLookup(str1 As Variant, rRng As Variant, Optional returnType) As Variant
'Return either the best match or the index of the best match
'depending on returnTYype parameter) between str1 and strings in rRng)
' returnType = 0 or omitted: returns the best matching string
' returnType = 1           : returns the index of the best matching string
' returnType = 2           : returns the similarity metric

    Dim sPairs1 As Collection
    Dim sPairs2 As Collection
    Dim metric, bestMetric As Double
    Dim i, iBest As String
    Const RETURN_STRING As Integer = 0
    Const RETURN_INDEX As Integer = 1
    Const RETURN_METRIC As Integer = 2

    If IsMissing(returnType) Then returnType = RETURN_STRING

    Set sPairs1 = New Collection

    WordLetterPairs CStr(str1), sPairs1

    bestMetric = -1
    iBest = ""

    For Each i In rRng
        Set sPairs2 = New Collection
        WordLetterPairs i, sPairs2
        metric = SimilarityMetric(sPairs1, sPairs2)
        If metric > bestMetric Then
            bestMetric = metric
            iBest = i
        End If
        Set sPairs2 = Nothing
    Next i

    If iBest = "" Then
        strSimLookup = CVErr(xlErrValue)
        Exit Function
    End If

    Select Case returnType
    Case RETURN_STRING
        strSimLookup = iBest
    Case RETURN_INDEX
        strSimLookup = iBest
    Case Else
        strSimLookup = bestMetric
    End Select

End Function

Public Function strSim(str1 As String, str2 As String) As Variant
    Dim ilen, iLen1, ilen2 As Integer

    iLen1 = Len(str1)
    ilen2 = Len(str2)

    If iLen1 >= ilen2 Then ilen = ilen2 Else ilen = iLen1

    strSim = stringSimilarity(Left(str1, ilen), Left(str2, ilen))

End Function

Sub WordLetterPairs(ByVal str As String, pairColl As Collection)
'Tokenize str into words, then add all letter pairs to pairColl

    Dim Words() As String
    Dim word, nPairs, pair As Integer

    Words = Split(str)

    If UBound(Words) < 0 Then
        Set pairColl = Nothing
        Exit Sub
    End If

    For word = 0 To UBound(Words)
        nPairs = Len(Words(word)) - 1
        If nPairs > 0 Then
            For pair = 1 To nPairs
                pairColl.Add Mid(Words(word), pair, 2)
            Next pair
        End If
    Next word

End Sub

Private Function SimilarityMetric(sPairs1 As Collection, sPairs2 As Collection) As Variant
'Helper function to calculate similarity metric given two collections of letter pairs.
'This function is designed to allow the pair collections to be set up separately as needed.
'NOTE: sPairs2 collection will be altered as pairs are removed; copy the collection
'if this is not the desired behavior.
'Also assumes that collections will be deallocated somewhere else

    Dim Intersect As Double
    Dim Union As Double
    Dim i, j As Long

    If sPairs1.Count = 0 Or sPairs2.Count = 0 Then
        SimilarityMetric = CVErr(xlErrNA)
        Exit Function
    End If

    Union = sPairs1.Count + sPairs2.Count
    Intersect = 0

    For i = 1 To sPairs1.Count
        For j = 1 To sPairs2.Count
            If StrComp(sPairs1(i), sPairs2(j)) = 0 Then
                Intersect = Intersect + 1
                sPairs2.Remove j
                Exit For
            End If
        Next j
    Next i

    SimilarityMetric = (2 * Intersect) / Union

End Function
