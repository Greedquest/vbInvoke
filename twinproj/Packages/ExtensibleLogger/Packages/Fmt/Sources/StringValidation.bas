Attribute VB_Name = "StringValidation"
'@Folder("CSharpishStringFormatter.Common")
Option Explicit
Option Private Module

Public Function StringContains(ByVal haystack As String, ByVal needle As String, Optional ByVal caseSensitive As Boolean = False) As Boolean

    Dim compareMethod As Long

    If caseSensitive Then
        compareMethod = vbBinaryCompare
    Else
        compareMethod = vbTextCompare
    End If
    'Have you thought about Null?
    StringContains = (InStr(1, haystack, needle, compareMethod) <> 0)

End Function

Public Function StringContainsAny(ByVal haystack As String, ByVal caseSensitive As Boolean, ParamArray needles() As Variant) As Boolean

    Dim i As Long

    For i = LBound(needles) To UBound(needles)
        If StringContains(haystack, needles(i), caseSensitive) Then
            StringContainsAny = True
            Exit Function
        End If
    Next

    StringContainsAny = False                    'Not really necessary, default is False..

End Function

Public Function StringMatchesAny(ByVal searchString As String, ParamArray possibleMatches() As Variant) As Boolean

    'String-typed local copies of passed parameter values:
    Dim i As Long

    StringMatchesAny = True
    For i = LBound(possibleMatches) To UBound(possibleMatches)
        If searchString = CStr(possibleMatches(i)) Then Exit Function
    Next

    StringMatchesAny = False

End Function

Public Function StringStartsWith(ByVal startingSequence As String, ByVal source As String) As Boolean
    StringStartsWith = Left$(source, Len(startingSequence)) = startingSequence
End Function

Public Function CopyCapitalisation(ByVal source As String, ByVal applyTo As String) As String
    If UCase$(source) = source Then
        CopyCapitalisation = UCase$(applyTo)
    Else
        CopyCapitalisation = LCase$(applyTo)
    End If
End Function

