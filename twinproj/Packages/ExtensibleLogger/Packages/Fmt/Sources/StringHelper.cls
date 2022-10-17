VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter")
'@PredeclaredId
Option Explicit

Private Const ERR_FORMAT_EXCEPTION As Long = vbObjectError Or 9001
Private Const ERR_SOURCE As String = "StringHelper"
Private Const ERR_MSG_INVALID_FORMAT_STRING As String = "Invalid format string."
Private Const ERR_MSG_FORMAT_EXCEPTION As String = "The number indicating an argument to format is less than zero, or greater than or equal to the length of the args array."

Private Type tString
    PaddingCharacter As String '* 1
    EscapeSequences As Collection
    NumericSpecifiers As Collection
    DateTimeSpecifiers As Collection
End Type

Private Const DEFAULT_PADDING_CHAR As String * 1 = " "

Private this As tString

Private Sub Class_Initialize()

    If this.PaddingCharacter = vbNullString Then this.PaddingCharacter = DEFAULT_PADDING_CHAR

    InitEscapeSequences
    InitNumericSpecifiers
    InitDateTimeSpecifiers

End Sub

Private Sub InitEscapeSequences()

    Set this.EscapeSequences = New Collection
    this.EscapeSequences.Add EscapeSequence.Create("\n", vbNewLine)
    this.EscapeSequences.Add EscapeSequence.Create("\q", Chr$(34))
    this.EscapeSequences.Add EscapeSequence.Create("\t", vbTab)
    this.EscapeSequences.Add EscapeSequence.Create("\a", Chr$(7))
    this.EscapeSequences.Add EscapeSequence.Create("\b", Chr$(8))
    this.EscapeSequences.Add EscapeSequence.Create("\v", Chr$(13))
    this.EscapeSequences.Add EscapeSequence.Create("\f", Chr$(14))
    this.EscapeSequences.Add EscapeSequence.Create("\r", Chr$(15))
    this.EscapeSequences.Add EscapeSequence.Create("\\x(\w{2})", 0, True, Hexadecimal)
    this.EscapeSequences.Add EscapeSequence.Create("\\(\d{3})", 0, True, Octal)

End Sub

Private Sub InitNumericSpecifiers()

    Dim myDummy As LongPtr
    myDummy = ObjPtr(CurrencySFI)
    myDummy = ObjPtr(DecimalSFI)
    myDummy = ObjPtr(GeneralNumericSFI)
    myDummy = ObjPtr(PercentSFI)
    myDummy = ObjPtr(FixedPointSFI)
    myDummy = ObjPtr(ExponentialSFI)
    myDummy = ObjPtr(HexSFI)
    myDummy = ObjPtr(RoundTripSFI)
    myDummy = ObjPtr(NumericPaddingSFI)
    
    Set this.NumericSpecifiers = New Collection
    
    With this.NumericSpecifiers
    
        .Add CurrencySFI
        .Add DecimalSFI
        .Add GeneralNumericSFI
        .Add PercentSFI
        .Add FixedPointSFI
        .Add ExponentialSFI
        .Add HexSFI
        .Add RoundTripSFI
        .Add NumericPaddingSFI
    
End With

End Sub

Private Sub InitDateTimeSpecifiers()

    Dim mydummy As LongPtr
    mydummy = ObjPtr(GenericDateTimeSFI)
    Set this.DateTimeSpecifiers = New Collection
    this.DateTimeSpecifiers.Add GenericDateTimeSFI
    '    this.DateTimeSpecifiers.Add New CustomDFI
    '    this.DateTimeSpecifiers.Add New FullDateLongStringFormatSpecifier
    '    this.DateTimeSpecifiers.Add New FullDateShortSFI
    '    this.DateTimeSpecifiers.Add New GeneralLongDateTimeSFI
    '    this.DateTimeSpecifiers.Add New GeneralShortDateTimeSFI
    '    this.DateTimeSpecifiers.Add New LongDFI
    '    this.DateTimeSpecifiers.Add New LongTimeSFI
    '    this.DateTimeSpecifiers.Add New ShortDFI
    '    this.DateTimeSpecifiers.Add New SortableDateTimeSFI

End Sub

Private Function GetFormatSpecifiers(ByVal format_string As String, valuesCount As Long) As Collection
    'executes a regular expression against format_string to extract all placeholders into a MatchCollection

    Dim regex As RegExp
    Set regex = New RegExp
    Dim matches As MatchCollection
    Dim thisMatch As Match

    Dim result As Collection
    Set result = New Collection
    Dim specifier As StringFormatSpecifier

    Dim csvIndices As String
    csvIndices = ","
    
    Dim uniqueCount As Long
    Dim largestIndex As Long

    regex.Pattern = "\{(\w+)(\,\-?\d+)?(\:[^}]+)?\}"

    ' literal {
    ' [1] numbered captured group, any number of repetitions (Index)
    '    alphanumeric, one or more repetitions
    ' [2] numbered captured group, zero or one repetitions (AlignmentSpecifier)
    '    literal ,
    '    literal -, zero or one repetitions
    '    any digit, one or more repetitions
    ' [3] numbered captured group, zero or one repetitions (FormatSpecifier)
    '    literal :
    '    any character except '}', one or more repetitions
    ' literal }

    regex.IgnoreCase = True
    regex.Global = True

    Set matches = regex.Execute(format_string)
    For Each thisMatch In matches                'loop through all

        'Extract index portion of ReGex match
        Set specifier = New StringFormatSpecifier
        specifier.Index = CInt(thisMatch.SubMatches(0)) 'TODO why just integer?

        If Not StringContains(csvIndices, "," & specifier.Index & ",") Then
            uniqueCount = uniqueCount + 1
            csvIndices = csvIndices & specifier.Index & ","
        End If
        If specifier.Index > largestIndex Then largestIndex = specifier.Index

        If Not thisMatch.SubMatches(1) = vbEmpty Then specifier.Alignment = Mid$(thisMatch.SubMatches(1), 2)
        If Not thisMatch.SubMatches(2) = vbEmpty Then
            specifier.identifier = Mid$(thisMatch.SubMatches(2), 2, 1)
            specifier.CustomSpecifier = Mid$(thisMatch.SubMatches(2), 3)
        End If

        result.Add specifier
    Next

    If matches.Count > 0 And (uniqueCount <> valuesCount Or largestIndex >= uniqueCount Or valuesCount = 0) Then Err.Raise ERR_FORMAT_EXCEPTION, ERR_SOURCE, ERR_MSG_FORMAT_EXCEPTION

    Set GetFormatSpecifiers = result

End Function

Public Function StringFormat(format_string As String, values() As Variant) As String

    Dim result As String
    result = format_string

    Dim specifiers As Collection
    Dim specifier As StringFormatSpecifier
    Set specifiers = GetFormatSpecifiers(result, UBound(values) + 1)

    Dim useLiteral As Boolean
    'when format_string starts with "@", escapes are not replaced
    '(string is treated as a literal string with placeholders)
    useLiteral = StringStartsWith("@", result)
    
    If useLiteral Then
        'remove the "@" literal specifier from the result string
        result = Right$(result, Len(result) - 1)
    
    ElseIf StringContains(result, "\\") Then
        'replace escaped backslashes with 'ESC' character [Chr$(27)]
        'to optimize escape sequences evaluation:
        result = Replace(result, "\\", Chr$(27))
        
    End If

    Dim formattedValue As String
    Dim alignmentPadding As Long
    Dim identifier As IStringFormatIdentifier
    Dim identifierFound As Boolean

    For Each specifier In specifiers

        formattedValue = values(specifier.Index)
        identifierFound = (specifier.identifier = vbNullString)

        If IsNumeric(values(specifier.Index)) Then
            For Each identifier In this.NumericSpecifiers
                If identifier.IsIdentifierMatch(specifier) Then
                    identifierFound = True
                    formattedValue = identifier.GetFormattedValue(values(specifier.Index), specifier)

                End If
                
            Next

        ElseIf TypeName(values(specifier.Index)) = "Date" Then
            For Each identifier In this.DateTimeSpecifiers
                If identifier.IsIdentifierMatch(specifier) Then
                    identifierFound = True
                    formattedValue = identifier.GetFormattedValue(values(specifier.Index), specifier)

                End If
                
            Next
            
        End If

        If Not identifierFound Then Err.Raise ERR_FORMAT_EXCEPTION, ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING

        alignmentPadding = Abs(specifier.Alignment)
        If specifier.Alignment < 0 Then
            'negative: left-justified alignment
            If alignmentPadding - Len(formattedValue) > 0 Then
                formattedValue = formattedValue & String$(alignmentPadding - Len(formattedValue), this.PaddingCharacter)
            End If

        ElseIf specifier.Alignment > 0 Then
            'positive: right-justified alignment
            If alignmentPadding - Len(formattedValue) > 0 Then
                formattedValue = String$(alignmentPadding - Len(formattedValue), this.PaddingCharacter) & formattedValue
            End If

        End If

        'replace all occurrences of placeholder {i} with their formatted values:
        result = Replace(result, specifier.ToString, formattedValue)

    Next

    Dim escape As EscapeSequence
    If Not useLiteral And StringContains(result, "\") Then
        For Each escape In this.EscapeSequences
            escape.Execute result
        Next
    End If

    If Not useLiteral And StringContains(result, Chr$(27)) Then result = Replace(result, Chr$(27), "\")
    StringFormat = result

End Function

