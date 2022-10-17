Attribute VB_Name = "FormerEntryPoint"
'@Folder("FormerCSharpishStringFormatter")
Option Explicit

Public Function FormerStringFormat(format_string As String, ParamArray values() as variant) As String
    'VB6 implementation of .net String.Format(), slightly customized.

    Dim return_value As String
    Dim values_count As Long
    Dim PADDING_CHAR As String

    'some error-handling constants:
    Const ERR_FORMAT_EXCEPTION As Long = vbObjectError Or 9001
    Const ERR_ARGUMENT_NULL_EXCEPTION As Long = vbObjectError Or 9002
    Const ERR_ARGUMENT_EXCEPTION As Long = vbObjectError Or 9003
    Const ERR_SOURCE As String = "FormerStringFormat"
    Const ERR_MSG_INVALID_FORMAT_STRING As String = "Invalid format string."
    Const ERR_MSG_FORMAT_EXCEPTION As String = "The number indicating an argument to format is less than zero, or greater than or equal to the length of the args array."
    Const ERR_MSG_NUMBER_ARGUMENT_EXCEPTION As String = "Invalid number argument."

    'use SPACE as default padding character
    If PADDING_CHAR = vbNullString Then PADDING_CHAR = Chr$(32)

    'figure out number of passed values:
    values_count = UBound(values) + 1

    Dim regex As RegExp
    Dim matches As MatchCollection
    Dim thisMatch As Match
    Dim thisString As String
    Dim thisFormat As String

    Dim useLiteral As Boolean                    'when format_string starts with "@", escapes are not replaced (string is treated as a literal string with placeholders)
    Dim escapeHex As Boolean                     'indicates whether HEX specifier "0x" is to be escaped or not

    'validate string_format:
    Set regex = New RegExp
    regex.Pattern = "{({{)*(\w+)(,-?\d+)?(:[^}]+)?}(}})*"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(format_string)

    'determine if values_count matches number of unique regex matches:
    Dim uniqueCount As Long
    Dim tmpCSV As String
    For Each thisMatch In matches
        If Not StringContains(tmpCSV, thisMatch.SubMatches(1)) Then
            uniqueCount = uniqueCount + 1
            tmpCSV = tmpCSV & thisMatch.SubMatches(1) & ","
        End If
    Next

    'unique indices count must match values_count:
    If matches.Count > 0 And uniqueCount <> values_count Then
        Err.Raise ERR_FORMAT_EXCEPTION, _
                  ERR_SOURCE, ERR_MSG_FORMAT_EXCEPTION
    End If

    useLiteral = StringStartsWith("@", format_string)
    If useLiteral Then format_string = Right$(format_string, Len(format_string) - 1) 'remove the "@" literal specifier

    If Not useLiteral And StringContains(format_string, "\\") Then _
                                                              format_string = Replace(format_string, "\\", Chr$(27))

    If matches.Count = 0 And format_string <> vbNullString And UBound(values) = -1 Then
        'only format_string was specified: skip to checking escape sequences:
        return_value = format_string
        GoTo checkEscapes
    ElseIf UBound(values) = -1 And matches.Count > 0 Then
        Err.Raise ERR_ARGUMENT_NULL_EXCEPTION, _
                  ERR_SOURCE, ERR_MSG_FORMAT_EXCEPTION
    End If

    return_value = format_string

    'dissect format_string:

    Dim i As Long, v As String, p As String   'i: iterator; v: value; p: placeholder
    Dim alignmentGroup As String, alignmentSpecifier As String
    Dim formattedValue As String, alignmentPadding As Long

    'iterate regex matches (each match is a placeholder):
    For i = 0 To matches.Count - 1

        'get the placeholder specified index:
        Set thisMatch = matches(i)
        p = thisMatch.SubMatches(1)

        'if specified index (0-based) > uniqueCount (1-based), something's wrong:
        If p > uniqueCount - 1 Then
            Err.Raise ERR_FORMAT_EXCEPTION, _
                      ERR_SOURCE, ERR_MSG_FORMAT_EXCEPTION
        End If
        v = values(p)

        'get the alignment specifier if it is specified:
        alignmentGroup = thisMatch.SubMatches(2)
        If alignmentGroup <> vbNullString Then _
           alignmentSpecifier = Right$(alignmentGroup, LenB(alignmentGroup) / 2 - 1)


        'get the format specifier if it is specified:
        thisString = thisMatch.Value
        If StringContains(thisString, ":") Then

            Dim formatGroup As String, precisionSpecifier As Long
            Dim formatSpecifier As String, precisionString As String

            'get the string between ":" and "}":
            formatGroup = Mid$(thisString, InStr(1, thisString, ":") + 1, (LenB(thisString) / 2) - 2)
            formatGroup = Left$(formatGroup, LenB(formatGroup) / 2 - 1)

            precisionString = Right$(formatGroup, LenB(formatGroup) / 2 - 1)
            formatSpecifier = Mid$(thisString, InStr(1, thisString, ":") + 1, 1)

            'applicable formatting depends on the type of the value (yes, GOTO!!):
            If TypeName(values(p)) = "Date" Then GoTo DateTimeFormatSpecifiers
            If v = vbNullString Then GoTo ApplyFormerStringFormat

NumberFormatSpecifiers:
            If precisionString <> vbNullString And Not IsNumeric(precisionString) Then
                Err.Raise ERR_FORMAT_EXCEPTION, _
                          ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
            End If

            If Not IsNumeric(v) Then
                Err.Raise ERR_ARGUMENT_EXCEPTION, _
                          ERR_SOURCE, ERR_MSG_NUMBER_ARGUMENT_EXCEPTION
            End If

            If precisionString = vbNullString Then precisionString = 0

            Select Case formatSpecifier

                Case "C", "c"                    'CURRENCY format, formats string as currency.
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings
                    '(hard-coded group separator, decimal separator and currency sign).

                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "#,##0.$"

                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = _
                                   Replace$(thisFormat, ".", "." & String$(precisionString, Chr$(48)))
                    Else
                        thisFormat = "#,##0.00$"
                    End If


                Case "D", "d"                    'DECIMAL format, formats string as integer number.
                    'Precision specifier determines number of digits in returned string.


                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    thisFormat = Right$(String$(precisionSpecifier, "0") & thisFormat, _
                                        IIf(precisionSpecifier = 0, Len(thisFormat), precisionSpecifier))


                Case "E", "e"                    'EXPONENTIAL NOTATION format (aka "Scientific Notation")
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).


                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0.00000#" & formatSpecifier & "-#" 'defaults to 6 decimals

                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "0." & String$(precisionSpecifier - 1, Chr$(48)) & "#" & formatSpecifier & "-#"

                    ElseIf LenB(formatGroup) > 2 And precisionSpecifier = 0 Then
                        Err.Raise ERR_FORMAT_EXCEPTION, _
                                  ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                    End If


                Case "F", "f"                    'FIXED-POINT format
                    'Precision specifier determines number of decimals in returned string.
                    'This implementation ignores regional settings'
                    '(hard-coded decimal separator).

                    precisionSpecifier = CInt(precisionString)
                    thisFormat = "0"
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))
                    Else
                        'no precision specified - default to 2 decimals:
                        thisFormat = "0.00"
                    End If


                Case "G", "g"                    'GENERAL format (recursive)
                    'returns the shortest of either FIXED-POINT or SCIENTIFIC formats in case of a Double.
                    'returns DECIMAL format in case of a Integer or Long.

                    Dim eNotation As String, ePower As Long, specifier As String
                    precisionSpecifier = IIf(CInt(precisionString) > 0, CInt(precisionString), _
                                             IIf(StringContains(v, "."), Len(v) - InStr(1, v, "."), 0))

                    'track character case of formatSpecifier:
                    specifier = IIf(formatSpecifier = "G", "D", "d")

                    If TypeName(values(p)) = "Integer" Or TypeName(values(p)) = "Long" Then
                        'Integer types: use {0:D} (recursive call):
                        formattedValue = FormerStringFormat("{0:" & specifier & "}", values(p))

                    ElseIf TypeName(values(p)) = "Double" Then
                        'Non-integer types: use {0:E}
                        specifier = IIf(formatSpecifier = "G", "E", "e")

                        'evaluate the exponential notation value (recursive call):
                        eNotation = FormerStringFormat("{0:" & specifier & "}", v)

                        'get the power of eNotation:
                        '@Ignore UntypedFunctionUsage:no sense zooming to and from Variant
                        ePower = Mid(eNotation, InStr(1, eNotation, "e", vbTextCompare) + 1)

                        If ePower > -5 And Abs(ePower) < precisionSpecifier Then
                            'use {0:F} when ePower > -5 and abs(ePower) < precisionSpecifier:
                            'evaluate the floating-point value (recursive call):
                            specifier = IIf(formatSpecifier = "G", "F", "f")
                            formattedValue = FormerStringFormat("{0:" & specifier & _
                                                                IIf(precisionSpecifier <> 0, precisionString, vbNullString) & "}", values(p))
                        Else
                            'fallback to {0:E} if previous rule didn't apply:
                            formattedValue = eNotation
                        End If

                    End If

                    GoTo AlignFormattedValue     'Skip the "ApplyFormerStringFormat" step, it's applied already.


                Case "N", "n"                    'NUMERIC format, formats string as an integer or decimal number.
                    'Precision specifier determines number of decimal digits.
                    'This implementation ignores regional settings'
                    '(hard-coded group and decimal separators).

                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))

                    Else                         'only the "D" is specified
                        thisFormat = "#,##0"
                    End If


                Case "P", "p"                    'PERCENT format. Formats string as a percentage.
                    'Value is multiplied by 100 and displayed with a percent symbol.
                    'Precision specifier determines number of decimal digits.

                    thisFormat = "#,##0%"
                    precisionSpecifier = CInt(precisionString)
                    If LenB(formatGroup) > 2 And precisionSpecifier > 0 Then
                        'if a non-zero precision is specified...
                        thisFormat = "#,##0"
                        thisFormat = (thisFormat & ".") & String$(precisionSpecifier, Chr$(48))

                    Else                         'only the "P" is specified
                        thisFormat = "#,##0"
                    End If

                    'Append the percentage sign to the format string:
                    thisFormat = thisFormat & "%"


                Case "R", "r"                    'ROUND-TRIP format (a string that can round-trip to an identical number)
                    'example: ?FormerStringFormat("{0:R}", 0.0000000001141596325677345362656)
                    '         ...returns "0.000000000114159632567735"

                    'convert value to a Double (chop off overflow digits):
                    v = CDbl(v)


                Case "X", "x"                    'HEX format. Formats a string as a Hexadecimal value.
                    'Precision specifier determines number of total digits.
                    'Returned string is prefixed with "&H" to specify Hex.

                    v = Hex$(v)
                    precisionSpecifier = CInt(precisionString)

                    If LenB(precisionString) > 0 Then 'precision here stands for left padding
                        v = Right$(String$(precisionSpecifier, "0") & v, IIf(precisionSpecifier = 0, Len(v), precisionSpecifier))
                    End If

                    'add C# hex specifier, apply specified casing:
                    '(VB6 hex specifier would cause Format() to reverse the formatting):
                    v = "0x" & IIf(formatSpecifier = "X", UCase$(v), LCase$(v))
                    escapeHex = True

                Case Else
                    If IsNumeric(formatSpecifier) And Val(formatGroup) = 0 Then
                        formatSpecifier = formatGroup
                        v = Format$(v, formatGroup)
                    Else
                        Err.Raise ERR_FORMAT_EXCEPTION, _
                                  ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
                    End If
            End Select

            GoTo ApplyFormerStringFormat


DateTimeFormatSpecifiers:
            Select Case formatSpecifier

                Case "c", "C"                    'CUSTOM date/time format
                    'let VB Format() parse precision specifier as is:
                    thisFormat = precisionString

                Case "d"                         'SHORT DATE format
                    thisFormat = "ddddd"

                Case "D"                         'LONG DATE format
                    thisFormat = "dddddd"

                Case "f"                         'FULL DATE format (short)
                    thisFormat = "dddddd h:mm AM/PM"

                Case "F"                         'FULL DATE format (long)
                    thisFormat = "dddddd ttttt"

                Case "g"
                    thisFormat = "ddddd hh:mm AM/PM"

                Case "G"
                    thisFormat = "ddddd ttttt"

                Case "s"                         'SORTABLE DATETIME format
                    thisFormat = "yyyy-mm-ddThh:mm:ss"

                Case "t"                         'SHORT TIME format
                    thisFormat = "hh:mm AM/PM"

                Case "T"                         'LONG TIME format
                    thisFormat = "ttttt"

                Case Else
                    Err.Raise ERR_FORMAT_EXCEPTION, _
                              ERR_SOURCE, ERR_MSG_INVALID_FORMAT_STRING
            End Select
            GoTo ApplyFormerStringFormat

        End If


ApplyFormerStringFormat:
        'apply computed format string:
        If thisFormat <> vbNullString Then
            formattedValue = Format$(v, thisFormat)
        Else
            formattedValue = v
        End If


AlignFormattedValue:
        'apply specified alignment specifier:
        If alignmentSpecifier <> vbNullString Then

            alignmentPadding = Abs(CInt(alignmentSpecifier))
            If CInt(alignmentSpecifier) < 0 Then
                'negative: left-justified alignment
                If alignmentPadding - Len(formattedValue) > 0 Then _
                                                          formattedValue = formattedValue & _
                                                          String$(alignmentPadding - Len(formattedValue), PADDING_CHAR)
            Else
                'positive: right-justified alignment
                If alignmentPadding - Len(formattedValue) > 0 Then _
                                                          formattedValue = String$(alignmentPadding - Len(formattedValue), PADDING_CHAR) & formattedValue
            End If
        End If

        'Replace C# hex specifier with VB6 hex specifier, only if hex specifier was introduced in this function:
        If (Not useLiteral And escapeHex) And StringContains(formattedValue, "0x") Then formattedValue = Replace$(formattedValue, "0x", "&H")

        'replace all occurrences of placeholder {i} with their formatted values:
        return_value = Replace(return_value, thisString, formattedValue, Count:=1)

        'reset before reiterating:
        thisFormat = vbNullString
    Next


checkEscapes:
    'if there's no more backslashes, don't bother checking for the rest:
    If useLiteral Or Not StringContains(return_value, "\") Then GoTo normalExit
    Dim escape As EscapeSequence
    Set escape = New EscapeSequence
    Dim escapes As Collection
    Set escapes = New Collection
    escapes.Add escape.Create("\n", vbNewLine), "0"
    escapes.Add escape.Create("\q", Chr$(34)), "1"
    escapes.Add escape.Create("\t", vbTab), "2"
    escapes.Add escape.Create("\a", Chr$(7)), "3"
    escapes.Add escape.Create("\b", Chr$(8)), "4"
    escapes.Add escape.Create("\v", Chr$(13)), "5"
    escapes.Add escape.Create("\f", Chr$(14)), "6"
    escapes.Add escape.Create("\r", Chr$(15)), "7"

    For i = 0 To escapes.Count - 1
        Set escape = escapes(CStr(i))
        If StringContains(return_value, escape.EscapeString) Then _
                                                             return_value = Replace(return_value, escape.EscapeString, escape.ReplacementString)

        If Not StringContains(return_value, "\") Then _
                                                 GoTo normalExit
    Next

    'replace "ASCII (oct)" escape sequence
    Set regex = New RegExp
    regex.Pattern = "\\(\d{3})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(format_string)

    Dim char As Long
    If matches.Count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the octal number representing the ASCII code we're after:
            p = "&O" & p                         'prepend octal prefix
            char = CLng(p)
            return_value = Replace(return_value, thisMatch.Value, Chr$(char))
        Next
    End If

    'if there's no more backslashes, don't bother checking for the rest:
    If Not StringContains("\", return_value) Then GoTo normalExit

    'replace "ASCII (hex)" escape sequence
    Set regex = New RegExp
    regex.Pattern = "\\x(\w{2})"
    regex.IgnoreCase = True
    regex.Global = True
    Set matches = regex.Execute(format_string)

    If matches.Count <> 0 Then
        For Each thisMatch In matches
            p = thisMatch.SubMatches(0)
            '"p" contains the hex value representing the ASCII code we're after:
            p = "&H" & p                         'prepend hex prefix
            char = CLng(p)
            return_value = Replace(return_value, thisMatch.Value, Chr$(char))
        Next
    End If

normalExit:
    Set escapes = Nothing
    Set escape = Nothing
    If Not useLiteral And StringContains(return_value, Chr$(27)) Then _
                                                                 return_value = Replace(return_value, Chr$(27), "\")
    FormerStringFormat = return_value
End Function


