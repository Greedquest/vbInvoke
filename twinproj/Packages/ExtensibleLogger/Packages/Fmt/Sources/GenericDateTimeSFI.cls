VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GenericDateTimeSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
If UCase$(specifier.identifier) = "C" Then
    IStringFormatIdentifier_GetFormattedValue = Format$(value, specifier.CustomSpecifier)
Else
    IStringFormatIdentifier_GetFormattedValue = Format$(value, formatStringFromIdentifierDict(specifier.identifier))
End If
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    If UCase$(specifier.identifier) = "C" Then
       IStringFormatIdentifier_IsIdentifierMatch = True
    ElseIf formatStringFromIdentifierDict.Exists(specifier.identifier) Then
        IStringFormatIdentifier_IsIdentifierMatch = True
    End If
End Function


Private Static Property Get formatStringFromIdentifierDict() As Dictionary
    Dim result As Dictionary
    If result Is Nothing Then
        Set result = New Dictionary
        With result
            .Add "d", "ddddd"                        'SHORT DATE format
            .Add "D", "dddddd"                         'LONG DATE format
            .Add "f", "dddddd h:mm AM/PM"                        'FULL DATE format (short)
            .Add "F", "dddddd ttttt"                        'FULL DATE format (long)
            .Add "g", "ddddd hh:mm AM/PM"
            .Add "G", "ddddd ttttt"
            .Add "s", "yyyy-mm-ddThh:mm:ss"                        'SORTABLE DATETIME format
            .Add "t", "hh:mm AM/PM"                        'SHORT TIME format
            .Add "T", "ttttt"                         'LONG TIME format
        End With
    End If
    Set formatStringFromIdentifierDict = result
End Property
