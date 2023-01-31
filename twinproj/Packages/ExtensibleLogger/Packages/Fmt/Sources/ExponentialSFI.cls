VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExponentialSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    'NOTE - formerly check LenB(formatGroup) > 2 when precision not specified to see if invalid specifier.identifier
    'But that should be checked earlier
    If specifier.Precision > 0 Then
        IStringFormatIdentifier_GetFormattedValue = Format$(value, "0." & String$(specifier.Precision - 1, Chr$(48)) & "#" & specifier.identifier & "-#")

    Else
        IStringFormatIdentifier_GetFormattedValue = Format$(value, "0.00000#" & specifier.identifier & "-#") 'defaults to 6 decimals

    End If
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "E"
End Function

'Should be static method really
Public Function ParseExponent(ByVal formattedValue As String) As Long
    ParseExponent = Mid$(formattedValue, InStr(1, formattedValue, "E", vbTextCompare) + 1)
End Function

