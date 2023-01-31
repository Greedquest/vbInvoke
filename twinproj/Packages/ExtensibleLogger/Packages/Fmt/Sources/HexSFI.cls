VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HexSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    Dim hexValue As String
    hexValue = Hex$(value)

    'add left padding and trim to size
    If specifier.Precision > 0 Then
        hexValue = Right$(String$(specifier.Precision, "0") & hexValue, specifier.Precision)
    End If

    IStringFormatIdentifier_GetFormattedValue = CopyCapitalisation(specifier.identifier, "&h" & hexValue)

End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "X"
End Function

