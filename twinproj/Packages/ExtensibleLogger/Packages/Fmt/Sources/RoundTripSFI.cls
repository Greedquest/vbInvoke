VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RoundTripSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    'ROUND-TRIP format (a string that can round-trip to an identical number)
    'example: ?StringFormat("{0:R}", 0.0000000001141596325677345362656)
    '         ...returns "0.000000000114159632567735"

    'convert value to a Double (chop off overflow digits):
    IStringFormatIdentifier_GetFormattedValue = CDbl(value)
    
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "R"
End Function

