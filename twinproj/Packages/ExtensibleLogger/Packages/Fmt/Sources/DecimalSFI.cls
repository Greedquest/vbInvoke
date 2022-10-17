VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    Const baseFormat As String = "0"
    IStringFormatIdentifier_GetFormattedValue = Format$(value, Right$(String$(specifier.Precision, "0") & baseFormat, IIf(specifier.Precision = 0, Len(baseFormat), specifier.Precision)))
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "D"
End Function

