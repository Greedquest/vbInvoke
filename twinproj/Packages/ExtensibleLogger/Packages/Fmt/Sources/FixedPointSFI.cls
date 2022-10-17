VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FixedPointSFI"
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
    If specifier.Precision > 0 Then
        IStringFormatIdentifier_GetFormattedValue = Format$(value, (baseFormat & ".") & String$(specifier.Precision, Chr$(48)))
    Else
        'no precision specified - default to 2 decimals:
        IStringFormatIdentifier_GetFormattedValue = Format$(value, "0.00")
    End If
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "F"
End Function

