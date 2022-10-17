VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GeneralNumericSFI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter.StringFormatIdentifiers")
'@PredeclaredID
Option Explicit

Implements IStringFormatIdentifier

Private Function IStringFormatIdentifier_GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    IStringFormatIdentifier_GetFormattedValue = GetFormattedValue(value, specifier)
End Function

Private Function IStringFormatIdentifier_IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
    IStringFormatIdentifier_IsIdentifierMatch = UCase$(specifier.identifier) = "G"
End Function

Public Function GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
    Dim result As String
    Dim exponentialNotation As String
    
    Dim formatSpecifier As StringFormatSpecifier
    Set formatSpecifier = New StringFormatSpecifier
    formatSpecifier.Alignment = specifier.Alignment
    formatSpecifier.CustomSpecifier = specifier.CustomSpecifier
    
    Dim stringFormatIdentifier As IStringFormatIdentifier

    If StringMatchesAny(TypeName(value), "Integer", "Long") Then
        formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "D")
        Set stringFormatIdentifier = DecimalSFI
        result = stringFormatIdentifier.GetFormattedValue(value, formatSpecifier)

    ElseIf TypeName(value) = "Double" Then
        formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "E")
        Set stringFormatIdentifier = ExponentialSFI
        exponentialNotation = stringFormatIdentifier.GetFormattedValue(value, formatSpecifier)
        
        'Check the power returned by this call, since General
        'only uses exponential notation in certain circumstances
        Dim exponent As Long
        exponent = ExponentialSFI.ParseExponent(exponentialNotation)
        
        If exponent > -5 And Abs(exponent) < specifier.Precision Then
            formatSpecifier.identifier = CopyCapitalisation(specifier.identifier, "F")
            Set stringFormatIdentifier = FixedPointSFI
            result = stringFormatIdentifier.GetFormattedValue(value, formatSpecifier)

        Else
            result = exponentialNotation

        End If

    End If
    
    GetFormattedValue = result
    
End Function

