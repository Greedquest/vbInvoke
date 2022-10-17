VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IStringFormatIdentifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter")
'@Interface
Option Explicit

'returns the formatted value.
Public Function GetFormattedValue(ByVal value As Variant, ByVal specifier As StringFormatSpecifier) As String
End Function

'compares specified format identifier with implementation-defined one,
'returns true if format is applicable.
Public Function IsIdentifierMatch(ByVal specifier As StringFormatSpecifier) As Boolean
End Function

