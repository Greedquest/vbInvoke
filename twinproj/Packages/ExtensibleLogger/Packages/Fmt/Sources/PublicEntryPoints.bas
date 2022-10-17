Attribute VB_Name = "PublicEntryPoints"
'@Folder("CSharpishStringFormatter")
Option Explicit

Public Function StringFormat(format_string As String, ParamArray values() as variant) As String
    Dim valuesArray() As Variant
    valuesArray = values
    StringFormat = StringHelper.StringFormat(format_string, valuesArray)
End Function

