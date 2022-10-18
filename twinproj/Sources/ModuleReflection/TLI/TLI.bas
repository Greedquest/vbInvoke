Attribute VB_Name = "TLI"
'@Folder("TLI")
Option Explicit

Public Const NULL_PTR As LongPtr = 0
Public Enum KnownMemberIDs
    MEMBERID_NIL = -1
End Enum

Public Function TypeLibInfoFromITypeLib(ByVal ITypeLib As ITypeLib) As TypeLibInfo
    Dim result As New TypeLibInfo
    Set result.ITypeLib = ITypeLib
    Set TypeLibInfoFromITypeLib = result
End Function
