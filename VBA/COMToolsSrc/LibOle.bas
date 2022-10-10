Attribute VB_Name = "LibOle"
Attribute VB_Description = "Credit: Ion Crisitian Buse https://gist.github.com/cristianbuse/b651a3cd740e27a78ea90bca9f7af4d1#file-libole-bas"
'@IgnoreModule ConstantNotUsed
'@Folder("OLE")
'@ModuleDescription("Credit: Ion Crisitian Buse https://gist.github.com/cristianbuse/b651a3cd740e27a78ea90bca9f7af4d1#file-libole-bas")
Option Explicit

#If Mac = False Then
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef pclsid As Any) As Long
    Private Declare PtrSafe Function ProgIDFromCLSID Lib "ole32.dll" (ByRef clsID As Any, ByRef lplpszProgID As LongPtr) As Long
    '@EntryPoint
    Private Declare PtrSafe Function StringFromCLSID Lib "ole32.dll" (ByRef rclsid As Any, ByRef lplpsz As LongPtr) As Long
    Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long
    Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (Optional ByVal pv As LongPtr)
#End If


'OLE Automation Protocol GUIDs
Public Const IID_IRecordInfo As String = "{0000002F-0000-0000-C000-000000000046}"
Public Const IID_IDispatch As String = "{00020400-0000-0000-C000-000000000046}"
Public Const IID_ITypeComp As String = "{00020403-0000-0000-C000-000000000046}"
Public Const IID_ITypeInfo As String = "{00020401-0000-0000-C000-000000000046}"
Public Const IID_ITypeInfo2 As String = "{00020412-0000-0000-C000-000000000046}"
Public Const IID_ITypeLib As String = "{00020402-0000-0000-C000-000000000046}"
Public Const IID_ITypeLib2 As String = "{00020411-0000-0000-C000-000000000046}"
Public Const IID_IUnknown As String = "{00000000-0000-0000-C000-000000000046}"
Public Const IID_IEnumVARIANT As String = "{00020404-0000-0000-C000-000000000046}"
Public Const IID_NULL As String = "{00000000-0000-0000-0000-000000000000}"

'*******************************************************************************
'Converts a string to a GUID struct
'Note that 'CLSIDFromString' win API is only slightly faster (<10%) compared
'   to the pure VB approach (used for MAc only) but it has the advantage of
'   raising other types of errors (like class is not in registry)
'*******************************************************************************
#If Mac Then
Public Function GUIDFromString(ByVal sGUID As String) As GUIDt
    Const methodName As String = "GUIDFromString"
    Const hexPrefix As String = "&H"
    Static pattern As String
    '
    If LenB(pattern) = 0 Then pattern = Replace(IID_NULL, "0", "[0-9A-F]")
    If Not sGUID Like pattern Then Err.Raise 5, methodName, "Invalid string"
    '
    Dim parts() As String: parts = Split(Mid$(sGUID, 2, Len(sGUID) - 2), "-")
    Dim i As Long
    '
    With GUIDFromString
        .Data1 = CLng(hexPrefix & parts(0))
        .Data2 = CInt(hexPrefix & parts(1))
        .Data3 = CInt(hexPrefix & parts(2))
        For i = 0 To 1
            .Data4(i) = CByte(hexPrefix & Mid$(parts(3), i * 2 + 1, 2))
        Next i
        For i = 2 To 7
            .Data4(i) = CByte(hexPrefix & Mid$(parts(4), (i - 1) * 2 - 1, 2))
        Next i
    End With
End Function
#Else
'https://docs.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-clsidfromstring
'@EntryPoint
'@Ignore NonReturningFunction
Public Function GUIDFromString(ByVal sGUID As String) As GUIDt
    Const methodName As String = "GUIDFromString"
    Dim hresult As Long: hresult = CLSIDFromString(StrPtr(sGUID), GUIDFromString)
    If hresult <> S_OK Then Err.Raise hresult, methodName, "Invalid string"
End Function
#End If

'*******************************************************************************
'Converts a GUID struct to a string
'Note that this approach is 4 times faster than running a combination of the
'   following 3 Windows APIs: StringFromCLSID, SysReAllocString, CoTaskMemFree
'*******************************************************************************
'@EntryPoint
Public Function GUIDToString(ByRef gID As GUIDt) As String
    Dim parts(0 To 4) As String
    '
    With gID
        parts(0) = AlignHex(Hex$(.Data1), 8)
        parts(1) = AlignHex(Hex$(.Data2), 4)
        parts(2) = AlignHex(Hex$(.Data3), 4)
        parts(3) = AlignHex(Hex$(.Data4(0) * 256& + .Data4(1)), 4)
        parts(4) = AlignHex(Hex$(.Data4(2) * 65536 + .Data4(3) * 256& + .Data4(4)) _
                          & Hex$(.Data4(5) * 65536 + .Data4(6) * 256& + .Data4(7)), 12)
    End With
    GUIDToString = "{" & Join(parts, "-") & "}"
End Function
Private Function AlignHex(ByRef h As String, ByVal charsCount As Long) As String
    Const maxHex As String = "0000000000000000" '16 chars (LongLong max chars)
    If Len(h) < charsCount Then
        AlignHex = Right$(maxHex & h, charsCount)
    Else
        AlignHex = h
    End If
End Function

'*******************************************************************************
'Converts a CLSID string to a progid string. Windows only
'Returns an empty string if not successful
'*******************************************************************************
#If Mac = False Then
'@EntryPoint
'@Ignore NonReturningFunction
Public Function GetProgIDFromCLSID(ByRef cID As GUIDt) As String
    '@Ignore VariableNotAssigned
    Dim resPtr As LongPtr

    If ProgIDFromCLSID(cID, resPtr) = S_OK Then
        SysReAllocString VarPtr(GetProgIDFromCLSID), resPtr
        CoTaskMemFree resPtr
    End If
End Function
#End If
