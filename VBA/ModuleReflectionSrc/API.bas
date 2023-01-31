Attribute VB_Name = "API"
'@Folder("TypeInfoInvoker")
Option Explicit

Public Function GetFancyAccessor(Optional ByVal moduleName As String = "ExampleModule", Optional ByVal projectName As Variant) As Object
    Dim project As String
    project = IIf(IsMissing(projectName), Application.VBE.ActiveVBProject.name, projectName)

    Dim moduleTypeInfo As TypeInfo
    Dim accessor As Object
    Dim pITypeLib As LongPtr
    Set accessor = StdModuleAccessor(moduleName, project, moduleTypeInfo, pITypeLib)

    'not sure why but not the same as moduleTypeInfo.ITypeInfo - different objects
    Dim moduleITypeInfo As IUnknown
    Set moduleITypeInfo = getITypeInfo(moduleName, pITypeLib)

    'calling ITypeInfo::GetIDsOfNames, DispGetIDsOfNames etc. does not work
    Set GetFancyAccessor = tryMakeFancyAccessor(accessor, moduleITypeInfo).ExtendedModuleAccessor

End Function

'The IModuleInfo interface gives simplified access to the accessor IDispatch interface
Private Function tryMakeFancyAccessor(ByVal baseAccessor As IUnknown, ByVal ITypeInfo As IUnknown) As IModuleInfo
    Dim result As SwapClass
    Set result = New SwapClass
    Set result.accessor = baseAccessor
    Set result.ITypeInfo = ITypeInfo
    Set tryMakeFancyAccessor = result
End Function

Private Function getITypeInfo(ByVal moduleName As String, ByVal pITypeLib As LongPtr) As IUnknown
    'HRESULT FindName(
    '  [in, out] LPOLESTR  szNameBuf,
    '  [in]      ULONG     lHashVal,
    '  [out]     ITypeInfo **ppTInfo,
    '  [out]     MEMBERID  *rgMemId,
    '  [in, out] USHORT * pcFound
    ');
    Dim hresult As hResultCode
    Dim pModuleITypeInfoArray(1 To 1) As LongPtr
    Dim memberIDArray(1 To 1) As Long
    '@Ignore IntegerDataType
    Dim pcFound As Integer 'number of matches
    pcFound = 1
    'call ITypeLib::FindName to get the module specific type info
    hresult = COMTools.CallFunction( _
        pITypeLib, ITypeLibVTableOffset(ITypeLibVTable.FindName), _
        CR_HRESULT, CC_STDCALL, _
        StrPtr(moduleName), _
        0&, _
        VarPtr(pModuleITypeInfoArray(1)), _
        VarPtr(memberIDArray(1)), _
        VarPtr(pcFound))

    If hresult <> S_OK Then Err.Raise hresult
    Set getITypeInfo = ObjectFromObjPtr(pModuleITypeInfoArray(1))
End Function



