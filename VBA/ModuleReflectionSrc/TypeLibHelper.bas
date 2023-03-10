Attribute VB_Name = "TypeLibHelper"
'@Folder "TLI"
Option Explicit

Public Function getITypeInfoByIndex(ByVal ITypeLib As IUnknown, ByVal index As Long) As IUnknown

'4      HRESULT  GetTypeInfo(
'            /* [in] */ UINT index,
'            /* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo) = 0;
    Dim hresult As hResultCode
    Dim pITypeInfo As LongPtr
    hresult = COMTools.CallCOMObjectVTableEntry(ITypeLib, ITypeLibVTableOffset(ITypeLibVTable.getTypeInfo), CR_HRESULT, index, VarPtr(pITypeInfo))
    If hresult <> S_OK Then Err.Raise hresult
    Set getITypeInfoByIndex = COMTools.ObjectFromObjPtr(pITypeInfo)
End Function

Public Function getTypeInfoCount(ByVal ITypeLib As IUnknown) As Long
'3      UINT     GetTypeInfoCount( void) = 0;
'TODO: assert not nothing
    getTypeInfoCount = COMTools.CallCOMObjectVTableEntry(ITypeLib, ITypeLibVTableOffset(ITypeLibVTable.getTypeInfoCount), CR_LONG)
End Function


Public Function getProjName(ByVal ITypeLib As IUnknown) As String
    getProjName = getDocumentation(ITypeLib, KnownMemberIDs.MEMBERID_NIL)
End Function
Private Function getDocumentation(ByVal ITypeLib As IUnknown, ByVal memid As dispid) As String
'        virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDocumentation(
'            /* [in] */ INT index,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrName,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrDocString,
'            /* [out] */ DWORD *pdwHelpContext,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrHelpFile) = 0;
    Dim hresult As hResultCode
    hresult = COMTools.CallCOMObjectVTableEntry(ITypeLib, ITypeLibVTableOffset(ITypeLibVTable.getDocumentation), CR_HRESULT, memid, VarPtr(getDocumentation), NULL_PTR, NULL_PTR, NULL_PTR)
    If hresult <> S_OK Then Err.Raise hresult
End Function
