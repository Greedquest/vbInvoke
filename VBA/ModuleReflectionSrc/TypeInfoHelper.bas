Attribute VB_Name = "TypeInfoHelper"
Attribute VB_Description = "ITypeInfo parsing/navigation without TLBINF32.dll. We don't want that because (1) It's no longer included in Windows, and (2) It ignores the type info marked as 'private', which we want to see"
'@Folder "TLI"
Option Explicit
Option Private Module
'Created by JAAFAR
'Src: https://www.vbforums.com/showthread.php?846947-RESOLVED-Ideas-Wanted-ITypeInfo-like-Solution&p=5449985&viewfull=1#post5449985
'Modified by wqweto 2020 (clean up)
'Modified by Greedo 2022 (refactor)
'@ModuleDescription("ITypeInfo parsing/navigation without TLBINF32.dll. We don't want that because (1) It's no longer included in Windows, and (2) It ignores the type info marked as 'private', which we want to see")



'@Description("Returns a map of funcName:dispid given a certain ITypeInfo without TLBINF32.dll")
Public Function GetFuncDispidFromTypeInfo(ByVal ITypeInfo As IUnknown) As Scripting.Dictionary
Attribute GetFuncDispidFromTypeInfo.VB_Description = "Returns a map of funcName:dispid given a certain ITypeInfo without TLBINF32.dll"
    Dim attrs As TYPEATTR
    attrs = getAttrs(ITypeInfo)

    Dim result As Scripting.Dictionary
    Set result = New Scripting.Dictionary
    result.CompareMode = TextCompare 'so we can look names up in a case insensitive manner

    Dim funcIndex As Long
    For funcIndex = 0 To attrs.cFuncs - 1
        Dim funcDescriptior As FUNCDESC
        funcDescriptior = getFuncDesc(ITypeInfo, funcIndex)
        Dim funcName As String
        funcName = getFuncNameFromDescriptor(ITypeInfo, funcDescriptior)
        With funcDescriptior
            Debug.Print "[INFO] "; funcName & vbTab & Switch( _
                .INVOKEKIND = INVOKE_METHOD, "VbMethod", _
                .INVOKEKIND = INVOKE_PROPERTYGET, "VbGet", _
                .INVOKEKIND = INVOKE_PROPERTYPUT, "VbLet", _
                .INVOKEKIND = INVOKE_PROPERTYPUTREF, "VbSet" _
                ) & "@" & .memid

            'property get/set all have the same dispid so only need to be here once
            If Not result.Exists(funcName) Then
                result.Add funcName, .memid
            ElseIf result(funcName) <> .memid Then
                Err.Raise 5, Description:=funcName & "is already associated with another dispid"
            Else
                Debug.Assert .INVOKEKIND <> INVOKE_METHOD 'this method & dispid should not appear twice
            End If

        End With
        funcName = vbNullString
    Next
    Set GetFuncDispidFromTypeInfo = result
End Function

Public Function getFuncNameFromDescriptor(ByVal ITypeInfo As IUnknown, ByRef inFuncDescriptor As FUNCDESC) As String
     getFuncNameFromDescriptor = getDocumentation(ITypeInfo, inFuncDescriptor.memid)
End Function

Public Function getModName(ByVal ITypeInfo As IUnknown) As String
    getModName = getDocumentation(ITypeInfo, KnownMemberIDs.MEMBERID_NIL)
End Function
Private Function getDocumentation(ByVal ITypeInfo As IUnknown, ByVal memid As dispid) As String
    'HRESULT  GetDocumentation( [in] MEMBERID memid, [out] BSTR *pBstrName, [out] BSTR *pBstrDocString, [out] DWORD *pdwHelpContext, [out] BSTR *pBstrHelpFile)
    Dim hresult As hResultCode
    hresult = COMTools.CallCOMObjectVTableEntry(ITypeInfo, ITypeInfoVTableOffset(ITypeInfoVTable.getDocumentation), CR_HRESULT, memid, VarPtr(getDocumentation), NULL_PTR, NULL_PTR, NULL_PTR)
    If hresult <> S_OK Then Err.Raise hresult
End Function

Public Function getAttrs(ByVal ITypeInfo As IUnknown) As TYPEATTR
    'HRESULT  GetTypeAttr([out] TYPEATTR **ppTypeAttr )
    Dim hresult As hResultCode
    Dim pTypeAttr As LongPtr
    hresult = COMTools.CallCOMObjectVTableEntry(ITypeInfo, ITypeInfoVTableOffset(ITypeInfoVTable.GetTypeAttr), CR_HRESULT, VarPtr(pTypeAttr))
    If hresult <> S_OK Then Err.Raise hresult

    'make a local copy of the data so we can safely release the reference to the type attrs object
    'TODO Is it safe? Does this make the info in the attrs structure invalid?
    CopyMemory getAttrs, ByVal pTypeAttr, LenB(getAttrs)

    'void ITypeInfo::ReleaseTypeAttr( [in] TYPEATTR *pTypeAttr)
    COMTools.CallCOMObjectVTableEntry ITypeInfo, ITypeInfoVTableOffset(ITypeInfoVTable.ReleaseTypeAttr), CR_None, pTypeAttr
    pTypeAttr = NULL_PTR 'good practice to null released pointers so we don't accidentally use them
End Function

Public Function getFuncDesc(ByVal ITypeInfo As IUnknown, ByVal index As Long) As FUNCDESC
    'HRESULT  GetFuncDesc([in] UINT index, [out] FUNCDESC **ppFuncDesc)
    Dim hresult As hResultCode
    Dim pFuncDesc As LongPtr
    hresult = COMTools.CallCOMObjectVTableEntry(ITypeInfo, ITypeInfoVTableOffset(ITypeInfoVTable.getFuncDesc), CR_HRESULT, index, VarPtr(pFuncDesc))
    If hresult <> S_OK Then Err.Raise hresult

    'logic same as in tryGetAttrs
    CopyMemory getFuncDesc, ByVal pFuncDesc, LenB(getFuncDesc)

    'void     ReleaseFuncDesc( [in] FUNCDESC *pFuncDesc)
    COMTools.CallCOMObjectVTableEntry ITypeInfo, ITypeInfoVTableOffset(ITypeInfoVTable.ReleaseFuncDesc), CR_None, pFuncDesc
    pFuncDesc = NULL_PTR
End Function
