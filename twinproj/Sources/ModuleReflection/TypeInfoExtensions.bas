Attribute VB_Name = "TypeInfoExtensions"
'@Folder "TypeInfoInvoker"
Option Private Module
Option Explicit

'<Summary> An internal interface exposed by VBA for all components (modules, class modules, etc)
'<remarks> This internal interface is known to be supported since the very earliest version of VBA6
'[ComImport(), Guid("DDD557E1-D96F-11CD-9570-00AA0051E5D4")]
'[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
'Public Enum IVBEComponentVTableOffsets           '+3 for the IUnknown
'    CompileComponentOffset = 12 + 3              'void CompileComponent();
'    GetStdModAccessorOffset = 14 + 3             'IDispatch GetStdModAccessor();
'    GetSomeRelatedTypeInfoPtrsOffset = 34 + 3    'void GetSomeRelatedTypeInfoPtrs(out IntPtr a, out IntPtr b);        // returns 2 TypeInfos, seemingly related to this ITypeInfo, but slightly different.
'End Enum

Public Type IVBEComponentVTable
    IUnknown As IUnknownVTable
    placeholder(1 To 12) As LongPtr
    CompileComponent As LongPtr
    placeholder2(1 To 1) As LongPtr
    GetStdModAccessor As LongPtr
    placeholder3(1 To 19) As LongPtr
    GetSomeRelatedTypeInfoPtrs As LongPtr
End Type: Public IVBEComponentVTable As IVBEComponentVTable

Public Property Get IVBEComponentVTableOffset(ByRef member As LongPtr) As LongPtr
    IVBEComponentVTableOffset = VarPtr(member) - VarPtr(IVBEComponentVTable)
End Property

'@Description("Invoke IVBEComponent::GetStdModAccessor - re-raise error codes as VBA errors")
Public Function GetStdModAccessor(ByVal pIVBEComponent As LongPtr) As Object
Attribute GetStdModAccessor.VB_Description = "Invoke IVBEComponent::GetStdModAccessor - re-raise error codes as VBA errors"
    Dim hresult As hResultCode
    hresult = CallFunction(pIVBEComponent, IVBEComponentVTableOffset(IVBEComponentVTable.GetStdModAccessor), CR_HRESULT, CC_STDCALL, VarPtr(GetStdModAccessor))
    If hresult = S_OK Then Exit Function
    Err.Raise hresult, "GetStdModAccessor", "Function did not succeed. IVBEComponent::GetStdModAccessor HRESULT: 0x" & Hex$(hresult)
End Function
