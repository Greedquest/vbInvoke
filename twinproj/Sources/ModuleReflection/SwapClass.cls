VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwapClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TypeInfoInvoker.DispatchWrapper"
Option Explicit

Implements IDispatchVB 'For the VTable swap
Implements IModuleInfo 'Easy access to additional methods
     
Private Declare PtrSafe Function SysReAllocString Lib "oleaut32.dll" (ByVal pBSTR As LongPtr, Optional ByVal pszStrPtr As LongPtr) As Long

Public accessor As Object
Public ITypeInfo As IUnknown

Private Sub Class_Initialize()
    Dim asDisp As IDispatchVB
    Set asDisp = Me
    Dim pAsDispVT As LongPtr
    pAsDispVT = MemLongPtr(ObjPtr(asDisp))
    Dim pInvokeVB As LongPtr, pInvokeOriginal As LongPtr
    pInvokeVB = pAsDispVT + IDispatchVBVTableOffset(IDispatchVBVTable.InvokeVB)
    pInvokeOriginal = pAsDispVT + IDispatchVBVTableOffset(IDispatchVBVTable.IDispatch.Invoke)
    
    Dim pGetIDsOfNamesVB As LongPtr, pGetIDsOfNamesOriginal As LongPtr
    pGetIDsOfNamesVB = pAsDispVT + IDispatchVBVTableOffset(IDispatchVBVTable.GetIDsOfNamesVB)
    pGetIDsOfNamesOriginal = pAsDispVT + IDispatchVBVTableOffset(IDispatchVBVTable.IDispatch.GetIDsOfNames)
    
    'swap the vtable entries
    MemLongPtr(pGetIDsOfNamesOriginal) = MemLongPtr(pGetIDsOfNamesVB)
    MemLongPtr(pInvokeOriginal) = MemLongPtr(pInvokeVB)
End Sub

Private Property Get funcs() as Dictionary
    'NOTE cached assuming you cannot modify typeinfo at all at runtime (i.e. you cannot edit a module while vba is running)
    'TODO Check if this holds True for VBComponents.Add
    Static result As Dictionary
    If result Is Nothing Then Set result = TypeInfoHelper.GetFuncDispidFromTypeInfo(ITypeInfo)
    Set funcs = result
End Property

Private Sub IDispatchVB_GetIDsOfNamesVB(ByVal riid As LongPtr, ByVal namesArray As LongPtr, ByVal cNames As Long, ByVal lcid As Long, ByVal dispidArray As LongPtr)
    'Debug.Assert cNames = 1
    Debug.Assert Not ITypeInfo Is Nothing
    Debug.Assert Not accessor Is Nothing
    Dim i As Long
    For i = 0 To cNames - 1
        Dim name As String
        name = GetStrFromPtrW(MemLongPtr(namesArray + PTR_SIZE * i))
        If funcs.Exists(name) Then
            MemLong(dispidArray + PTR_SIZE * i) = CLng(funcs(name))
        Else
            MemLong(dispidArray + PTR_SIZE * i) = -1 'unrecognised
            'REVIEW: SetLastError DISPGetIDsOfNamesErrors.DISP_E_UNKNOWNNAME ?
            Err.Raise DISPGetIDsOfNamesErrors.DISP_E_UNKNOWNNAME
        End If
    Next i
End Sub

Private Sub IDispatchVB_InvokeVB(ByVal dispIDMember As Long, ByVal riid As LongPtr, ByVal lcid As Long, ByVal wFlags As Integer, ByVal pDispParams As LongPtr, ByVal pVarResult As LongPtr, ByVal pExcepInfo As LongPtr, ByVal puArgErr As LongPtr)
    Debug.Assert Not accessor Is Nothing
    Dim hresult As hResultCode
    hresult = CallCOMObjectVTableEntry( _
        accessor, IDispatchVTableOffset(IDispatchVTable.Invoke), _
        CR_LONG, _
        dispIDMember, _
        riid, lcid, wFlags, _
        pDispParams, _
        pVarResult, pExcepInfo, puArgErr _
        )
End Sub

Private Property Get IModuleInfo_ExtendedITypeInfo() As IUnknown
    Set IModuleInfo_ExtendedITypeInfo = ITypeInfo
End Property

Private Property Get IModuleInfo_ModuleFuncInfoMap() As Dictionary
    Set IModuleInfo_ModuleFuncInfoMap = funcs
End Property

Private Property Get IModuleInfo_PublicOnlyModuleAccessor() As Object
    Set IModuleInfo_PublicOnlyModuleAccessor = accessor
End Property

Private Property Get IModuleInfo_ExtendedModuleAccessor() As Object
    Dim dipatchInterface As IDispatchVB 'need to cast me to the correct interface as only IDispatchVB is overloaded
    Set dipatchInterface = Me
    Set IModuleInfo_ExtendedModuleAccessor = dipatchInterface
End Property

'Returns a copy of a null-terminated Unicode string (LPWSTR/LPCWSTR) from the given pointer
Private Function GetStrFromPtrW(ByVal Ptr As LongPtr) As String
    SysReAllocString VarPtr(GetStrFromPtrW), Ptr
End Function

