' Attribute VB_Name = "CrashingCode"
' '@Folder "_Excel.LegacySamples"
' '@IgnoreModule
' Option Explicit
' Option Private Module

' 'Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" (ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, ByVal cc As tagCALLCONV, ByVal vtReturn As Integer, ByVal cActuals As Long, ByRef prgvt As Integer, ByRef prgpvarg As LongPtr, ByRef pvargResult As Variant) As Long
' 'Private Declare PtrSafe Function DispGetIDsOfNames Lib "oleaut32.dll" (ByVal ptinfo As LongPtr, ByVal rgszNames As LongPtr, ByVal cNames As Long, ByVal rgDispId As LongPtr) As Long


' Private Type INVOKE_ARGS
'     args() As Variant
'     argsVT() As Integer
'     #If VBA7 Then
'         argsPtrs() As LongPtr
'     #Else
'         argsPtrs() As Long
'     #End If
'     argsCount As Long
' End Type

' #If Win64 Then
'     Private Const PTR_SIZE As Long = 8
' #Else
'     Private Const PTR_SIZE As Long = 4
' #End If

' 'IDispatch derives from the IUnknown interface
' Private Enum IDispatchVtblOffset
'     oQueryInterface = PTR_SIZE * 0   'IUnknown
'     oAddRef = PTR_SIZE * 1           'IUnknown
'     oRelease = PTR_SIZE * 2          'IUnknown
'     oGetTypeInfoCount = PTR_SIZE * 3 'IDispatch
'     oGetTypeInfo = PTR_SIZE * 4      'IDispatch
'     oGetIDsOfNames = PTR_SIZE * 5    'IDispatch
'     oInvoke = PTR_SIZE * 6           'IDispatch
' End Enum

' 'ITypeInfo derives from the IUnknown interface
' Private Enum ITypeInfoVtblOffset
'     oQueryInterface = PTR_SIZE * 0   'IUnknown
'     oAddRef = PTR_SIZE * 1           'IUnknown
'     oRelease = PTR_SIZE * 2          'IUnknown
'     oGetTypeAttr = PTR_SIZE * 3
'     oGetTypeComp = PTR_SIZE * 4
'     oGetFuncDesc = PTR_SIZE * 5
'     oGetVarDesc = PTR_SIZE * 6
'     oGetNames = PTR_SIZE * 7
'     oGetRefTypeOfImplType = PTR_SIZE * 8
'     oGetImplTypeFlags = PTR_SIZE * 9
'     oGetIDsOfNames = PTR_SIZE * 10
'     oInvoke = PTR_SIZE * 11
'     oGetDocumentation = PTR_SIZE * 12
'     oGetDllEntry = PTR_SIZE * 13
'     oGetRefTypeInfo = PTR_SIZE * 14
'     oAddressOfMember = PTR_SIZE * 15
'     oCreateInstance = PTR_SIZE * 16
'     oGetMops = PTR_SIZE * 17
'     oGetContainingTypeLib = PTR_SIZE * 18
'     oReleaseTypeAttr = PTR_SIZE * 19
'     oReleaseFuncDesc = PTR_SIZE * 20
'     oReleaseVarDesc = PTR_SIZE * 21
' End Enum


' #If VBA7 Then
' Public Function GetAddressOfClassMethod(ByVal classInstance As Object, ByVal methodName As String) As LongPtr
' #Else
' Public Function GetAddressOfClassMethod(ByVal classInstance As Object, ByVal methodName As String) As Long
' #End If
'     #If VBA7 Then
'         Dim iDispatchPtr As LongPtr
'         Dim iTypeInfoPtr As LongPtr
'     #Else
'         Dim iDispatchPtr As Long
'         Dim iTypeInfoPtr As Long
'     #End If
'     Dim localeID As Long 'Not really needed. Could pass 0 instead
'     '
'     'Get a pointer to the IDispatch interface
'     iDispatchPtr = ObjPtr(GetDefaultInterface(classInstance))
'     '
'     'Get a pointer to the ITypeInfo interface
'     localeID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'     IDispatch_GetTypeInfo iDispatchPtr, 0, localeID, iTypeInfoPtr
'     '
'     Dim arrNames(0 To 0) As String: arrNames(0) = methodName
'     Dim arrIDs(0 To 0) As Long
'     '
'     'Get ID of required member
'     DispGetIDsOfNames iTypeInfoPtr, VarPtr(arrNames(0)), 1, VarPtr(arrIDs(0))
'     '
'     'Get address of member
'     ITypeInfo_AddressOfMember iTypeInfoPtr, arrIDs(0), INVOKE_METHOD, GetAddressOfClassMethod
' End Function

' ''*******************************************************************************
' ''Returns the default interface for an object
' ''All VB intefaces are dual interfaces meaning all interfaces are derived from
' ''   IDispatch which in turn is derived from IUnknown. In VB the Object datatype
' ''   stands for the IDispatch interface.
' ''Casting from a custom interface (derived only from IUnknown) to IDispatch
' ''   forces a call to QueryInterface for the IDispatch interface (which knows
' ''   about the default interface)
' ''*******************************************************************************
' 'Private Function GetDefaultInterface(obj As IUnknown) As Object
' '    Set GetDefaultInterface = obj
' 'End Function

' '*******************************************************************************
' 'IDispatch::GetTypeInfo
' '*******************************************************************************
' Private Function IDispatch_GetTypeInfo(ByVal iDispatchPtr As LongPtr, ByVal iTInfo As Long, ByVal lcid As Long, ByRef ppTInfo As LongPtr) As Long
'     Dim hresult As Long
'     '
'     With CreateInvokeArgs(iTInfo, lcid, VarPtr(ppTInfo))
'         hresult = DispCallFunc(iDispatchPtr, IDispatchVtblOffset.oGetTypeInfo, CC_STDCALL, vbLong, .argsCount, .argsVT(0), .argsPtrs(0), IDispatch_GetTypeInfo)
'     End With
'     If hresult <> S_OK Then Err.Raise hresult, "IDispatch_GetTypeInfo"
' End Function

' '*******************************************************************************
' 'ITypeInfo::AddressOfMember
' '*******************************************************************************

' Private Function ITypeInfo_AddressOfMember(ByVal iTypeInfoPtr As LongPtr, ByVal memid As Long, ByVal invKind As tagINVOKEKIND, ByRef ppv As LongPtr) As Long
'     Dim hresult As Long
'     '
'     With CreateInvokeArgs(memid, invKind, VarPtr(ppv))
'         hresult = DispCallFunc(iTypeInfoPtr, ITypeInfoVtblOffset.oAddressOfMember, CC_STDCALL, vbLong, .argsCount, .argsVT(0), .argsPtrs(0), ITypeInfo_AddressOfMember)
'     End With
'     If hresult <> S_OK Then Err.Raise hresult, "ITypeInfo_AddressOfMember"
' End Function

' '*******************************************************************************
' 'Helper function that creates the necessary arrays to use with DispCallFunc
' 'Passing arguments:
' '   - ByVal: pass the arg
' '   - ByRef: pass VarPtr(arg)
' '*******************************************************************************
' Private Function CreateInvokeArgs(ParamArray args() As Variant) As INVOKE_ARGS
'     With CreateInvokeArgs
'         .argsCount = UBound(args) + 1 'ParamArray is always 0-based (LBound)
'         If .argsCount = 0 Then
'             ReDim .argsVT(0 To 0)
'             ReDim .argsPtrs(0 To 0)
'             Exit Function
'         End If
'         '
'         .args = args 'Avoid ByRef issues by making a copy
'         ReDim .argsVT(0 To .argsCount - 1)
'         ReDim .argsPtrs(0 To .argsCount - 1)
'         Dim i As Long
'         '
'         'For Each is not used because it does copies of the values inside the
'         '   array and we need the actual addresses of the values (ByRef)
'         For i = 0 To .argsCount - 1
'             .argsVT(i) = VarType(.args(i))
'             .argsPtrs(i) = VarPtr(.args(i))
'         Next i
'     End With
' End Function
