Attribute VB_Name = "InterfaceQuerying"
'@IgnoreModule UseMeaningfulName
'@Folder "TypeInfoInvoker"
Option Explicit

Public Declare PtrSafe Function IIDFromString Lib "ole32.dll" (ByVal lpsz As LongPtr, ByRef lpiid As GUIDt) As Long

' Public Function ObjectFromObjPtr(ByVal Address As LongPtr) As IUnknown
'     '@Ignore VariableNotAssigned: Assigned ByRef
'     Dim result As IUnknown
'     '@Ignore ValueRequired: False positive
'     MemLongPtr(VarPtr(result)) = Address
'     Set ObjectFromObjPtr = result
'     '@Ignore ValueRequired: False positive
'     MemLongPtr(VarPtr(result)) = 0
' End Function

'deref without calling QI
Public Function TypeLibFromObjPtr(ByVal Address As LongPtr) As ITypeLib
    '@Ignore VariableNotAssigned: Assigned ByRef
    Dim result As ITypeLib
    '@Ignore ValueRequired: False positive
    MemLongPtr(VarPtr(result)) = Address
    Set TypeLibFromObjPtr = result
    '@Ignore ValueRequired: False positive
    MemLongPtr(VarPtr(result)) = 0
End Function



'@Ignore ParameterCanBeByVal: Passing ByVal would trigger an additional QueryInterface
Public Function QueryInterface(ByRef ClassInstance As IUnknown, ByVal InterfaceIID As String) As LongPtr

    Dim InterfaceGUID As GUIDt
    IIDFromString StrPtr(InterfaceIID), InterfaceGUID

    Dim valueWrapper0 As Variant
    Dim valueWrapper1 As Variant

    valueWrapper0 = VarPtr(InterfaceGUID)
    '@Ignore VariableNotAssigned: False Positive
    Dim retVal As LongPtr
    valueWrapper1 = VarPtr(retVal)

    Dim ptrVarValues(1) As LongPtr
    ptrVarValues(0) = VarPtr(valueWrapper0)
    ptrVarValues(1) = VarPtr(valueWrapper1)
    
    '@Ignore IntegerDataType: Integer is correct here
    Dim varTypes(1) As Integer
    varTypes(0) = VbVarType.vbLong
    varTypes(1) = VarType(retVal)
    
    Const paramCount As Long = 2
    
    Dim objAddr As LongPtr
    objAddr = ObjPtr(ClassInstance)
    
    '@Ignore VariableNotAssigned: False Positive
    Dim apiRetVal As Variant
    Dim hresult As hResultCode

    hresult = DispCallFunc(objAddr, IUnknownVTableOffset(IUnknownVTable.QueryInterface), CC_STDCALL, VbVarType.vbLong, paramCount, varTypes(0), ptrVarValues(0), apiRetVal)

    If hresult = S_OK Then
        hresult = apiRetVal
        
        If hresult = S_OK Then
        
            QueryInterface = retVal
        Else
            Err.Raise hresult, "QueryInterface", "Failed to cast to interface pointer. IUnknown::QueryInterface HRESULT: 0x" & Hex$(hresult)
        End If
    Else
        Err.Raise hresult, "DispCallFunc", "Failed to cast to interface pointer. DispCallFunc HRESULT: 0x" & Hex$(hresult)
    End If
        
End Function

'@EntryPoint
Public Function QueryInterfaceObject(ByRef ClassInstance As IUnknown, ByVal InterfaceIID As String) As IUnknown
    Set QueryInterfaceObject = ObjectFromObjPtr(QueryInterface(ClassInstance, InterfaceIID))
End Function

