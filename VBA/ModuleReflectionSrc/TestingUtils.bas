Attribute VB_Name = "TestingUtils"
'@Folder("Testing")
'@IgnoreModule
Option Explicit
Option Private Module

Public Sub DispInvokeMethod(ByVal accessor As Object, ByVal dispid As Long) 'orig as object

    Dim localeID As Long 'Not really needed. Could pass 0 instead
    localeID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

    Dim outExcepInfo As COMTools.EXCEPINFOt
    
    Dim guidIID_NULL As GUIDt
    guidIID_NULL = GUIDFromString(IID_NULL)
    '@Ignore IntegerDataType
    Dim flags As Integer
    flags = tagINVOKEKIND.INVOKE_METHOD
    
    Dim params As COMTools.DISPPARAMSt 'this empty should be sufficient if no params

    Dim outResult As Variant
    Dim outFirstBadArgIndex As Long
    
    'HRESULT Invoke(
    '  [in]      DISPID     dispIdMember,
    '  [in]      REFIID     riid,
    '  [in]      LCID       lcid,
    '  [in]      WORD       wFlags,
    '  [in, out] DISPPARAMS *pDispParams,
    '  [out]     VARIANT    *pVarResult,
    '  [out]     EXCEPINFO  *pExcepInfo,
    '  [out] UINT * puArgErr
    ');
    Debug.Print "INVOKED="; ObjPtr(accessor)
    Dim hresult As hResultCode
    On Error Resume Next
    hresult = COMTools.CallFunction( _
        ObjPtr(accessor), IDispatchVTableOffset(IDispatchVTable.Invoke), _
        CR_HRESULT, CC_STDCALL, _
        dispid, _
        VarPtr(guidIID_NULL), localeID, flags, _
        VarPtr(params), _
        VarPtr(outResult), VarPtr(outExcepInfo), VarPtr(outFirstBadArgIndex) _
        )
    
    If hresult <> S_OK Then Stop
    On Error GoTo 0
End Sub

Public Sub InvokeParamaterlessSub(ByVal projectName As String, ByVal moduleName As String, ByVal methodName As String)
    Dim accessor As Object
    Set accessor = StdModuleAccessor(moduleName, projectName)
    On Error GoTo logErr
    Debug.Print "Before"
    CallByName accessor, methodName, VbMethod
    Debug.Print "After"
    Exit Sub
    
logErr:
    MsgBox Err.Number & "-" & Err.Description, vbCritical + vbOKOnly, "Error when Invoking Sub"
    Resume Next
    
End Sub
