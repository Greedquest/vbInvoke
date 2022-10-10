Attribute VB_Name = "IDispatchHook"
'@Folder "_Excel.LegacySamples"
'@IgnoreModule
Option Explicit
Option Private Module

#If VBA7 Then
    Private newInvokePtr As LongPtr
    Private oldInvokePtr As LongPtr
    Private invokeVtblPtr As LongPtr
#Else
    Private newInvokePtr As Long
    Private oldInvokePtr As Long
    Private invokeVtblPtr As Long
#End If

'https://docs.microsoft.com/en-us/windows/win32/api/oaidl/nf-oaidl-idispatch-invoke
Function IDispatch_Invoke(ByVal this As Object _
    , ByVal dispIDMember As Long _
    , ByVal riid As LongPtr _
    , ByVal lcid As Long _
    , ByVal wFlags As Integer _
    , ByVal pDispParams As LongPtr _
    , ByVal pVarResult As LongPtr _
    , ByVal pExcepInfo As LongPtr _
    , ByRef puArgErr As Long _
) As Long
    Const DISP_E_MEMBERNOTFOUND = &H80020003
    '
    Debug.Print "The IDispatch::Invoke return address " & VarPtr(IDispatch_Invoke) & " should be outside of the"
    IDispatch_Invoke = DISP_E_MEMBERNOTFOUND
End Function

Sub HookInvoke(obj As Object)
    If obj Is Nothing Then Exit Sub
    #If VBA7 Then
        Dim vTablePtr As LongPtr
    #Else
        Dim vTablePtr As Long
    #End If
    '
    newInvokePtr = VBA.Int(AddressOf IDispatch_Invoke)
    CopyMemory vTablePtr, ByVal ObjPtr(obj), PTR_SIZE
    '
    invokeVtblPtr = vTablePtr + 6 * PTR_SIZE
    CopyMemory oldInvokePtr, ByVal invokeVtblPtr, PTR_SIZE
    CopyMemory ByVal invokeVtblPtr, newInvokePtr, PTR_SIZE
End Sub

Sub RestoreInvoke()
    If invokeVtblPtr = 0 Then Exit Sub
    '
    CopyMemory ByVal invokeVtblPtr, oldInvokePtr, PTR_SIZE
    invokeVtblPtr = 0
    oldInvokePtr = 0
    newInvokePtr = 0
End Sub


