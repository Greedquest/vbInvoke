Attribute VB_Name = "FunctionInvocation"
'@Folder("TypeInfoInvoker")
Option Explicit

' for documentation on the main API DispCallFunc... http://msdn.microsoft.com/en-us/library/windows/desktop/ms221473%28v=vs.85%29.aspx
Public Enum hResultCode
    S_OK = 0
End Enum

Public Enum CALLINGCONVENTION_ENUM
    ' http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.comtypes.callconv%28v=vs.110%29.aspx
    CC_FASTCALL = 0&
    CC_CDECL

    CC_PASCAL
    CC_MACPASCAL
    CC_STDCALL                                   ' typical windows APIs
    CC_FPFASTCALL
    CC_SYSCALL
    CC_MPWCDECL
    CC_MPWPASCAL
End Enum

Public Enum CALLRETURNTUYPE_ENUM
    CR_None = vbEmpty
    CR_LONG = vbLong
    CR_BYTE = vbByte
    CR_INTEGER = vbInteger
    CR_SINGLE = vbSingle
    CR_DOUBLE = vbDouble
    CR_CURRENCY = vbCurrency
    CR_HRESULT = CR_LONG 'alias because it comes up so often
    [_CR_Dispatch] = vbObject
    ' if the value you need isn't in above list, you can pass the value manually to the
    ' CallFunction_DLL method below. For additional values, see:
    ' http://msdn.microsoft.com/en-us/library/cc237865.aspx
End Enum

Public Enum STRINGPARAMS_ENUM
    STR_NONE = 0&
    STR_ANSI
    STR_UNICODE
End Enum


#If VBA7 Then
    Public Declare PtrSafe Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare PtrSafe Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
    Private Declare PtrSafe Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
#Else
    Public Declare Sub SetLastError Lib "kernel32.dll" (ByVal dwErrCode As Long)
    Private Declare Function lstrlenA Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As LongPtr) As Long
#End If

Public Function CallFunction(ByVal InterfacePointer As LongPtr, ByVal VTableByteOffsetOrFunction As LongPtr, _
                             ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                             ByVal CallConvention As CALLINGCONVENTION_ENUM, _
                             ParamArray FunctionParameters() As Variant) As Variant
    Dim vParams() As Variant
    '@Ignore DefaultMemberRequired: apparently not since this code works fine
    vParams() = FunctionParameters()             ' copy passed parameters, if any
    CallFunction = DispCallFunctionWrapper(InterfacePointer, VTableByteOffsetOrFunction, FunctionReturnType, CallConvention, vParams)

End Function

Public Function CallVBAFuncPtr(ByVal FuncPtr As LongPtr, _
                             ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                             ParamArray FunctionParameters() As Variant) As Variant
    Dim vParams() As Variant
    '@Ignore DefaultMemberRequired: apparently not since this code works fine
    vParams() = FunctionParameters()             ' copy passed parameters, if any
    CallVBAFuncPtr = DispCallFunctionWrapper(0, FuncPtr, FunctionReturnType, CC_STDCALL, vParams)
End Function

Public Function CallCOMObjectVTableEntry(ByRef COMInterface As IUnknown, ByVal VTableByteOffset As LongPtr, _
                             ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                             ParamArray FunctionParameters() As Variant) As Variant
    Dim vParams() As Variant
    '@Ignore DefaultMemberRequired: apparently not since this code works fine
    vParams() = FunctionParameters()             ' copy passed parameters, if any
    CallCOMObjectVTableEntry = DispCallFunctionWrapper(ObjPtr(COMInterface), VTableByteOffset, FunctionReturnType, CC_STDCALL, vParams)
End Function

Private Function DispCallFunctionWrapper(ByVal InterfacePointer As LongPtr, ByVal VTableByteOffsetOrFunction As LongPtr, _
                             ByVal FunctionReturnType As CALLRETURNTUYPE_ENUM, _
                             ByVal CallConvention As CALLINGCONVENTION_ENUM, _
                             ByRef vParams() As Variant) As Variant

    ' Used to call active-x or COM objects, not standard dlls

    ' Return value. Will be a variant containing a value of FunctionReturnType
    '   If this method fails, the return value will always be Empty. This can be verified by checking
    '       the Err.LastDLLError value. It will be non-zero if the function failed else zero.
    '   If the method succeeds, there is no guarantee that the Interface function you called succeeded. The
    '       success/failure of that function would be indicated by this method's return value.
    '       Typically, success is returned as S_OK (zero) and any other value is an error code.
    '   If calling a sub vs function & this method succeeds, the return value will be zero.
    '   Summarizing: if method fails to execute, Err.LastDLLError value will be non-zero
    '       If method executes ok, if the return value is zero, method succeeded else return is error code

    ' Parameters:
    '   InterfacePointer (Optional). A pointer to an object/class, i.e., ObjPtr(IPicture)
    '       Passing invalid pointers likely to result in crashes
    '   VTableOffsetOrFunction. The offset from the passed InterfacePointer where the virtual function exists.
    '       These offsets are generally in multiples of PTR_SIZE. Value cannot be negative.
    '   For the remaining parameters, see the details withn the CallFunction_DLL method.
    '       They are the same with one exception: strings. Pass the string variable name or value

    '// minimal sanity check for these 4 parameters:
    If VTableByteOffsetOrFunction < 0& Or ((InterfacePointer = 0) And (VTableByteOffsetOrFunction = 0)) Then Exit Function
    If Not (FunctionReturnType And &HFFFF0000) = 0& Then Exit Function ' can only be 4 bytes


    Dim pCount As Long
    pCount = Abs(UBound(vParams) - LBound(vParams) + 1&)

    Dim vParamPtr() As LongPtr
    Dim vParamType() As Integer
    If pCount = 0& Then                          ' no return value (sub vs function)
        ReDim vParamPtr(0 To 0)
        ReDim vParamType(0 To 0)
    Else
        ReDim vParamPtr(0 To pCount - 1&)        ' need matching array of parameter types
        ReDim vParamType(0 To pCount - 1&)       ' and pointers to the parameters
        Dim pIndex As Long
        For pIndex = 0& To pCount - 1&
            vParamPtr(pIndex) = VarPtr(vParams(pIndex))
            vParamType(pIndex) = VarType(vParams(pIndex))
        Next
    End If
    ' call the function now
    Dim vRtn As Variant
    Dim hresult As hResultCode
    hresult = DispCallFunc(InterfacePointer, VTableByteOffsetOrFunction, CallConvention, FunctionReturnType, _
                          pCount, vParamType(0), vParamPtr(0), vRtn)

    If hresult = S_OK Then
        DispCallFunctionWrapper = vRtn                      ' return result
    Else
        SetLastError hresult                      ' set error & return Empty
    End If

End Function

'@EntryPoint
Public Function PointerToStringA(ByVal ANSIpointer As LongPtr) As String
    ' courtesy function provided for your use as needed
    ' ANSIpointer must be a pointer to an ANSI string (1 byte per character)
    Dim lSize As Long, sANSI As String
    If Not ANSIpointer = 0& Then
        lSize = lstrlenA(ANSIpointer)
        If lSize > 0& Then
            sANSI = String$(lSize \ 2& + 1&, vbNullChar)
            CopyMemory ByVal StrPtr(sANSI), ByVal ANSIpointer, lSize
            PointerToStringA = Left$(StrConv(sANSI, vbUnicode), lSize)
        End If
    End If
End Function

'@EntryPoint
Public Function PointerToStringW(ByVal UnicodePointer As LongPtr) As String
    ' courtesy function provided for your use as needed
    ' UnicodePointer must be a pointer to an unicode string (2 bytes per character)
    Dim lSize As Long
    If Not UnicodePointer = 0 Then
        lSize = lstrlenW(UnicodePointer)
        If lSize > 0& Then
            PointerToStringW = Space$(lSize)
            CopyMemory ByVal StrPtr(PointerToStringW), ByVal UnicodePointer, lSize * 2&
        End If
    End If
End Function


