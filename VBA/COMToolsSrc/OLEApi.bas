Attribute VB_Name = "OLEApi"
'@Folder("OLE")
Option Explicit

'#define DISPATCH_METHOD         0x1
'#define DISPATCH_PROPERTYGET    0x2
'#define DISPATCH_PROPERTYPUT    0x4
'#define DISPATCH_PROPERTYPUTREF 0x8
Public Enum tagINVOKEKIND
    INVOKE_METHOD = &H1
    INVOKE_PROPERTYGET = &H2
    INVOKE_PROPERTYPUT = &H4
    INVOKE_PROPERTYPUTREF = &H8
End Enum

'HRESULT DispGetIDsOfNames(
'        ITypeInfo *ptinfo,
'  [in]  LPOLESTR  *rgszNames,
'        UINT      cNames,
'  [out] DISPID * rgDispId
');
'@EntryPoint
Public Declare PtrSafe Function DispGetIDsOfNames Lib "oleaut32.dll" ( _
    ByVal ptinfo As LongPtr, ByVal rgszNames As LongPtr, ByVal cNames As Long, ByVal rgDispId As LongPtr _
) As hResultCode

'CreateStdDispatch(
'  IUnknown  *punkOuter,
'  void      *pvThis,
'  ITypeInfo *ptinfo,
'  IUnknown  **ppunkStdDisp
');
'@EntryPoint
Public Declare PtrSafe Function CreateStdDispatch Lib "oleaut32.dll" ( _
    ByVal punkOuter As LongPtr, ByVal pvThis As LongPtr, ByVal ptinfo As LongPtr, ByRef outDispUnk As LongPtr _
) As hResultCode


Public Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, ByVal offsetVtable As LongPtr, ByVal CallConv As CALLINGCONVENTION_ENUM, ByVal vartypeReturn As Integer, _
    ByVal paramCount As Long, ByRef paramTypes As Integer, ByRef paramValues As LongPtr, ByRef returnValue As Variant _
) As hResultCode


