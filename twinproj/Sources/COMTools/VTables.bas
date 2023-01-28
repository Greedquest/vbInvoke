Attribute VB_Name = "VTables"
'@Folder("TLITypes")
Option Explicit

'[IUnknown](https://en.wikipedia.org/wiki/IUnknown)
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
Public Type IUnknownVTable
    QueryInterface As LongPtr
    AddRef As LongPtr
    ReleaseRef As LongPtr
End Type: Public IUnknownVTable As IUnknownVTable

'[IDispatch](https://en.wikipedia.org/wiki/IDispatch)  extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
'3      HRESULT  GetTypeInfoCount(unsigned int * pctinfo)
'4      HRESULT  GetTypeInfo(unsigned int iTInfo, LCID lcid, ITypeInfo ** ppTInfo)
'5      HRESULT  GetIDsOfNames(REFIID riid, OLECHAR ** rgszNames, unsigned int cNames, LCID lcid, DISPID * rgDispId)
'6      HRESULT  Invoke(DISPID dispIdMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS * pDispParams, VARIANT * pVarResult, EXCEPINFO * pExcepInfo, unsigned int * puArgErr)
Public Type IDispatchVTable
    IUnknown As IUnknownVTable
    GetTypeInfoCount As LongPtr
    GetTypeInfo As LongPtr
    GetIDsOfNames As LongPtr
    Invoke As LongPtr
End Type: Public IDispatchVTable As IDispatchVTable


Public Property Get IUnknownVTableOffset(ByRef member As LongPtr) As LongPtr
    IUnknownVTableOffset = VarPtr(member) - VarPtr(IUnknownVTable)
End Property

'@EntryPoint
Public Property Get IDispatchVTableOffset(ByRef member As LongPtr) As LongPtr
    IDispatchVTableOffset = VarPtr(member) - VarPtr(IDispatchVTable)
End Property

