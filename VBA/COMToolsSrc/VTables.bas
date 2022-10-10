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

'TODO tidy this up
'    MIDL_INTERFACE ("00020402-0000-0000-C000-000000000046")
'ITypeLib:      Public IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
'3      UINT     GetTypeInfoCount( void) = 0;
'4      HRESULT  GetTypeInfo(
'            /* [in] */ UINT index,
'            /* [out] */ __RPC__deref_out_opt ITypeInfo **ppTInfo) = 0;
'
'        virtual HRESULT STDMETHODCALLTYPE GetTypeInfoType(
'            /* [in] */ UINT index,
'            /* [out] */ __RPC__out TYPEKIND *pTKind) = 0;
'
'        virtual HRESULT STDMETHODCALLTYPE GetTypeInfoOfGuid(
'            /* [in] */ __RPC__in REFGUID guid,
'            /* [out] */ __RPC__deref_out_opt ITypeInfo **ppTinfo) = 0;
'
'        virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetLibAttr(
'            /* [out] */ TLIBATTR **ppTLibAttr) = 0;
'
'        virtual HRESULT STDMETHODCALLTYPE GetTypeComp(
'            /* [out] */ __RPC__deref_out_opt ITypeComp **ppTComp) = 0;
'
'        virtual /* [local] */ HRESULT STDMETHODCALLTYPE GetDocumentation(
'            /* [in] */ INT index,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrName,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrDocString,
'            /* [out] */ DWORD *pdwHelpContext,
'            /* [annotation][out] */
'            _Outptr_opt_  BSTR *pBstrHelpFile) = 0;
'
'        virtual /* [local] */ HRESULT STDMETHODCALLTYPE IsName(
'            /* [annotation][out][in] */
'            __RPC__inout  LPOLESTR szNameBuf,
'            /* [in] */ ULONG lHashVal,
'            /* [out] */ BOOL *pfName) = 0;
'
'        virtual /* [local] */ HRESULT STDMETHODCALLTYPE FindName(
'            /* [annotation][out][in] */
'            __RPC__inout  LPOLESTR szNameBuf,
'            /* [in] */ ULONG lHashVal,
'            /* [length_is][size_is][out] */ ITypeInfo **ppTInfo,
'            /* [length_is][size_is][out] */ MEMBERID *rgMemId,
'            /* [out][in] */ USHORT *pcFound) = 0;
'
'        virtual /* [local] */ void STDMETHODCALLTYPE ReleaseTLibAttr(
'            /* [in] */ TLIBATTR *pTLibAttr) = 0;
'
'    };

Public Type ITypeLibVTable
    IUnknown As IUnknownVTable
    GetTypeInfoCount As LongPtr
    GetTypeInfo As LongPtr
    GetTypeInfoType As LongPtr
    GetTypeInfoOfGuid As LongPtr
    GetLibAttr As LongPtr
    GetTypeComp As LongPtr
    GetDocumentation As LongPtr
    IsName As LongPtr
    FindName As LongPtr
    ReleaseTLibAttr As LongPtr
End Type: Public ITypeLibVTable As ITypeLibVTable

'[ITypeInfo](https://github.com/tpn/winsdk-10/blob/master/Include/10.0.16299.0/um/OAIdl.h#L2683) extends IUnknown
'0      HRESULT  QueryInterface ([in] REFIID riid, [out] void **ppvObject)
'1      ULONG    AddRef ()
'2      ULONG    Release ()
'3      HRESULT  GetTypeAttr([out] TYPEATTR **ppTypeAttr )
'4      HRESULT  GetTypeComp([out] ITypeComp **ppTComp )
'5      HRESULT  GetFuncDesc([in] UINT index, [out] FUNCDESC **ppFuncDesc)
'6      HRESULT  GetVarDesc([in] UINT index, [out] VARDESC **ppVarDesc)
'7      HRESULT  GetNames([in] MEMBERID memid, [out] BSTR *rgBstrNames, [in] UINT cMaxNames, [out] UINT *pcNames)
'8      HRESULT  GetRefTypeOfImplType( [in] UINT index, [out] HREFTYPE *pRefType)
'9      HRESULT  GetImplTypeFlags( [in] UINT index, [out] INT *pImplTypeFlags)
'10     HRESULT  GetIDsOfNames( [in] LPOLESTR *rgszNames, [in] UINT cNames, [out] MEMBERID *pMemId)
'11     HRESULT  Invoke( [in] PVOID pvInstance, [in] MEMBERID memid, [in] WORD wFlags, [out][in] DISPPARAMS *pDispParams, [out] VARIANT *pVarResult, [out] EXCEPINFO *pExcepInfo, [out] UINT *puArgErr)
'12     HRESULT  GetDocumentation( [in] MEMBERID memid, [out] BSTR *pBstrName, [out] BSTR *pBstrDocString, [out] DWORD *pdwHelpContext, [out] BSTR *pBstrHelpFile)
'13     HRESULT  GetDllEntry( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] BSTR *pBstrDllName, [out] BSTR *pBstrName, [out] WORD *pwOrdinal)
'14     HRESULT  GetRefTypeInfo( [in] HREFTYPE hRefType, [out] ITypeInfo **ppTInfo)
'15     HRESULT  AddressOfMember( [in] MEMBERID memid, [in] INVOKEKIND invKind, [out] PVOID *ppv)
'16     HRESULT  CreateInstance( [in] IUnknown *pUnkOuter, [in] REFIID riid, [out] PVOID *ppvObj)
'17     HRESULT  GetMops( [in] MEMBERID memid, [out] BSTR *pBstrMops)
'18     HRESULT  GetContainingTypeLib( [out] ITypeLib **ppTLib, [out] UINT *pIndex)
'19     void     ReleaseTypeAttr( [in] TYPEATTR *pTypeAttr)
'20     void     ReleaseFuncDesc( [in] FUNCDESC *pFuncDesc)
'21     void     ReleaseVarDesc( [in] VARDESC *pVarDesc)
Public Type ITypeInfoVTable
    IUnknown As IUnknownVTable
    GetTypeAttr As LongPtr
    GetTypeComp As LongPtr
    GetFuncDesc As LongPtr
    GetVarDesc As LongPtr
    GetNames As LongPtr
    GetRefTypeOfImplType As LongPtr
    GetImplTypeFlags As LongPtr
    GetIDsOfNames As LongPtr
    Invoke As LongPtr
    GetDocumentation As LongPtr
    GetDllEntry As LongPtr
    GetRefTypeInfo As LongPtr
    AddressOfMember As LongPtr
    CreateInstance As LongPtr
    GetMops As LongPtr
    GetContainingTypeLib As LongPtr
    ReleaseTypeAttr As LongPtr
    ReleaseFuncDesc As LongPtr
    ReleaseVarDesc As LongPtr
End Type: Public ITypeInfoVTable As ITypeInfoVTable

Public Property Get IUnknownVTableOffset(ByRef member As LongPtr) As LongPtr
    IUnknownVTableOffset = VarPtr(member) - VarPtr(IUnknownVTable)
End Property

'@EntryPoint
Public Property Get IDispatchVTableOffset(ByRef member As LongPtr) As LongPtr
    IDispatchVTableOffset = VarPtr(member) - VarPtr(IDispatchVTable)
End Property

'@EntryPoint
Public Property Get ITypeInfoVTableOffset(ByRef member As LongPtr) As LongPtr
    ITypeInfoVTableOffset = VarPtr(member) - VarPtr(ITypeInfoVTable)
End Property

'@EntryPoint
Public Property Get ITypeLibVTableOffset(ByRef member As LongPtr) As LongPtr
    ITypeLibVTableOffset = VarPtr(member) - VarPtr(ITypeLibVTable)
End Property

