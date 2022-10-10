VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDispatchVB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TypeInfoInvoker.DispatchWrapper"
Option Explicit
'@Interface
'IDispatch:: GetIDsOfNames method
'IDispatch:: GetTypeInfo method
'IDispatch:: GetTypeInfoCount method
'IDispatch:: invoke method


Public Sub GetIDsOfNamesVB( _
    ByVal riid As LongPtr, _
    ByVal namesArray As LongPtr, _
    ByVal cNames As Long, _
    ByVal lcid As Long, _
    ByVal dispidArray As LongPtr _
    )
    'HRESULT GetIDsOfNames(
    '  [in]  REFIID   riid,
    '  [in]  LPOLESTR *rgszNames,
    '  [in]  UINT     cNames,
    '  [in]  LCID     lcid,
    '  [out] dispid * rgDispId
    ');
End Sub

'Public Sub GetTypeInfoVB( _
'    ByVal iTInfo As Long, _
'    ByVal lcid As Long, _
'    ByVal ppTInfo As LongPtr _
')
'    'HRESULT GetTypeInfo(
'    '  [in]  UINT      iTInfo,
'    '  [in]  LCID      lcid,
'    '  [out] ITypeInfo **ppTInfo
'    ');
'End Sub

Public Sub InvokeVB( _
    ByVal dispIDMember As Long, _
    ByVal riid As LongPtr, _
    ByVal lcid As Long, _
    ByVal wFlags As Integer, _
    ByVal pDispParams As LongPtr, _
    ByVal pVarResult As LongPtr, _
    ByVal pExcepInfo As LongPtr, _
    ByVal puArgErr As LongPtr _
)
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
End Sub


