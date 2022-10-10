Attribute VB_Name = "Structures"
'@Folder("TLITypes")
Option Explicit

Public Type GUIDt
    Data1 As Long
    '@Ignore IntegerDataType
    Data2 As Integer
    '@Ignore IntegerDataType
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type TTYPEDESC
    pTypeDesc As LongPtr
    vt As Integer
End Type

Private Type TPARAMDESC
    pPARAMDESCEX As LongPtr
    wParamFlags As Integer
End Type

Private Type TELEMDESC
    tdesc  As TTYPEDESC
    pdesc  As TPARAMDESC
End Type

Public Type TYPEATTR
    aGUID As GUIDt
    LCID As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Integer
    typekind As Long
    cFuncs As Integer
    cVars As Integer
    cImplTypes As Integer
    cbSizeVft As Integer
    cbAlignment As Integer
    wTypeFlags As Integer
    wMajorVerNum As Integer
    wMinorVerNum As Integer
    tdescAlias As Long
    idldescType As Long
End Type

Public Type FUNCDESC
    memid As Long                                'The function member ID (DispId).
    lprgscode As LongPtr                         'Pointer to status code
    lprgelemdescParam As LongPtr                 'Pointer to description of the element.
    funckind As Long                             'virtual, static, or dispatch-only
    INVOKEKIND As Long                           'VbMethod / VbGet / VbSet / VbLet
    CallConv As Long                             'typically will be stdecl
    cParams As Integer                           'number of parameters
    cParamsOpt As Integer                        'number of optional parameters
    oVft As Integer                              'For FUNC_VIRTUAL, specifies the offset in the VTBL.
    cScodes As Integer                           'The number of possible return values.
    elemdescFunc As TELEMDESC                    'The function return type
    wFuncFlags As Integer                        'The function flags. See FUNCFLAGS.
End Type

Public Type EXCEPINFO
    wCode As Integer
    wReserved As Integer
    bstrSource As String
    bstrDescription As String
    bstrHelpFile As String
    dwHelpContext As Long
    pvReserved As LongPtr
    pfnDeferredFillIn As LongPtr
    scode As Long
End Type

