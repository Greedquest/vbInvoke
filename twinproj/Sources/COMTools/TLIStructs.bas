Attribute VB_Name = "TLIStructs"
'@IgnoreModule IntegerDataType
'@Folder("TLITypes")
Option Explicit

Public Type TTYPEDESC
    pTypeDesc As LongPtr
    vt As Integer
End Type

Public Type TPARAMDESC
    pPARAMDESCEX As LongPtr
    wParamFlags As Integer
End Type

Public Type TELEMDESC
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
    memid As DISPID                              'The function member ID (DispId).
    lprgscode As LongPtr                         'Pointer to status code
    lprgelemdescParam As LongPtr                 'Pointer to description of the element.
    funckind As Long                             'virtual, static, or dispatch-only
    INVOKEKIND As tagINVOKEKIND                           'VbMethod / VbGet / VbSet / VbLet
    CallConv As CALLINGCONVENTION_ENUM                             'typically will be stdecl
    cParams As Integer                           'number of parameters
    cParamsOpt As Integer                        'number of optional parameters
    oVft As Integer                              'For FUNC_VIRTUAL, specifies the offset in the VTBL.
    cScodes As Integer                           'The number of possible return values.
    elemdescFunc As TELEMDESC                    'The function return type
    wFuncFlags As Integer                        'The function flags. See FUNCFLAGS.
End Type

