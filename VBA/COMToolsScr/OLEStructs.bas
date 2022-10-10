Attribute VB_Name = "OLEStructs"
'@IgnoreModule IntegerDataType
'@Folder("OLE")
Option Explicit

Public Type GUIDt
    Data1 As Long
    '@Ignore IntegerDataType
    Data2 As Integer
    '@Ignore IntegerDataType
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type EXCEPINFOt
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

Public Type DISPPARAMSt
    rgvarg As LongPtr                            '  VARIANTARG *rgvarg;
    rgdispidNamedArgs As LongPtr                 '  DISPID     *rgdispidNamedArgs;
    cArgs As Long                                '  UINT       cArgs;
    cNamedArgs As Long                           '  UINT       cNamedArgs;
End Type

Public Enum DISPID
    [_]
End Enum
