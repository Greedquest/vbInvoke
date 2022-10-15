VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TLI")
Option Explicit

Private Type TTypeInfo
    ITypeInfo As ITypeInfo
End Type

Private this As TTypeInfo

Public Property Get ITypeInfo() As ITypeInfo
    Debug.Assert Not this.ITypeInfo Is Nothing
    Set ITypeInfo = this.ITypeInfo
End Property

Public Property Set ITypeInfo(ByVal RHS As ITypeInfo)
    Set this.ITypeInfo = RHS
End Property

Public Property Get name() As String
    name = getModName(ITypeInfo)
End Property

Private Function attrs() As TYPEATTR
    Static result As TYPEATTR
    'check if already set
    If result.aGUID.Data1 = 0 Then result = TypeInfoHelper.getAttrs(ITypeInfo)
    attrs = result
End Function
