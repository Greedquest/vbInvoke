VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeLibInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("TLI")
Option Explicit

Private Type TTypeLibInfo
    ITypeLib As IUnknown
    typeInfos As TypeInfoCollection
End Type

Private this As TTypeLibInfo

Public Property Get name() As String
    name = TypeLibHelper.getProjName(ITypeLib)
End Property

Public Property Get ITypeLib() As IUnknown
    Debug.Assert Not this.ITypeLib Is Nothing
    Set ITypeLib = this.ITypeLib
End Property

Public Property Set ITypeLib(ByVal RHS As IUnknown)
    Set this.ITypeLib = RHS
    Set this.typeInfos = TypeInfoCollection.Create(ITypeLib)
End Property

Public Function getTypeInfoByName(ByVal name As String) As TypeInfo
    Set getTypeInfoByName = this.typeInfos.Find(name)
End Function
