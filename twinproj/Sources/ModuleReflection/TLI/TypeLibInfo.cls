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
    ITypeLib As ITypeLib
    typeInfos As TypeInfoCollection
End Type

Private this As TTypeLibInfo

Public Sub New(ByVal pTypeLib As LongPtr)
    set me.ITypeLib = ObjectFromObjPtr(pTypeLib)
End Sub

Public Property Get name() As String
    name = TypeLibHelper.getProjName(ITypeLib)
End Property

Public Property Get ITypeLib() As ITypeLib
    Assert Not this.ITypeLib Is Nothing
    Set ITypeLib = this.ITypeLib
End Property

Public Property Set ITypeLib(ByVal RHS As ITypeLib)
    Set this.ITypeLib = RHS
    Set this.typeInfos = New TypeInfoCollection(ITypeLib)
End Property

Public Function tryGetTypeInfoByName(ByVal name As String, ByRef outTI As ITypeInfo) As Boolean
    On Error GoTo logErr
    this.typeInfos.Find name, outTI
    Return True
    
    logErr:
    Logger.Log ErrorLevel, printf("Find raised error {}: {}", Err.Number, Err.Description)
    Return False
End Function
