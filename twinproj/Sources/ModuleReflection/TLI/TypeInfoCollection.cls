VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeInfoCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("TLI")
Option Explicit

Private Type TTypeInfoCollection
    ITypeLib As ITypeLib
    typeInfos As New Dictionary
    count As Long
End Type

Private this As TTypeInfoCollection

' Public Property Get ITypeLib() As ITypeLib
'     Debug.Assert Not this.ITypeLib Is Nothing
'     Set ITypeLib = this.ITypeLib
' End Property

Public Property Set ITypeLib(ByVal RHS As ITypeLib)
    Set this.ITypeLib = RHS
    this.count = TypeLibHelper.getTypeInfoCount(RHS)
    Logger.Log DebugLevel, TypeLibHelper.getProjName(RHS) & " TypeInfo.count = " & this.count
End Property

Private Function tryGenerateNext(ByRef outITypeInfo As ITypeInfo) As Boolean
    Static i As Long 'zero indexed
    If i >= this.count Then Return False
    
    'if there are errors here we want to know about them
    Dim rawITypeInfo As ITypeInfo
    Set rawITypeInfo = TypeLibHelper.getITypeInfoByIndex(this.ITypeLib, i)
    i = i + 1

    Set outITypeInfo = rawITypeInfo
    Return True
End Function

Public Function Create(ByVal wrappedITypeLib As ITypeLib) As TypeInfoCollection
    Dim result As New TypeInfoCollection
    Set result.ITypeLib = wrappedITypeLib
    Set Create = result
End Function

Public Function Find(ByVal name As String) As ITypeInfo
    Logger.Log TraceLevel, "Find ... " & name
    Do While Not this.typeInfos.Exists(name)
        Dim wrappedTI As ITypeInfo
        If Not tryGenerateNext(wrappedTI) Then Err.Raise 5, Description:="That name can't be found"
        Dim tiName As String = TypeInfoHelper.getModName(wrappedTI)
        Logger.Log DebugLevel, "found ti for " & tiName
        this.typeInfos.Add tiName, wrappedTI
    Loop
    Set Find = this.typeInfos.Item(name)
End Function
