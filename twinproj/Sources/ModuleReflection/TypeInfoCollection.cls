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

Public Property Get ITypeLib() As ITypeLib
    Debug.Assert Not this.ITypeLib Is Nothing
    Set ITypeLib = this.ITypeLib
End Property

Public Property Set ITypeLib(ByVal RHS As ITypeLib)
    Set this.ITypeLib = RHS
    this.count = TypeLibHelper.getTypeInfoCount(ITypeLib)
End Property

Private Function tryGenerateNext(ByRef outITypeInfo As TypeInfo) As Boolean
    Static i As Long 'zero indexed
    If i >= this.count Then Exit Function
    On Error Resume Next
        Dim rawITypeInfo As ITypeInfo
        
        Set rawITypeInfo = TypeLibHelper.getITypeInfoByIndex(ITypeLib, i)
        i = i + 1
        
        Dim noErrors As Boolean
        noErrors = Err.Number = 0
    On Error GoTo 0
    
    If noErrors Then
        Set outITypeInfo = New TypeInfo
        Set outITypeInfo.ITypeInfo = rawITypeInfo
        tryGenerateNext = True
    End If
End Function

Public Function Create(ByVal wrappedITypeLib As ITypeLib) As TypeInfoCollection
    Dim result As New TypeInfoCollection
    Set result.ITypeLib = wrappedITypeLib
    Set Create = result
End Function

Public Function Find(ByVal name As String) As TypeInfo
    Do While Not this.typeInfos.Exists(name)
        Dim wrappedTI As TypeInfo
        If Not tryGenerateNext(wrappedTI) Then Err.Raise 5, Description:="That name can't be found"
        MsgBox "found ti for " & wrappedTI.name
        this.typeInfos.Add wrappedTI.name, wrappedTI
    Loop
    Set Find = this.typeInfos.Item(name)
End Function
