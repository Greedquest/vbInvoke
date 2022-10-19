Attribute VB_Name = "VBETypeLib"
'@Folder "TypeInfoInvoker"
Option Explicit

Public Type VBEReferencesObj
    vTable1 As LongPtr                           'To _References vtable
    vTable2 As LongPtr
    vTable3 As LongPtr
    object1 As LongPtr
    object2 As LongPtr
    typeLib As LongPtr
    placeholder1 As LongPtr
    placeholder2 As LongPtr
    RefCount As LongPtr
End Type

Public Type VBETypeLibObj
    vTable1 As LongPtr                           'To ITypeLib vtable
    vTable2 As LongPtr
    vTable3 As LongPtr
    Prev As LongPtr
    '@Ignore KeywordsUsedAsMember: Looks nice, sorry ThunderFrame
    Next As LongPtr
End Type

[ DllExport ]
Public Sub boxIt(ByVal s As String)
    MsgBox s
End Sub

[ DllExport ]
Public Function somethingSimple(ByVal x As String, ByVal proj As VBProject) As Object
    Logger.Log InfoLevel, "Invoking tB dll function..."
    Return StdModuleAccessor(x, proj, proj.Name)
End Function


Public Function StdModuleAccessor(ByVal moduleName As String, ByVal vbProj As VBProject, ByVal projectName As String, Optional ByRef outModuleTypeInfo As TypeInfo, Optional ByRef outITypeLib As LongPtr) As Object
    Dim referencesInstancePtr As LongPtr
    referencesInstancePtr = ObjPtr(vbProj.References)
    Debug.Assert referencesInstancePtr <> 0
    
    'The references object instance looks like this, and has a raw pointer contained within it to the typelibs it uses
    Dim refData As VBEReferencesObj
    CopyMemory refData, ByVal referencesInstancePtr, LenB(refData)
    Debug.Assert refData.vTable1 = MemLongPtr(referencesInstancePtr)
    
    Dim typeLibInstanceTable As VBETypeLibObj
    CopyMemory typeLibInstanceTable, ByVal refData.typeLib, LenB(typeLibInstanceTable)

    'Create a class to iterate over the doubly linked list
    Dim typeLibPtrs As New TypeLibIterator
    typeLibPtrs.baseTypeLib = refData.typeLib
    
    'Now we could use proj.module.sub to find something in particular
    'For now though, we just want a reference to the typeInfo for the ExampleModule
    Dim projectTypeLib As TypeLibInfo
    Dim found As Boolean

    Do While typeLibPtrs.TryGetNext(projectTypeLib)
        Debug.Assert typeLibPtrs.tryGetCurrentRawITypeLibPtr(outITypeLib)
        LogManager.Log InfoLevel, "Discovered: " & projectTypeLib.name
        If projectTypeLib.name = projectName Then
            'we have found the project typelib, check for the correct module within it
            Dim moduleTI As TypeInfo
            If TryGetTypeInfo(projectTypeLib, moduleName, outTI:=moduleTI) Then
                Logger.Log DebugLevel, "Got a typeinfo @" & ObjPtr(moduleTI)
                found = True
                Exit Do
            Else
                Err.Raise vbObjectError + 5, Description:="Module with name '" & moduleName & "' not found in project " & projectName
            End If
        End If
    Loop
    If Not found Then Err.Raise vbObjectError + 5, Description:="No project found with that name"

    'Cast to IVBEComponent Guid("DDD557E1-D96F-11CD-9570-00AA0051E5D4")
    '   In RD this is done via Aggregation
    '   Meaning an object is made by merging the COM interface with a managed C# interface
    '   We don't have to worry about this, it is just to avoid some bug with C# reflection I think
    
    Logger.Log DebugLevel, "Trying to cast object @" & ObjPtr(moduleTI.ITypeInfo)
    Dim vbeComponent As IVBEComponent
    Set vbeComponent = moduleTI.ITypeInfo
    Logger.Log DebugLevel, "Got an IVBEComponent @" & ObjPtr(vbeComponent)
    
    Set outModuleTypeInfo = moduleTI
    Set StdModuleAccessor = vbeComponent.GetStdModAccessor()

End Function

Private Function TryGetTypeInfo(ByVal typeLib As TypeLibInfo, ByVal moduleName As String, ByRef outTI As TypeInfo) As Boolean
    On Error Resume Next
    Set outTI = typeLib.getTypeInfoByName(moduleName)
    TryGetTypeInfo = Err.Number = 0
    On Error GoTo 0
End Function

