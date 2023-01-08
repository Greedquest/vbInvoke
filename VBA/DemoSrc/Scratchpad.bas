Attribute VB_Name = "Scratchpad"
'@Folder("_Scratch")
Option Explicit

'Public Declare PtrSafe Function StdModuleAccessor Lib "C:\Users\guy\Documents\GitHub\vbInvoke\Build\vbInvoke_win64.dll" _
'(ByVal moduleName As String, ByVal vbProj As Object, ByVal projectName As String, _
'Optional ByRef outPModuleTypeInfo As LongPtr, Optional ByRef outITypeLib As LongPtr) As Object
'

Public Declare PtrSafe Function GetFullAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject, ByRef outPrivateTI As IUnknown) As Object
Public Declare PtrSafe Function GetPublicAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject) As Object
Public Declare PtrSafe Function makeProxy Lib "vbInvoke_win64" (ByVal baseObj As Object) As Object

Public Declare PtrSafe Sub CallMe Lib "vbInvoke_win64" ()

Private Type TAddLibDemo
    DllMan As DLLManager
    VBInvoke64 As TempDll
End Type

Public this As TAddLibDemo

Private Const FULL_DLL_PATH As String = "C:\GitHub\vbInvoke\Build\vbInvoke_win64.dll"

Public Sub LoadDll()
    If this.VBInvoke64 Is Nothing Then Set this.VBInvoke64 = TempDll.Create(FULL_DLL_PATH)
End Sub

Public Sub UnLoadDll()
    Set this.VBInvoke64 = Nothing
End Sub

Public Sub ForceKillDll()
    TempDll.Kill FULL_DLL_PATH
End Sub

'@EntryPoint
Public Sub TestPubAccessor()
    On Error GoTo ReleaseRef
    LoadDll
  
    Dim accessor As Object
    Set accessor = GetPublicAccessor("ExampleModule", ThisWorkbook.VBProject)
    Debug.Print ObjPtr(accessor)
    accessor.Baz
    
ReleaseRef:
    Set accessor = Nothing
    UnLoadDll
    If Err.Number <> 0 Then Debug.Print Err.Number & " - " & Err.Description
End Sub

'@EntryPoint
Public Sub TestPubPrivAccessorHardcoded()
    On Error GoTo ReleaseRef
    LoadDll

    Dim outTi As IUnknown
  
    Dim accessor As Object
    Set accessor = GetFullAccessor("ExampleModule", ThisWorkbook.VBProject, outTi)
    Debug.Print ObjPtr(accessor)
    Debug.Print ObjPtr(outTi)
        
    Debug.Print accessor.Foo(13)
    
    
ReleaseRef:
    Set accessor = Nothing
    UnLoadDll
    If Err.Number <> 0 Then Debug.Print Err.Number & " - " & Err.Description
    
End Sub

'@EntryPoint
Public Sub TestProxyHardcode()
    On Error GoTo ReleaseRef
    LoadDll
    Dim baseObject As Object
    Set baseObject = New Collection
    baseObject.Add "foo"
            
    'should call our swapped function!
        
    Dim proxy As Object
    Set proxy = makeProxy(baseObject)
    proxy.Add "bar"
    
            
    Dim item As Variant
    For Each item In baseObject
        Logger.Logg "Got a " & item
    Next
ReleaseRef:
    Set proxy = Nothing
    UnLoadDll
    If Err.Number <> 0 Then Debug.Print Err.Number & " - " & Err.Description

End Sub


