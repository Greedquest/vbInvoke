Attribute VB_Name = "Scratchpad"
'@Folder("_Scratch")
Option Explicit

Public Declare PtrSafe Function GetFullAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject, ByRef outPrivateTI As IUnknown) As Object
Public Declare PtrSafe Function GetPublicAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject) As Object
Public Declare PtrSafe Function makeProxy Lib "vbInvoke_win64" (ByVal baseObj As Object) As Object

Public Declare PtrSafe Sub CallMe Lib "vbInvoke_win64" ()

'@EntryPoint
Public Sub TestPubAccessor()
    Dim accessor As Object
    Set accessor = GetPublicAccessor("ExampleModule", ThisWorkbook.VBProject)
    Debug.Print ObjPtr(accessor)
    accessor.Baz
End Sub

'@EntryPoint
Public Sub TestPubPrivAccessorHardcoded()
    Dim outTi As IUnknown
    Dim accessor As Object
    Set accessor = GetFullAccessor("ExampleModule", ThisWorkbook.VBProject, outTi)
    Debug.Print "accessor@"; ObjPtr(accessor)
    Debug.Print "ti@"; ObjPtr(outTi)
    Debug.Print accessor.Foo(13)
End Sub

'@EntryPoint
Public Sub TestProxyHardcode()
    Dim baseObject As Object
    Set baseObject = New Collection
    baseObject.Add "foo"
            
    'test we can use any IDispatch interface
    Dim proxy As stdole.IUnknown
    Set proxy = makeProxy(baseObject)
    'should call our swapped function!
    
    Dim proxyObj As Object
    Set proxyObj = proxy
    proxyObj.Add "bar"
      
    Dim item As Variant
    For Each item In baseObject
        Debug.Print "Got a " & item
    Next
End Sub


