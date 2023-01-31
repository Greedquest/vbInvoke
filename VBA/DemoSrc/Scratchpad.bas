Attribute VB_Name = "Scratchpad"
'@Folder("_Scratch")
Option Explicit

Public Declare PtrSafe Function GetExtendedModuleAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject, ByRef outPrivateTI As IUnknown) As Object
Public Declare PtrSafe Function GetStandardModuleAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject) As Object
Public Declare PtrSafe Function makeProxy Lib "vbInvoke_win64" (ByVal baseObj As Object) As Object

Public Declare PtrSafe Sub CallMe Lib "vbInvoke_win64" ()

'@EntryPoint
Public Sub TestPubAccessor()
    Dim accessor As Object
    Set accessor = GetStandardModuleAccessor("ExampleModule", ThisWorkbook.VBProject)
    Debug.Print ObjPtr(accessor)
    accessor.Baz
End Sub


'@EntryPoint
Public Sub TestPubPrivAccessorHardcoded()
    Dim outTi As IUnknown
    Dim accessor As Object
    Set accessor = GetExtendedModuleAccessor("ExampleModule", ThisWorkbook.VBProject, outTi)
    Debug.Print "accessor@"; ObjPtr(accessor)
    Debug.Print "ti@"; ObjPtr(outTi)
    Debug.Print accessor.foo(13)
    Debug.Print "accessor is: "; TypeName(accessor)
    Dim x As Variant
    For Each x In accessor
        On Error Resume Next
        Debug.Print "Calling "; x;
        CallByName accessor, x, VbMethod
        If Err.Number <> 0 Then Debug.Print " raised error: "; Err.Description;
        Debug.Print
        On Error GoTo 0
    Next x
End Sub

'@EntryPoint
Public Sub TestProxyHardcode()
    Dim baseObject As Object
    Set baseObject = New Collection
    Debug.Print "TypeName(baseObject)="; TypeName(baseObject)
    baseObject.Add "foo"

    'test we can use any IDispatch interface
    Dim proxy As stdole.IUnknown
    Set proxy = makeProxy(baseObject)
    Debug.Print "TypeName(proxy)="; TypeName(proxy)
    'should call our swapped function!

    Dim proxyObj As Object
    Set proxyObj = proxy
    Debug.Print "TypeName(proxyObj)="; TypeName(proxyObj)
    proxyObj.Add "bar"

    Dim item As Variant
    For Each item In baseObject
        Debug.Print "Got a " & item
    Next
End Sub

Sub VBInvokeTHingy()
    Dim a As Object, foo As vbInvoke.[_ITypeInfo]
    Set a = vbInvoke.GetExtendedModuleAccessor("ExampleModule", ThisWorkbook.VBProject, foo)
    Dim outName As String
    foo.GetDocumentation 1610612739, outName
    Debug.Print outName
    Debug.Print a.foo(11)

End Sub

Sub VBInvokeTHingyOptionalArg()
    Dim a As Object
    Set a = vbInvoke.GetExtendedModuleAccessor("ExampleModule", ThisWorkbook.VBProject)
    Debug.Print a.foo(11), TypeName(a)
End Sub



