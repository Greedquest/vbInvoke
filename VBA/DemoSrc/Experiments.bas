Attribute VB_Name = "Experiments"
'@Folder("DemoProject")
Option Explicit

Private test_val As Double

Public Sub testCallingPrivateMethod()
    Dim exampleMod As Object
    Set exampleMod = GetFancyAccessor("ExampleModule", ThisWorkbook.VBProject.Name)
    Debug.Assert exampleMod.Foo(2) = "HelloWorld!!"
End Sub

Public Sub testCallingErrorMethod()
    Dim thisMod As Object
    Set thisMod = GetFancyAccessor("Experiments")
    On Error Resume Next 'untrappable errors unfortunately, but also does not crash which is very good
    thisMod.raisesError
    Debug.Assert Err.Number = 5
End Sub

Public Sub testTerminateDoesNotCrash()
    GetFancyAccessor("Experiments").terminate
    Debug.Assert False 'unreachable
End Sub

Public Sub testModuleReflection()
    Dim info As IModuleInfo
    Set info = GetFancyAccessor("ExampleModule")
    Debug.Assert Join(info.ModuleFuncInfoMap.Keys()) = Join(Array("Foo", "Baz", "val")) 'val appears twice in module but only once here since it is a Let/Set
End Sub

Public Sub testProps()
    test_val = 9
    With GetFancyAccessor("Experiments")
        .val = 1.7
        Debug.Assert .val = 1.7
    End With
End Sub

Private Sub raisesError()
    Err.Raise 5
End Sub

Private Sub terminate()
    End
End Sub

Public Property Get val() As Double
    val = test_val
End Property

Public Property Let val(ByVal rhs As Double)
    test_val = rhs
End Property
