Attribute VB_Name = "ExperimentalTests"
'@Folder("Tests")
'@IgnoreModule
Option Explicit
Option Private Module

Public Sub BasicInvoke()
    InvokeParamaterlessSub "ModuleReflection", "ExampleModule", "CallME"
End Sub

Public Sub PrivateInvoke(Optional ByVal functionName As String = "CallNothing")
    CallByName GetFancyAccessor("ExampleModule"), functionName, VbMethod
End Sub

Sub testanotherlib()
    Dim a As LongPtr
    a = 52
    GetFancyAccessor("LibMemory", "MemoryTools").memlongptr(VarPtr(a)) = 72
    Debug.Assert a = 72
End Sub
