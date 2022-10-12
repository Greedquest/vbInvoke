Attribute VB_Name = "ExampleModule"
'@Folder("VBAProject")
Option Explicit

Private Function Foo(ByVal bar As Long) As String
    Foo = "HelloWorld" & String(bar, "!")
End Function

Public Sub Baz()
    MsgBox "Hi!"
End Sub

Public Property Get val() As Double
    val = Rnd()
End Property

Public Property Let val(ByVal rhs As Double)
    Debug.Print "Ignoring val="; rhs
End Property

