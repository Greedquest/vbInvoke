Attribute VB_Name = "ExampleModule"
'@Folder("DemoProject")
Option Explicit

Private Function foo(ByVal bar As Long) As String
    foo = "HelloWorld" & String(bar, "!")
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

Private Property Get lemon() As String
    lemon = String$(val * 10, "l")
End Property
