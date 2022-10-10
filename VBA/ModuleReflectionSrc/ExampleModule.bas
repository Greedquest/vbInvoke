Attribute VB_Name = "ExampleModule"
'@Folder "Testing"
'@IgnoreModule
Option Explicit

Private Type ummmmmm
    case As Long
End Type

Private Sub CallNothing()
    Debug.Print "Hi from ExampleModule (noargs)"
End Sub

Private Sub CallME(Optional ByVal readme As Long = 2)
    Debug.Print "Hi from ExampleModule", "Readme="; readme
End Sub

Private Sub Whisper()
    Debug.Print "You can't call me I'm private"
End Sub

Private Sub Ouch()
    Err.Raise 5, Description:="You can't call me"
End Sub

Private Property Let props(ByVal thing As Variant)
    Debug.Print "props to you" & thing
End Property

