Attribute VB_Name = "ExampleModule2"
'@Folder("Testing")
'@IgnoreModule
Option Explicit
Option Private Module

Public Sub CallNothing()
    Debug.Print "Hi from ExampleModule2 (noargs)"
End Sub

Public Sub CallME(Optional ByVal readme As Long = 2)
    Debug.Print "Hi from ExampleModule2", "Readme="; readme
End Sub

Private Sub Whisper()
    Debug.Print "You can't call me I'm private"
End Sub

Private Sub Ouch()
    Err.Raise 5, Description:="You can't call me"
End Sub

Public Function Hi(ByVal person As String) As String
    Hi = "Hello hello " & person
End Function
