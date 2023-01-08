VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempDll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@PredeclaredId
'@Folder("_Scratch")
Option Explicit

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal libFilepath As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal libFilepath As String) As LongPtr

Public ModuleHandle As LongPtr

Public Sub Kill(ByVal libFullPath As String)
    FreeLibrary GetModuleHandle(libFullPath)
    FreeLibrary GetModuleHandle(libFullPath)
End Sub

Public Function Create(ByVal libFullPath As String) As TempDll
    Debug.Assert GetModuleHandle(libFullPath) = 0 'module shouldn't already be loaded
    Dim result As New TempDll
    result.ModuleHandle = LoadLibrary(libFullPath)
    Set Create = result
End Function


Private Sub Class_Terminate()
    If ModuleHandle <> 0 Then
        FreeLibrary ModuleHandle 'LoadLibrary
        FreeLibrary ModuleHandle 'Implicit load library if it was ever used
    End If
End Sub
