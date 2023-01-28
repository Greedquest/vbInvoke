Attribute VB_Name = "DllTools"
'@Folder("_Scratch")
Option Explicit

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal libFilepath As String) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal libFilepath As String) As LongPtr

Public Const DLLNAME As String = "vbInvoke_win64.dll"
Public Const FILEPATH As String = "C:\GitHub\vbInvoke\Build\" & DLLNAME
Public Const TEMP_PATH As String = "C:\Users\guy\AppData\Local\" & DLLNAME

Public Sub RefreshDll()
    Dim existingHandle As LongPtr
    existingHandle = GetModuleHandle(DLLNAME)
    
    If existingHandle <> 0 Then
        Debug.Assert FreeLibrary(existingHandle) = 1 'released our LoadLibrary call at least
        If FreeLibrary(existingHandle) = 0 Then Debug.Print "[INFO] DLL was not used by VBA since last refresh"
        Debug.Assert GetModuleHandle(DLLNAME) = 0 'DLL has been loaded more than twice, which is an issue
    Else
        Debug.Print "[WARN] DLL not already loaded"
    End If
    
    FileCopy FILEPATH, TEMP_PATH 'allows overwriting temp path unlike Name...As
    Kill FILEPATH 'ensures we call refresh only when there is a new build
    Dim newHandle As LongPtr
    newHandle = LoadLibrary(TEMP_PATH)
    Debug.Assert newHandle <> 0
    Debug.Assert newHandle = GetModuleHandle(DLLNAME)
End Sub
