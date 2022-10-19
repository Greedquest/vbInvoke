Attribute VB_Name = "DispatchVBTypes"
'@Folder "TypeInfoInvoker.DispatchWrapper"
Option Private Module
Option Explicit

'https://github.com/wine-mirror/wine/blob/master/include/winerror.h
'TODO move to COMtools
Public Enum DISPGetIDsOfNamesErrors
      DISP_E_UNKNOWNNAME = &H80020006
      DISP_E_UNKNOWNLCID = &H8002000C
End Enum

' Public Type IDispatchVBVTable
'     IDispatch As IDispatchVTable
'     GetIDsOfNamesVB As LongPtr
'     InvokeVB As LongPtr
' End Type: Public IDispatchVBVTable As IDispatchVBVTable

' Public Property Get IDispatchVBVTableOffset(ByRef member As LongPtr) As LongPtr
'     IDispatchVBVTableOffset = VarPtr(member) - VarPtr(IDispatchVBVTable)
' End Property

