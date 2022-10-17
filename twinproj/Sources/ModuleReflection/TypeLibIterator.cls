VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TypeLibIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "TypeInfoInvoker"
Option Explicit

Private Type TIterator
    currentTL As VBETypeLibObj
    pCurrentTL As LongPtr
End Type

Private this As TIterator

Public Property Let baseTypeLib(ByVal rawptr As LongPtr)
    currentTL = rawptr
    ResetIteration
End Property

Private Property Let currentTL(ByVal rawptr As LongPtr)
    this.pCurrentTL = rawptr
    CopyMemory this.currentTL, ByVal rawptr, LenB(this.currentTL)
End Property

Public Sub ResetIteration()
    Do While this.currentTL.Prev <> 0
        currentTL = this.currentTL.Prev
    Loop
End Sub

Private Function NextTypeLib() As LongPtr
    If this.currentTL.Next = 0 Then Err.Raise 5, Description:="We've reached the end of the line"
    NextTypeLib = this.currentTL.Next
    currentTL = this.currentTL.Next 'move the iterator along
End Function

'@Description("Gets type library com objects from list")
Public Function TryGetNext(ByRef outTypeLib As TypeLibInfo) As Boolean
Attribute TryGetNext.VB_Description = "Gets type library com objects from list"
    On Error GoTo cleanFail
    Dim tlPtr As LongPtr
    tlPtr = NextTypeLib
    Set outTypeLib = TLI.TypeLibInfoFromITypeLib(TypeLibFromObjPtr(tlPtr))
    TryGetNext = True
    
cleanExit:
    Exit Function
    
cleanFail:
    TryGetNext = False
    Set outTypeLib = Nothing
    Resume cleanExit
End Function

'@Description("Returns the raw ITypeLib interface; this is because TLI.TypeLibInfo is a slightly more restricted view than the pointer here and hides private members")
Public Function tryGetCurrentRawITypeLibPtr(ByRef outITypeLib As LongPtr) As Boolean
    If this.pCurrentTL <= 0 Then Exit Function
    outITypeLib = this.pCurrentTL
    tryGetCurrentRawITypeLibPtr = True
End Function
