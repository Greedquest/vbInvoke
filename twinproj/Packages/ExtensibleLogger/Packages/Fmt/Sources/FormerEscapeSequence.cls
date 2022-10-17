VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormerEscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("FormerCSharpishStringFormatter")

Private Type tFormerEscapeSequence
    EscapeString As String
    ReplacementString As String
End Type

Private this As tFormerEscapeSequence
Option Explicit

Public Property Get EscapeString() As String
    EscapeString = this.EscapeString
End Property

Friend Property Let EscapeString(value As String)
    this.EscapeString = value
End Property

Public Property Get ReplacementString() As String
    ReplacementString = this.ReplacementString
End Property

Friend Property Let ReplacementString(value As String)
    this.ReplacementString = value
End Property

'Lord I wish VB6 had constructors!
Public Function Create(escape As String, replacement As String) As FormerEscapeSequence
    Dim result As FormerEscapeSequence
    Set result = New FormerEscapeSequence
    result.EscapeString = escape
    result.ReplacementString = replacement
    Set Create = result
End Function

