VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EscapeSequence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("CSharpishStringFormatter")
'@PredeclaredID
Option Explicit

Private Type tEscapeSequence
    EscapeString As String
    ReplacementString As String
    IsAsciiCharacter As Boolean
    AsciiBase As AsciiEscapeBase
End Type

Public Enum AsciiEscapeBase
    Octal
    Hexadecimal
End Enum

Private this As tEscapeSequence

Public Function Create(ByVal escape As String, ByVal replacement As String, _
                       Optional ByVal isAsciiReplacement As Boolean = False, _
                       Optional ByVal base As AsciiEscapeBase = Octal) As EscapeSequence

    With New EscapeSequence
        .EscapeString = escape
        .ReplacementString = replacement
        .IsAsciiCharacter = isAsciiReplacement
        .AsciiBase = base

        Set Create = .Self
    End With
End Function

Friend Property Get Self() As EscapeSequence
    Set Self = Me
End Property

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

Public Property Get IsAsciiCharacter() As Boolean
    IsAsciiCharacter = this.IsAsciiCharacter
End Property

Friend Property Let IsAsciiCharacter(value As Boolean)
    this.IsAsciiCharacter = value
End Property

Public Property Get AsciiBase() As AsciiEscapeBase
    AsciiBase = this.AsciiBase
End Property

Friend Property Let AsciiBase(value As AsciiEscapeBase)
    this.AsciiBase = value
End Property

Public Sub Execute(ByRef string_value As String)

    If this.IsAsciiCharacter Then
        ProcessAsciiEscape string_value, this.EscapeString

    ElseIf StringContains(string_value, this.EscapeString) Then
        string_value = Replace(string_value, this.EscapeString, this.ReplacementString)

    End If

End Sub

Private Sub ProcessAsciiEscape(ByRef format_string As String, ByVal regexPattern As String)

    Dim regex As RegExp
    Dim matches As MatchCollection, thisMatch As Match
    Dim prefix As String, char As Long

    If Not StringContains(format_string, "\") Then Exit Sub

    Set regex = New RegExp
    regex.Pattern = regexPattern
    regex.IgnoreCase = True
    regex.Global = True

    Select Case this.AsciiBase
        Case AsciiEscapeBase.Octal
            prefix = "&O"

        Case AsciiEscapeBase.Hexadecimal
            prefix = "&H"

    End Select

    Set matches = regex.Execute(format_string)
    For Each thisMatch In matches
        char = CLng(prefix & thisMatch.SubMatches(0))
        format_string = Replace(format_string, thisMatch.value, Chr$(char))
    Next

    Set regex = Nothing
    Set matches = Nothing

End Sub



