Attribute VB_Name = "typeLibPrinter"
'@Folder "_Excel.LegacySamples"
'@IgnoreModule
''   Adding the reference to the typelib:
''
''    It seems that earlier versions of Windows or Office were able to add the type lib from a (registered?) guid:
''      - application.VBE.activeVBProject.references.addFromGuid "{8B217740-717D-11CE-AB5B-D41203C10000}", 1, 0
''
''    wheares newer versions of Windows or Office require an installation of Visual Studio(?):
''
''      - application.VBE.activeVBProject.references.AddFromFile "c:\Program Files (x86)\Microsoft Visual Studio 14.0\Common7\IDE\vstlbinf.dll"
''
''https://renenyffenegger.ch/notes/development/languages/VBA/Useful-object-libraries/TypeLib-Information/index
''
'Option Explicit
'Option Private Module
'
'Sub propertiesOfObj(ByVal obj As Object)
'    Dim tlApp  As New TLI.TLIApplication
'    Dim tlInfo As TLI.TypeInfo
'    Set tlInfo = tlApp.InterfaceInfoFromObject(obj)
'    dumpTypeInfo tlInfo
'End Sub
'
'Sub dumpTypeInfo(ByVal tlInfo As TLI.TypeInfo)
'    Dim n As Integer
'    n = FreeFile()
'    Open ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.name & " output.txt" For Append As #n
'
'    Dim attributes() As String
'    Dim ix           As Long
'    Dim nofAttrs     As Long
'
'    nofAttrs = tlInfo.AttributeStrings(attributes)
'
'    Print #n, "Name             = " & tlInfo.name
'    Print #n, "GUID             = " & tlInfo.GUID
'    Print #n, "Kind             = " & tlInfo.TypeKindString
'    Print #n, "AttributeMask    = " & tlInfo.AttributeMask
'
'
'    If nofAttrs = 0 Then
'        Print #n, "No attributes"
'    Else
'        For ix = LBound(attributes) To UBound(attributes)
'            Print #n, "                   " & attributes(ix)
'        Next ix
'    End If
'
'
'
'    On Error Resume Next
'    Dim interfaceCount As Long
'    interfaceCount = tlInfo.Interfaces.Count
'    On Error GoTo 0
'    Print #n, "nof Interfaces   = " & interfaceCount
'
'    Print #n, "-----------------------"
'
'
'
'    Dim mbrInfo As TLI.MemberInfo
'
'    Dim i  As Long: i = 0
'    For Each mbrInfo In tlInfo.Members           ' {
'        i = i + 1
'        Dim memberName As String
'        memberName = "<ERR>"
'        On Error Resume Next
'        memberName = mbrInfo.name
'        On Error GoTo 0
'
'        Print #n, lpad(mbrInfo.MemberId, 11) & " " & tlMemberKind(mbrInfo) & " @VTable: " & mbrInfo.VTableOffset & rpad(memberName, 40) & ": " & tlTypeName(mbrInfo.ReturnType)
'
'        Dim parInfo As TLI.ParameterInfo
'        For Each parInfo In mbrInfo.Parameters   ' {
'
'            Print #n, "   " & tlParamKind(parInfo) & " " & rpad(parInfo.name, 40) & ": " & tlTypeName(parInfo.VarTypeInfo)
'
'        Next parInfo                             ' }
'
'        '       debug.print "   " & callingConvention(mbrInfo.callConv)
'        '       if i > 20 then exit sub
'
'        Print #n, ""
'
'    Next mbrInfo
'
'    Close #n
'    Debug.Print "Finished", "See: """; ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.name & " output.txt"""
'End Sub                                          ' }
'
'Private Function tlMemberKind(mbr As TLI.MemberInfo) As String ' {
'
'    Select Case mbr.DescKind                     ' {
'        Case TLI.DESCKIND_FUNCDESC
'            Select Case mbr.INVOKEKIND
'                Case TLI.INVOKE_FUNC             ' {
'
'                    If mbr.ReturnType.VarType = VT_VOID Then
'                        tlMemberKind = "sub              "
'                    Else
'                        tlMemberKind = "function         "
'                    End If
'
'
'                Case TLI.INVOKE_PROPERTYGET: tlMemberKind = "property get     "
'                Case TLI.INVOKE_PROPERTYPUT: tlMemberKind = "property put     "
'                Case Else: tlMemberKind = "?                "
'            End Select                           ' }
'
'        Case TLI.DESCKIND_VARDESC: tlMemberKind = "variable     "
'        Case TLI.DESCKIND_NONE: tlMemberKind = "             "
'        Case Else: tlMemberKind = "?            "
'    End Select                                   ' }
'
'End Function                                     ' }
'
'Private Function tlParamKind(par As TLI.ParameterInfo) As String ' {
'
'    If par.flags And PARAMFLAG_FOPT Then         ' {
'        tlParamKind = "optional "
'    Else
'        tlParamKind = ".        "
'    End If                                       ' }
'
'    If par.flags And PARAMFLAG_FOUT Then         ' {
'        tlParamKind = tlParamKind & "byRef "
'    Else
'        tlParamKind = tlParamKind & "byVal "
'    End If                                       ' }
'
'End Function                                     ' }
'
'Private Function tlTypeName(var As TLI.VarTypeInfo) As String ' {
'    Dim vType As TliVarType
'
'    vType = var.VarType
'
'    If vType And VT_ARRAY Then
'        tlTypeName = "()"
'        vType = vType And Not VT_ARRAY
'    End If
'
'    Select Case vType
'        Case TLI.TliVarType.VT_BOOL: tlTypeName = tlTypeName & "boolean"
'        Case TLI.TliVarType.VT_BSTR: tlTypeName = tlTypeName & "string"
'        Case TLI.TliVarType.VT_CY: tlTypeName = tlTypeName & "currency"
'        Case TLI.TliVarType.VT_DATE: tlTypeName = tlTypeName & "date"
'        Case TLI.TliVarType.VT_DISPATCH: tlTypeName = tlTypeName & "object"
'        Case TLI.TliVarType.VT_HRESULT: tlTypeName = tlTypeName & "HRESULT"
'        Case TLI.TliVarType.VT_I2: tlTypeName = tlTypeName & "integer"
'        Case TLI.TliVarType.VT_I4: tlTypeName = tlTypeName & "long"
'        Case TLI.TliVarType.VT_R4: tlTypeName = tlTypeName & "single"
'        Case TLI.TliVarType.VT_R8: tlTypeName = tlTypeName & "double"
'        Case TLI.TliVarType.VT_UI1: tlTypeName = tlTypeName & "byte"
'        Case TLI.TliVarType.VT_UNKNOWN: tlTypeName = tlTypeName & "IUnknown"
'        Case TLI.TliVarType.VT_VARIANT: tlTypeName = tlTypeName & "variant"
'        Case TLI.TliVarType.VT_VOID: tlTypeName = tlTypeName & "any"
'
'        Case TliVarType.VT_EMPTY                 ' {
'
'            If Not var.TypeInfo Is Nothing Then
'                tlTypeName = tlTypeName & var.TypeInfo.name
'            End If
'
'    End Select                                   ' }
'
'End Function                                     ' }
'
'Private Function callingConvention(cc As Long) As String ' {
'
'    Select Case cc
'        Case TLI.CC_CDECL: callingConvention = "cdecl   "
'        Case TLI.CC_FASTCALL: callingConvention = "fastcall"
'        Case TLI.CC_STDCALL: callingConvention = "stdcall "
'        Case TLI.CC_SYSCALL: callingConvention = "syscall "
'        Case Else: callingConvention = "TODO: implement me!"
'    End Select
'
'End Function                                     ' }
'
'Function rpad(text As String, Length As Long, Optional padChar As String = " ") ' {
'    '
'    '   https://renenyffenegger.ch/notes/development/languages/VBA/modules/Common/Text
'    '
'    rpad = text & String(WorksheetFunction.Max(Length - Len(text), 0), padChar)
'End Function                                     ' }
'
'Function lpad(text As String, Length As Long, Optional padChar As String = " ") ' {
'    lpad = String(WorksheetFunction.Max(Length - Len(text), 0), padChar) & text
'End Function                                     ' }
'
