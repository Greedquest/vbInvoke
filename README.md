# vbInvoke [![Code Review](http://www.zomis.net/codereview/shield/?qid=274532&mode=score)](http://codereview.stackexchange.com/q/274532/146810) [![GitHub Latest Release)](https://img.shields.io/github/v/release/Greedquest/vbInvoke?logo=github)](https://github.com/Greedquest/vbInvoke/releases/latest)

Library for calling methods in VBA modules with several key benefits:
  - Host agnostic (unlike `Application.Run`)
  - Works with standard .bas Modules (unlike `CallByName` that calls classes)
  - Uses the standard `dot.Notation()` to call methods (unlike `AddressOf` which uses pointers invoked by `DispCallFunc`)
  - Can call public _or_ private methods
  
  
Under the hood it uses the same technique as [Rubberduck's](https://github.com/rubberduck-vba/Rubberduck) test execution engine to access the addresses of the public and private methods, ported from C# to VBA to twinBASIC (thanks RD team & Wayne for all the help). That foundation is then supplemented with a new technique to make calling them from VBA or tB much more natural (with dot notation). Click on the code review shield at top of post for more detailed code/technique walkthrough (written before tB port, but the concepts are the same)

## Quickstart
Compile an ActiveX Dll from the `vbInvoke.twinproj`, or use the precompiled `vbInvoke_win[32/64].dll` (all found in the [latest release Assets](https://github.com/Greedquest/vbInvoke/releases/latest#:~:text=Assets) - whichever matches your VBA bitness). The library has 2 main methods and 2 ways of calling them:

### ActiveX DLL (`Tools⇒Add Reference`)

```vba
Function GetStandardModuleAccessor(ByVal moduleName As String, ByVal proj As VBProject) As Object
Function GetExtendedModuleAccessor(ByVal moduleName As String, ByVal proj As VBProject, Optional ByRef outPrivateTI As vbInvoke.[_ITypeInfo]) As Object
```


_<ins>or</ins>_
### Standard DLL Declare

```vba
Declare PtrSafe Function GetStandardModuleAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject) As Object
Declare PtrSafe Function GetExtendedModuleAccessor Lib "vbInvoke_win64" (ByVal moduleName As Variant, ByVal proj As VBProject, ByRef outPrivateTI As IUnknown) As Object
```

<sub>_Note: The standard DLL version uses `Variant` for the module name and has no optional arguments. The Library name in the object browser and intellisense is `vbInvoke`_</sub>

---

These 2 functions create "Accessors": `IDispatch` Objects that can be used with dot notation `accessor.Foo` or call by name `CallByName(accessor, "Foo", ...)` to invoke
 - Public methods/functions/properties of modules (Standard Accessor)
 - Public _or Private_ methods/functions/properties (Extended Accessor)

You can also use the Extended Accessor in a For-Each loop to print all the public & private methods (although this API may change to be more useful):

```vba
Dim exampleModuleAccessor As Object
Set exampleModuleAccessor = GetExtendedModuleAccessor("ExampleModule", ThisWorkbook.VBProject)

For Each methodName In exampleModuleAccessor
	CallByName exampleModuleAccessor, methodName, vbMethod  ' Or simply exampleModuleAccessor.methodName if the method is known at design time 
Next methodName
```

---

Finally you can reference this library as a `.twinpack`. _**This library only works when compiled into in-process DLLs or VBE Addins - it cannot be used to create a standalone EXE**_ as it relies on sharing memory with the active VBProject.
