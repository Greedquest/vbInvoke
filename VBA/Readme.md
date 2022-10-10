# QUICKSTART:
 - Download [ModuleReflection.xlam](./ModuleReflection.xlam?raw=True), [COMTools.xlam](./COMTools.xlam?raw=True) and [MemoryTools.xlam](./MemoryTools.xlam?raw=True), and put all three in the folder `C:\ProgramData\Temp\VBAHack`
 - Download [VBAHack Demo.xlsm](./VBAHack%20Demo.xlsm?raw=True) anywhere you like (this is the file you will open)
 - Open VBAHack Demo.xlsm and press <kbd>Alt</kbd>+<kbd>F11</kbd> to view the code. Navigate to the experiments module where there is some demo code showcasing the reflection/ module accessor library.


# Detailed version:
The [VBAHack Demo.xlsm](./VBAHack%20Demo.xlsm?raw=True) contains a demo for the addin.

[ModuleReflection.xlam](./ModuleReflection.xlam?raw=True) is the actual code from the [CR post](https://codereview.stackexchange.com/questions/274532/low-level-vba-hacking-making-private-functions-public) packaged as an addin which you can add a reference to.
It requires references to the files below:

### These are the References
 
 - [MemoryTools.xlam](./MemoryTools.xlam?raw=True) - This is an addin which wraps [cristianbuse](https://github.com/cristianbuse)/**[VBA-MemoryTools](https://github.com/cristianbuse/VBA-MemoryTools)** which I'm using to read/write memory e.g. `MemByte(address As LongPtr) = value` because it is both performant and has a really nice API design in my opinion.

 - [COMTools.xlam](./COMTools.xlam?raw=True)  - This is an addin I wrote myself for this project and contains all the types and library functions to make working with COM possible in VBA. In particular:
	 - VTables** for `IUnknown`, `IDispatch` and the other various interfaces that crop up
	 - Standard methods like `ObjectFromObjPtr` and `QueryInterface` for dealing with interfaces
	 - Methods `CallFunction`, `CallCOMObjectVTableEntry` & `CallVBAFuncPtr` which wrap `DispCallFunc` and allow you to invoke function pointers
	 - _NOTE you must open this and add a reference to `MemoryTools.xlam` since it relies on that addin too_

**ALL vba projects are password protected; password = "1"**

  [3]: https://stackoverflow.com/a/42581513/6609896
  [4]: https://www.dll-files.com/tlbinf32.dll.html
