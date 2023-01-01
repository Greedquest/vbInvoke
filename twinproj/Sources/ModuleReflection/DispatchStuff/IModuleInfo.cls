VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IModuleInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "TypeInfoInvoker.DispatchWrapper"
'@Exposed
Option Explicit

'@Description("The Extended ITypeInfo interface for the module this accessor refers to")
Public Property Get ExtendedITypeInfo() As IUnknown
Attribute ExtendedITypeInfo.VB_Description = "The Extended ITypeInfo interface for the module this accessor refers to"
End Property

'@Description("Parsed map of ProcedureName:Info for methods of the extended module accessor (public/private)")
Public Property Get ModuleFuncInfoMap() As Dictionary
Attribute ModuleFuncInfoMap.VB_Description = "Parsed map of ProcedureName:Info for methods of the extended module accessor (public/private)"
End Property

'@Description("Base accessor for accessing public members of a module using standard late binding")
Public Property Get PublicOnlyModuleAccessor() As Object
Attribute PublicOnlyModuleAccessor.VB_Description = "Base accessor for accessing public members of a module using standard late binding"
End Property

'@Description("Rich accessor for accessing public and private members of any module based on extended ITypeInfo")
Public Property Get ExtendedModuleAccessor() As Object
Attribute ExtendedModuleAccessor.VB_Description = "Rich accessor for accessing public and private members of any module based on extended ITypeInfo"
End Property
