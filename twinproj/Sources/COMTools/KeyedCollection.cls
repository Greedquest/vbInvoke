VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyedCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@PredeclaredId
'@Exposed
'@Folder "Common.KeyedCollection"

'''=============================================================================
''' VBA KeyedCollection
''' ---------------------------------------------------
''' https://github.com/cristianbuse/VBA-KeyedCollection
''' ---------------------------------------------------
''' MIT License
'''
''' Copyright (c) 2017 Ion Cristian Buse
'''
''' Permission is hereby granted, free of charge, to any person obtaining a copy
''' of this software and associated documentation files (the "Software"), to
''' deal in the Software without restriction, including without limitation the
''' rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
''' sell copies of the Software, and to permit persons to whom the Software is
''' furnished to do so, subject to the following conditions:
'''
''' The above copyright notice and this permission notice shall be included in
''' all copies or substantial portions of the Software.
'''
''' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
''' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
''' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
''' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
''' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
''' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
''' IN THE SOFTWARE.
'''=============================================================================

''==============================================================================
'' Description:
''    * Keyed Collection is a combination of functionality found in a regular
''      VBA.Collection and a Scripting.Dictionary plus some extra utilities
''    * It uses 2 internal collections in order to achieve the functionality
''      (one collection for items and one for keys). The 2 are synchronized
''    * Forces the use of keys for all items and also allows the retrieval of
''      keys like a Dictionary
''    * Although Keys are of String type and are case insensitive, this Custom
''      Collection is suitable to plenty of tasks where a simple Collection
''      does not provide some of the much-needed functionality: editing keys and
''      items, retrieving keys and checking if a specific key exists
''    * The items/keys are ordered allowing the use of Before/After parameters
''      just like a Collection
''    * Removes the need for extra libraries (i.e. Microsoft Scripting Runtime)
''      in most cases and obviously works on Mac as well
''    * Items and keys can both be seen in the Locals Window, unlike Dictionary
'' Methods:
''    * Add
''         - Adds a new Key-Item pair to the internal Collections
''         - Key is mandatory! (as in Dictionary)
''         - Key is the first parameter (as in Dictionary)
''         - Before or After position can be specified (as in Collection)
''         - Keys are String type only! (as in Collection)
''    * CompareMode <Get>
''         - Returns the compare mode for keys: vbTextCompare
''         - Cannot be changed. Keys are case insensitive (as in Collection)
''    * Count <Get>
''         - Returns the number of key-item pairs
''         - Same functionality as in Collection and Dictionary
''    * Exists
''         - Checks if a specified Key exists (as in Dictionary)
''         - Enhancement over a regular Collection
''    * GetKeyAtIndex
''         - Returns the key at the specified index
''    * Item <Get>
''         - Returns an Item by key or index (as in Collection)
''         - Default Member. Can be omitted:
''           y = x.Item(v) can be called as y = x(v)
''           Set y = x.Item(v) can be called as Set y = x(v)
''    * Item <Let><Set>
''         - Changes the value of an item
''         - Default Member. Can be omitted:
''           x.Item(v) = y can be called as x(v) = y
''           Set x.Item(v) = y can be called as Set x(v) = y
''         - Enhancement over a regular Collection
''    * Items
''         - Returns a 1D array of all the items (as the Dictionary.Items)
''         - Enhancement over a regular Collection
''    * ItemsEnum
''         - Allows iteration of items with:
''           For Each itemValue In keyedCollection.ItemsEnum
''         - Avoids bug on x64. Link available in method header
''    * Key <Let>
''         - Allows a key value to be changed (as in Dictionary)
''         - If the key to be replaced does not exist then an error is raised,
''           unlike a Dictionary that adds a new Empty Item with that key
''         - Enhancement over a regular Collection
''    * KeyItemPairs
''         - Returns a 2D array of all the keys (1st col) and items (2nd col)
''         - Enhancement over a regular Collection
''    * Keys
''         - Returns a 1D array of all the keys (as the Dictionary.Keys)
''         - Enhancement over a regular Collection
''    * KeysEnum
''         - Allows iteration of keys with:
''           For Each keyValue In keyedCollection.KeysEnum
''         - Avoids bug on x64. Link available in method header
''    * Remove
''         - Removes an item by index or key (as in Collection)
''    * RemoveAll
''         - Removes all Items (as in Dictionary)
''         - Enhancement over a regular Collection
''==============================================================================

Option Explicit
Option Compare Text

'Class members
Private m_collItems As Collection
Private m_collKeys As Collection

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Initialize Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Initialize()
    Set m_collItems = New Collection
    Set m_collKeys = New Collection
End Sub

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'Class Terminate Event
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Private Sub Class_Terminate()
    Set m_collItems = Nothing
    Set m_collKeys = Nothing
End Sub

'*******************************************************************************
'Adds a Value or Object to Collection
'Raises Error:
'   -   5: both 'beforeIndex' and 'afterIndex' were specified
'          or one of them is an invalid key
'          or one of them is an invalid index (while collection has no elements)
'   -   9: 'beforeIndex' or 'afterIndex' invalid (while collection has elements)
'   -  13: invalid data type for 'beforeIndex' or 'afterIndex'
'   - 457: a duplicated key was specified
'*******************************************************************************
Public Sub Add(ByVal keyValue As String _
             , ByVal itemValue As Variant _
             , Optional ByVal beforeIndex As Variant _
             , Optional ByVal afterIndex As Variant)
    m_collItems.Add itemValue, keyValue, beforeIndex, afterIndex
    m_collKeys.Add keyValue, keyValue, beforeIndex, afterIndex
End Sub

'===============================================================================
'Compare mode property (Get only). Text Compare only!
'===============================================================================
Public Property Get CompareMode() As Long
    CompareMode = VBA.VbCompareMethod.vbTextCompare
End Property

'===============================================================================
'Get the number of Items/Keys
'===============================================================================
Public Property Get Count() As Long
    Count = m_collItems.Count
End Property

'*******************************************************************************
'Checks if a Key exists. Does not raise errors
'Note that keys are case insensitive
'*******************************************************************************
Public Function Exists(ByVal keyValue As String) As Boolean
    On Error Resume Next
    m_collKeys.Item keyValue
    Exists = (Err.Number = 0)
    On Error GoTo 0
End Function

'===============================================================================
'Get Item by index or key
'Raises Error:
'   -   5: invalid key (key is not associated with any element)
'          or invalid index (while collection has no elements)
'   -   9: invalid index (while collection has elements)
'   -  13: invalid data type for 'indexOrKey'
'Default Class Member. KeyedCollection.Item(i) same as KeyedCollection(i)
'===============================================================================
'@DefaultMember
Public Property Get Item(ByVal indexOrKey As Variant) As Variant
Attribute Item.VB_UserMemId = 0
    LetSet(Item) = m_collItems(indexOrKey) 'Avoids extra check for vbDataObject
End Property
Private Property Let LetSet(ByRef result As Variant, ByVal v As Variant)
    If IsObject(v) Then Set result = v Else result = v
End Property
'===============================================================================
'Let Item by index or key
'Raises Error:
'   -    5: invalid key (key is not associated with any element)
'           or invalid index (while collection has no elements)
'   -    9: invalid index (while collection has elements)
'   -   13: invalid data type for 'indexOrKey'
'   -  450: 'Set' is missing when assigning an object
'Note that only the Items Collection is updated
'===============================================================================
'@DefaultMember
Public Property Let Item(ByVal indexOrKey As Variant, ByVal itemValue As Variant)
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(itemValue) Then
        Err.Raise 450, TypeName(Me) & ".Item", "Invalid property assignment"
    End If
    UpdateItem indexOrKey, itemValue
End Property
'===============================================================================
'Set Item by index or key
'Raises Error:
'   -   5: invalid key (key is not associated with any element)
'          or invalid index (while collection has no elements)
'   -   9: invalid index (while collection has elements)
'   -  13: invalid data type for 'indexOrKey'
'Note that only the Items Collection is updated
'===============================================================================
'@DefaultMember
Public Property Set Item(ByVal indexOrKey As Variant, ByVal itemValue As Object)
Attribute Item.VB_UserMemId = 0
    UpdateItem indexOrKey, itemValue
End Property

'*******************************************************************************
'Utility for Item property <Let> and <Set>
'*******************************************************************************
Private Sub UpdateItem(ByVal indexOrKey As Variant, ByVal itemValue As Variant)
    Dim keyValue As String
    Dim tempKey As String
    '
    'Get Key
    keyValue = m_collKeys(indexOrKey)
    'Generate random temporary key
    tempKey = RandomKey()
    'Add tempKey after keyValue to preserve the position
    m_collItems.Add Empty, tempKey, , keyValue
    'Remove old item
    m_collItems.Remove keyValue
    'Insert new item after the tempKey
    m_collItems.Add itemValue, keyValue, , tempKey
    'Remove tempKey
    m_collItems.Remove tempKey
End Sub

'*******************************************************************************
'Utility. Generate a random key that is not part of the Collection
'*******************************************************************************
Private Function RandomKey() As String
    Dim tempKey As String
    '
    Do
        tempKey = CStr(VBA.Rnd())
    Loop Until Not Me.Exists(tempKey)
    RandomKey = tempKey
End Function

'*******************************************************************************
'Returns a 1-Dimensional Array of Items
'*******************************************************************************
Public Function Items() As Variant()
    If m_collItems.Count = 0 Then
        Items = Array()
        Exit Function
    End If
    '
    Dim arr() As Variant: ReDim arr(1 To m_collItems.Count)
    Dim v As Variant
    Dim i As Long: i = 1
    '
    'A Variant variable of type VbVarType.vbDataObject does NOT need 'Set' when
    '   assigned from one Variant variable to another
    For Each v In m_collItems
        If VBA.IsObject(v) Then Set arr(i) = v Else arr(i) = v
        i = i + 1
    Next v
    '
    Items = arr
End Function

'*******************************************************************************
'Provides For Each... support for Items while avoiding bug on x64:
'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
'*******************************************************************************
Public Function ItemsEnum() As EnumHelper
    With New EnumHelper
        Set .EnumVariant = m_collItems.[_NewEnum]
        Set ItemsEnum = .self
    End With
End Function

'*******************************************************************************
'Returns a key by it's index position
'Raises Error:
'   -   5: invalid index (while collection has no elements)
'   -   9: invalid index (while collection has elements)
'*******************************************************************************
Public Function GetKeyAtIndex(ByVal indexValue As Long) As String
    GetKeyAtIndex = m_collKeys(indexValue)
End Function

'===============================================================================
'Change existing key value
'Raises Error:
'   -   5: invalid oldKey (key is not associated with any element)
'   - 457: invalid newKey (key is already associated with an element)
'===============================================================================
Public Property Let Key(ByVal oldKey As String, ByVal newKey As String)
    If oldKey = newKey Then Exit Property
    '
    'Add newKey after the oldKey to preserve the position
    m_collItems.Add m_collItems.Item(oldKey), newKey, , oldKey
    m_collKeys.Add newKey, newKey, , oldKey
    '
    'Remove oldKey
    m_collItems.Remove oldKey
    m_collKeys.Remove oldKey
End Property

'*******************************************************************************
'Returns a 1-Dimensional Array of String Keys
'Variant() is returned instead of String() for flexibility of use
'*******************************************************************************
Public Function Keys() As Variant()
    If m_collKeys.Count = 0 Then
        Keys = Array()
        Exit Function
    End If
    '
    Dim arr() As Variant: ReDim arr(1 To m_collKeys.Count)
    Dim v As Variant
    Dim i As Long: i = 1
    '
    For Each v In m_collKeys
        arr(i) = v
        i = i + 1
    Next v
    '
    Keys = arr
End Function

'*******************************************************************************
'Provides For Each... support for Keys while avoiding bug on x64:
'https://stackoverflow.com/questions/63848617/bug-with-for-each-enumeration-on-x64-custom-classes
'*******************************************************************************
Public Function KeysEnum() As EnumHelper
    With New EnumHelper
        Set .EnumVariant = m_collKeys.[_NewEnum]
        Set KeysEnum = .self
    End With
End Function

'*******************************************************************************
'Returns a 2-Dimensional Array with 2 columns for String Keys and Items
'Note that using 2 For Each... loops is about 15x faster than using just one on
'   the Keys collection and then retreiving each item by key from the Items
'   collection. The latter would also require an extra check for vbDataObject
'   for the items or the use of LetSet
'*******************************************************************************
Public Function KeyItemPairs() As Variant()
    If m_collKeys.Count = 0 Then
        KeyItemPairs = Array()
        Exit Function
    End If
    '
    Dim arrPairs() As Variant
    Dim i As Long
    Dim v As Variant
    '
    ReDim arrPairs(1 To m_collKeys.Count, 1 To 2)
    i = 1
    For Each v In m_collKeys
        arrPairs(i, 1) = v
        i = i + 1
    Next v
    i = 1
    For Each v In m_collItems
        If VBA.IsObject(v) Then Set arrPairs(i, 2) = v Else arrPairs(i, 2) = v
        i = i + 1
    Next v
    '
    KeyItemPairs = arrPairs
End Function

'*******************************************************************************
'Remove item
'Raises Error:
'   -   5: invalid key (key is not associated with any element)
'          or invalid index (while collection has no elements)
'   -   9: invalid index (while collection has elements)
'   -  13: invalid data type for 'indexOrKey'
'*******************************************************************************
Public Sub Remove(ByVal indexOrKey As Variant)
    m_collItems.Remove indexOrKey
    m_collKeys.Remove indexOrKey
End Sub

'*******************************************************************************
'Remove all items
'*******************************************************************************
Public Sub RemoveAll()
    If m_collItems.Count > 0 Then
        Set m_collItems = New Collection
        Set m_collKeys = New Collection
    End If
End Sub

'*******************************************************************************
'Get self instance. Useful in 'With New KeyedCollection' blocks
'*******************************************************************************
Public Function self() As KeyedCollection
    Set self = Me
End Function

Public Function Create() As KeyedCollection
    Set Create = New KeyedCollection
End Function
