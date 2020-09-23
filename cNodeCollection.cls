VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNodeCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Full Version  Full Version  Full Version  Full Version
' ********************************************************
'       FULL VERSION NODE-BASED COLLECTION CLASS
' ********************************************************
' Full Version  Full Version  Full Version  Full Version


' DESCRIPTION:  A node-based, parent-child, collection class.
'   This class combines treeview-like collections with VB's collection & dictionary objects. Collection items have heirarchy.
'   Can be used without hierarchy by simply not provding a Parent/Sibling Key when adding items


' CHANGE HISTORY
'--------------------------------------------------------------------------------------------------------------
'  Updated the MethodsProps.rtf file, as needed, with every change
' 17 Dec 06. Added support to for saving persistent objects in SaveCollection
'   - persistent objects include stdPictures and stdFonts
' 15 Dec 06. Finalizing project.
'   - Found in specific scenarios, a deserialized object's reference count was one more than it should be. Corrected
'   - When creating arrays, array features are now determined and added to the array attributes for all array types
'   - Variants containing UDTs (user-defined types) are not persisted when saving collection
'       -- UDTs can be saved to a Variant only if the UDT is public in a class within a compiled reference (DLL, OCX, etc)
'   - Variants returning a VarType of vbError or vbDataObject are not persisted when saving collection
'       -- Haven't been able to break those down at this point. IUnknown objects appear, from testing, to be vbDataObjects
' 14 Dec 06.
'   - Completely revamped the serialization and deserialization routines
'     (See GetSerializedItem & SetSerializedItem)
'       -- Now all array types, including variant arrays, are supported
'       -- Arrays are no longer restricted to number of dimensions
'       -- Objects (i.e., Classes) within arrays still need to be serialized by user
'       -- Parameters in Public Event SerializeObject has been changed
'   - Added tons of remarks to the serialization & deserialization routines
' 10 Dec 06.
'   - Playing more with manipulating multidimensional arrays via pointers.
'       -- Removed logic that wrapped the collection item in a variant. This actually
'          forces VB to copy the array and we don't need a copy;
'          we can work directly off the collection item vs a variant wrapper
' 9 Dec 06.
'   - Added SiblingOffsetIndex method
'   - Changed SortCollection to return True if successful
'   - Changed SaveCollection & LoadCollection to return True if successful
'   - Upgraded SaveCollection to auto-save multidimensional arrays of all data types
'       except Variants and Objects/UDTs (those still require SerializeObject event)
'       -- With above exceptions, any multidimensional array containing more than
'          10 dimensions will no longer be able to be saved/persisted
'       -- Resulted in permanent removal of Public Events SerializeStringArray & SerializeArray
'   - Changed ClearKeys method to clear individual branch levels or entire collection
'   - Renamed CountCollection to ItemCount
'   - Renamed CountChildren to ChildCount
'   - Removed IsSiblingOf. Calling myCol.Parent(Index)=myCol.Parent(OtherIndex) results in same thing
'   - Removed CountSiblings. Calling myCol.ChildCount(ParentIndex) results in same thing
'       -- when ParentIndex is zero or vbNullString, root level nodes are counted
' 8 Dec 06.
'   - Found minor bugs in the ShellSortString routine. Fixed
'   - Added optional parameter in MoveItem to allow a new key to be assigned along with the move
'   - Merged the AddKey routine into the FindKey routine; no need to keep it separate
' 7 Dec 06.
'   - the FamilyTreeStructure, when used, was being dimensioned with 3 elements; only 2 are needed. Changed
'   - Added SortCollection method
'       -- Required new supporting routines: ShellSortNumbers, ShellSortString, SortRecursiveNodes, SortNodeLevel
'   - Added additional Relationship enumerations to allow sorting when a child node is added or moved
'       -- Required new supporting routine: GetChildInsertPosition
'   - Added IndexFromKeyIndex method
'   - Changed Parent function to a read/write property
'   - Removed obsolete enumerations
' 6 Dec 06 - Initial release
'--------------------------------------------------------------------------------------------------------------

' Highlights include....
' ============================================================================================================
' 1) Key or key-less items. Keys can be numeric or not, case-sensitive or not. vbNullString is reserved and cannot be used.
' 2) TreeView-Like methods:  AddItem(item,key,parent or sibling), First/Next/PreviousSibling, First/LastChild, etc
' 3) Assigned keys can be changed as desired, but still must be unique if used.
' 4) Nodes, and any children, are moved/deleted in one call
' 5) Provides nice-to-have functions like KeyExists, KeyToIndex, HasChildren, MoveItem (from one node to another) & more
' 6) Change contents of an item without having to delete/re-add
' 7) Fast navigation via doubly linked lists
' 8) Value-based collection; therefore, instantiated objects (i.e., Classes) added to collection are kept alive until removed from collection
' 9) Has ability to save/load collection with restrictions. If you use this ability, declare this class WithEvents
'     See SaveCollection, LoadCollection, and the Public Event
' 10) Now has sorting ability too

' SPECIAL CONSIDERATIONS.....
' ============================================================================================================
' 1) Collection heirarchy are linked lists. Indexes of items should not be cached by you.
'    Indexes change dynamically when items are deleted or a new collection is loaded
' 2) Above being said, keys should be used. If keys aren't used, then you should use the
'    following methods to obtain indexes or to loop thru the collection:
'    - FirstChild, LastChild returns relative indexes based on a passed parent node index/key
'    - FirstSibling, LastSibling returns relative indexes based on passed sibling node index/key
'    - NextSibling, PreviousSibling returns relative indexes based on passed sibling/child node index/key
'    - SiblingOffsetIndex returns sibling x positions from the first or last sibling
'       Whenever a Child or Sibling index method returns zero, you are at the beginning or end of the branch
' 3) Speaking of keys. Since numeric keys can be used too, all methods, that allow
'    a Variant Key/Index, have no way of knowing if "1234" is a numeric key or an index.
'    The type of passed variable determines whether the class treats the parameter as a
'    key or index. VarType of String are ALWAYS treated as Keys, whereas,
'    VarType of Byte,Integer,Long,Single,Double,Currency are ALWAYS treated as Indexes
'    Example:
'       Dim strIndex As String, lngIndex As Long
'       Dim strKey As String, lngKey As Long
'       strIndex="99": lngIndex=99
'       strKey="-123899": lngIndex=-123899
'    Getting Item by Index:  Either Class.Item(lngIndex) or Class.Item(CLng(strIndex))
'    Getting Item by Key: Either Class.Item(strKey) or Class.Item(CStr(lngKey))
' 4) Sorting
'    - The collection can now be sorted. See SortCollection method for restrictions in what can and cannot be sorted.
'    - Another option is to assign keys which are sorted (remember keys can be numeric or not), you can request an
'       item by its key index. Keys are always sorted ascending order. To reverse sort, get the item's key index
'       in descending order.  See IndexFromKeyIndex method.
'    - Yet another option is to dictate where an item is to be added to a collection by properly identifying
'       which Sibling the item should be added before or after.
'   Once items are set in their positions, their positions can only be changed using 4 methods. Even if an item's
'   index changes due to collection item deletes, the remaining items' positions relative to their siblings do not
'   change. The 4 methods that change positions of items are AddItem, MoveItem, MoveChildren, SortCollection

' RESTRICTIONS/LIMITATIONS
' ============================================================================================================
' 1) The LoadCollection, SaveCollection, SaveProperties, GetSerializedItem & the Public Events should be removed for VB5 users.
'    - VB5 does not support creating property bags. i.e., Set ppBag=New PropertyBag is an error in VB5
'    - Should you still want to replicate the LoadCollection/SaveCollection functions, you will have to write your own.
'       Suggest possibly using something similiar to an INI file format and caching the bytes of that file?
'    - Likewise, the example below for SerializeObject, of serailizing a class, will need to be reworked/rethought for VB5 users.
' 2) Sorting on items containing primarily objects can only be successful 100% of the time when all of the objects
'    have a common public property whose value can be compared/sorted. Otherwise, you will need to sort the objects yourself

' PUBLIC EVENTS
' ============================================================================================================
' Should you want this class to cache the collection, it can, except for objects; is that true?....
' NOTE: This class doesn't save anything, it organizes the entire collection into an array
' that is passed via the SaveCollection method. That array can be written by you to a user file or,
' if this is used within a usercontrol, to the property bag (yes, property bags can store byte arrays,
' but not arrays of other data types). That same array must not be altered thereafter and then can be
' passed to the LoadCollection method to re-create the collection. Interested?  Read just a little bit more.

' When your collection contains objects (i.e., Classes), you will need a way of serializing your class
' into a byte array and pass the byte array as a response to the SerializeObject Event when it is
' called.  This event is called from SaveCollection method to get your serialized object and
' also from LoadCollection to pass the serialized array back to you for deserializing

' You will get this event called for each collection item or arrayed item of a collection item that contains an Object
Public Event SerializeObject(ByVal Serialize As Boolean, ByRef collectionObject As Object, ByRef DataArray() As Byte, ByRef ObjectID As String)

'   If Serialize=True :: You are to serialize the passed object into the DataArray()
'       collectionObject. A non-Nothing object within your collection.
'           -- Test the type object, if needed, by using TypeOf(collectionObject) Is classWhatever
'       DataArray(). An empty array you are to fill with the contents of the serailized object
'           -- The serialized object in DataArray can be multidimensional & any LBound.
'       ObjectID (Optional) :: a name you provide that uniquely distinguishes the kind of object it is. This is usually a Class Type.
'           -- When you get the array back for deserialization, you may or many not included flags to indicate which class/object
'               type it was created from (i.e., "cMyClass"). Therefore knowing which class to create from array may be impossible

'   If Serialize=False :: You are to deserialize the DataArray() into the passed collectionObject
'       collectionObject. An empty object to be set from the deserailized array
'           -- i.e, deserialize DataArray into a new classWhatever, then Set collectionObject=classWhatever
'       DataArray(). A populated array you are to use to create a new object
'           -- The DataArray will be same size, dimensions, and have same LBound/UBound you passed when Serialize=True
'       ObjectID (Optional) :: If the ObjectID was provided when Serialize=True, the the value you provided else vbNullString

' How do you serialize an object? For a class object, it is rather simple. This is just one possible way.
' - Create a Friend Function SaveClass(ClassBytes() As Byte) As Boolean
    ' - Within that function Dim myData As New PropertyBag
    ' - Just like an OCX, write your class values into the property bag
    ' - Pass the property bag like so:  ClassBytes() = myData.Contents()
' - Create a Friend Function LoadClass(ClassBytes() As Byte) As Boolean
    ' - Within that function Dim myData As New PropertyBag
    ' - Set the propety bag contents: myData.Contents() = ClassBytes(). Use On Error statement(s)
    ' - Just like an OCX, read your class values from the property bag



' TYPE DECLARATIONS, ENUMERATIONS AND CONSTANTS
' ============================================================================================================
' only API used; think this is a first for me--just one API ;)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Type KeyStructure
    Index As Long                   ' index into the m_Col() array
    Key As String                   ' the key
End Type
Private Type FamilyTreeStructure
    Tree() As Long                  ' For parent nodes only, array contains First Child, Last Child indexes for fast insert/delete/enumeration
End Type
Private Type CollectStructure       ' contains relationships between node siblings/children
    Parent As Long                  ' index value of the parent node, if zero then this is a main level node
    Elder As Long                   ' if zero then this is the last sibling, else index of next sibling (linked list)
    Younger As Long                 ' if zero then this is the first sibling, else index of previous sibling (reverse linked list)
    Key As Long                     ' if zero then item is not keyed, else StrPtr of the Key: See VBStringFromBSTRpointer function
    Family As FamilyTreeStructure   ' See Structure above ^^
    Item As Variant                 ' your custom collection data item
End Type
    ' ^^ 20 bytes overhead/admin per item + 16 for the Variant + (if keyed add 8 bytes & if has children add another 8 bytes)
    
Private Enum eTreeMembers            ' internal use only
    treefirst = 1
    treelast = 2
End Enum
Public Enum eNodeRelationship   ' see AddItem for description of meaning
    relFirstSibling = 0
    relLastSibling = 1
    relNextSibling = 2
    relPrevSibling = 3
    relFirstChild = 4
    relLastChild = 5
    relChildSortedAsText_Asc = 6
    relChildSortedBinaryText_Asc = 8
    relChildSortedAsNumeric_Asc = 10
    relChildSortedAsText_Desc = 7
    relChildSortedBinaryText_Desc = 9
    relChildSortedAsNumeric_Desc = 11
End Enum
Public Enum eCompareEnum
    keyCaseSensitive = 0
    keyCaseInsensitive = 1
End Enum
Public Enum eSortLevel          ' see SortCollection & ClearKeys for description of meaning
    brchAllLevels = 0
    brchSiblings = 1
    brchChildren = 2
End Enum
Public Enum eSortOptions        ' see SortCollection for description of meaning
    srtTextText = 0
    srtTextBinary = 1
    srtAsNumeric = 2
End Enum

' possible return errors when adding/editing/removing items
Public Enum ErrorEnum
    ERROR_NOT_FOUND = &H80000500
    ERROR_KEY_DUPLICATED = &H80000501
    ERROR_INVALID_KEYINDEX = &H80000502
    ERROR_INVALID_RELATIVEKEY = &H80000503
    ERROR_INVALID_DATATYPE = &H80000504
    ERROR_INVALID_PROPERTY = &H80000505
    ERROR_INVALID_COLLECTION = &H80000506
    ERROR_MAXIMUM_ITEMS_REACHED = &H80000507
End Enum

Private Const MAXINT_4 As Long = 2147483647 ' max number of collection items
                                            ' also used in SafeOffset routine
' class wide variables
' ============================================================================================================
Private m_CompareType As eCompareEnum   ' key comparision; cannot be changed after 1st key is added to collection. Makes keys case sensitive or insensitive
Private m_Buffer As Long                ' buffer of n items (default is 10) to help prevent Redimensioning collection after each add
Private m_ItemCount As Long             ' number of items in the entire collection
Private m_KeyCount As Long              ' number of keys (items do not require keys)

Private m_Keys() As KeyStructure        ' array of keys used
Private m_Col() As CollectStructure     ' the collection
Private m_Objects() As Object           ' may be used temporarily when loading a previously saved object. Referenced in LoadColllection, SetSerializedItem

Public Function AddItem(Item As Variant, Optional ByVal Key As String = vbNullString, _
            Optional RelativeKeyIndex As Variant = vbNullString, _
            Optional ByVal Relationship As eNodeRelationship = relNextSibling) As Long

    ' PURPOSE: Add/append/insert an item into the collection
    
    ' Item.{Required} Any VB number, string, control, variant, object, etc
    ' Key.{Optional} A key to reference this new item by. The key must be unique, can be numeric or not. Cannot be Null
    ' RelativeKeyIndex.{Optional} The key or index to the node that the Relationship parameter applies too
    '       -- If is vbNullString or zero, then the root will be used
    ' Relationship.{Optional}. Default is relNextSibling
    '       -- relFirstChild. The new node will become the first child of the node ref'd by RelativeKeyIndex
    '       -- relLastChild. The new node will become the last child of the node ref'd by RelativeKeyIndex
    '       -- relFirstSibling. The new node will become the first child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relLastSibling. The new node will become the last child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relNextSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its next sibling
    '       -- relPrevSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its previous sibling
    '   The following do not apply for items containing arrays or objects and will reset to relLastChild if that case
    '       -- relChildSortedAsNumeric_Asc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsNumeric_Desc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsText_Asc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedAsText_Desc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedBinaryText_Asc. If item is String, place child in sorted position comparing siblings case-sensitive
    '       -- relChildSortedBinaryText_Desc. If item is String, place child in sorted position comparing siblings case-sensitive
    ' Return value. If no errors encounterd, return value is current Index of newly added item

    ' Note about Overflow checks in AddItem, DeleteNode and FindKey(adding a key)...
    ' IMO they are probably useless. If the collection got so big as to contain over
    ' a million items, user would probably run out of memory first. But we make the
    ' checks anyway as a measure of professionalism.
    
    Dim newIndex As Long, KeyIndex As Long, relIndex As Long, pIndex As Long
    Dim vType As Integer
    
    If Not RelativeKeyIndex = vbNullString Then ' validate relative's key/index exists
        relIndex = ConvertKeyToIndex(RelativeKeyIndex, True)
        If relIndex < 0 Then
            Err.Raise ERROR_INVALID_RELATIVEKEY, "Collection.AddItem", "Relative Key is Invalid"
        End If
    End If
    
    ' determine position of the new item
    If m_ItemCount = 0 Then
        newIndex = 1
    ElseIf m_ItemCount = MAXINT_4 Then
        ' what are the odds; better off checking anyway
        Err.Raise ERROR_MAXIMUM_ITEMS_REACHED, "Collection.AddItem", "The Collection Is Full"
    Else
        newIndex = m_ItemCount + 1
    End If
    
    If Not Key = vbNullString Then          ' validate the passed key is unique
        KeyIndex = FindKey(Key, newIndex)
        If KeyIndex < 0 Then
            Err.Raise ERROR_KEY_DUPLICATED, "Collection.AddItem", "Key Already Exists in this Collection"
        End If
    End If
    
    ' add the item to our m_Col() array
    If newIndex = 1 Then
        ReDim m_Col(0 To newIndex + m_Buffer)
    ElseIf UBound(m_Col) = m_ItemCount Then
        If newIndex >= MAXINT_4 - m_Buffer Then ' over 2.4 trillion items?
            ReDim Preserve m_Col(0 To MAXINT_4)
        Else
            ReDim Preserve m_Col(0 To newIndex + m_Buffer)
        End If
    End If
    
    m_ItemCount = m_ItemCount + 1
    ' the Key member are pointers; this allows caching key as 4 bytes vs Len(Key)*2+4 bytes (BSTR String)
    ' Also prevents reindexing m_Keys() array when it is added to or deleted from
    m_Col(newIndex).Key = StrPtr(m_Keys(KeyIndex).Key) ' requires a little work getting key. See VBStringFromBSTRpointer
    
    ' get the true varType.
    ' For example. VarType on a stdPicture returns vbLong not vbObject
    CopyMemory vType, ByVal VarPtr(Item), 2&
    vType = vType And Not &H4000    ' objects can be ByRef
    
    ' when not an object, use VB's VarType function
    If Not vType = vbObject Then vType = VarType(Item)
    Select Case vType
    Case vbObject, vbDataObject, vbError
        Set m_Col(newIndex).Item = Item
    Case Else
        m_Col(newIndex).Item = Item
    End Select
    
    If Relationship > relLastChild And Relationship < relChildSortedAsNumeric_Desc + 1 Then
        ' place in a sorted position.... determine the position among its other siblings
        pIndex = GetChildInsertPosition(newIndex, relIndex, (Relationship - 6 + (Relationship Mod 2 = 1)), (Relationship Mod 2 = 0))
        If pIndex = 0 Then ' this will be the last child
            Relationship = relLastChild
        Else    ' inserted before
            relIndex = pIndex                   ' cache the existing node's index
            Relationship = relPrevSibling       ' insert beore
            pIndex = 0
        End If
    End If
    
    ' add parent and sibling links as needed
    Select Case Relationship
    
        Case relFirstChild   ' this node will be 1st child of relIndex
            If IsArrayEmpty(Not m_Col(relIndex).Family.Tree) Then
                InitializeTree relIndex, newIndex
            Else
                With m_Col(relIndex).Family
                    m_Col(newIndex).Elder = .Tree(treefirst)    ' fwd index
                    m_Col(.Tree(treefirst)).Younger = newIndex  ' rev index
                    .Tree(treefirst) = newIndex
                End With
                m_Col(newIndex).Parent = relIndex
            End If
            
        Case relLastChild ' this node will be last child of relIndex
            If IsArrayEmpty(Not m_Col(relIndex).Family.Tree) Then
                InitializeTree relIndex, newIndex
            Else
                With m_Col(relIndex).Family
                    m_Col(.Tree(treelast)).Elder = newIndex     ' fwd index
                    m_Col(newIndex).Younger = .Tree(treelast)   ' rev index
                    .Tree(treelast) = newIndex
                End With
                m_Col(newIndex).Parent = relIndex
            End If
            
        Case relFirstSibling  ' (Insert At Beginning)
            pIndex = m_Col(relIndex).Parent
            If IsArrayEmpty(Not m_Col(pIndex).Family.Tree) Then
                InitializeTree pIndex, newIndex
            Else
                With m_Col(pIndex).Family
                    m_Col(newIndex).Elder = .Tree(treefirst)    ' fwd index
                    m_Col(.Tree(treefirst)).Younger = newIndex  ' rev index
                    .Tree(treefirst) = newIndex
                End With
                m_Col(newIndex).Parent = pIndex
            End If
            
        Case relLastSibling  ' (Append To End)
            pIndex = m_Col(relIndex).Parent
            If IsArrayEmpty(Not m_Col(pIndex).Family.Tree) Then
                InitializeTree pIndex, newIndex
            Else
                With m_Col(pIndex).Family
                    m_Col(.Tree(treelast)).Elder = newIndex     ' fwd index
                    m_Col(newIndex).Younger = .Tree(treelast)   ' rev index
                    .Tree(treelast) = newIndex
                End With
                m_Col(newIndex).Parent = pIndex
            End If
            
        Case relPrevSibling ' (InsertBefore)
            If relIndex = 0 Then    ' core root (Default Parameter) cannot have siblings. Use its children
                If IsArrayEmpty(Not m_Col(relIndex).Family.Tree) Then
                    InitializeTree relIndex, newIndex   ' collection not yet started, initialize it
                Else
                    relIndex = m_Col(relIndex).Family.Tree(treefirst) ' set current sibling to first sibling
                End If
            Else
                pIndex = m_Col(relIndex).Parent         ' user gave a non-root node, use it
            End If
            If Not relIndex < 1 Then
                ' update child links as needed
                If m_Col(pIndex).Family.Tree(treefirst) = relIndex Then m_Col(pIndex).Family.Tree(treefirst) = newIndex
                ' inserting node, set new node's linked list members
                m_Col(newIndex).Elder = relIndex                    ' fwd index
                m_Col(newIndex).Younger = m_Col(relIndex).Younger   ' rev index
                ' adjust moved node's linked list members
                If Not m_Col(relIndex).Younger = 0 Then ' not first sibling
                    ' make moved nodes younger sibling point to new node as its older sibling
                    m_Col(m_Col(relIndex).Younger).Elder = newIndex ' fwd index
                End If
                ' make moved node's younger sibling point to new node
                m_Col(relIndex).Younger = newIndex                  ' rev index
                m_Col(newIndex).Parent = pIndex
            End If
            
        Case Else   'relNextSibling ' (InsertAfter)
            If relIndex = 0 Then    ' core root (Default Parameter) cannot have siblings. Use its children
                If IsArrayEmpty(Not m_Col(relIndex).Family.Tree) Then
                    InitializeTree relIndex, newIndex   ' collection not yet started, initialize it
                Else
                    relIndex = m_Col(relIndex).Family.Tree(treelast) ' set current sibling to first sibling
                End If
            Else
                pIndex = m_Col(relIndex).Parent         ' user gave a non-root node, use it
            End If
            If Not relIndex < 1 Then
                ' update child stats as needed
                If m_Col(pIndex).Family.Tree(treelast) = relIndex Then m_Col(pIndex).Family.Tree(treelast) = newIndex
                ' inserting node, set new node's linked list members
                m_Col(newIndex).Elder = m_Col(relIndex).Elder       ' fwd index
                m_Col(newIndex).Younger = relIndex                  ' rev index
                ' adjust moved node's linked list members
                If Not m_Col(relIndex).Elder = 0 Then ' not last sibling
                    ' make moved nodes elder sibling point to new node as its younger sibling
                    m_Col(m_Col(relIndex).Elder).Younger = newIndex ' rev index
                End If
                ' make moved node's elder sibling point to new node
                m_Col(relIndex).Elder = newIndex                 ' fwd index
                m_Col(newIndex).Parent = pIndex
            End If
        
    End Select
    AddItem = newIndex
    
End Function

Public Function RemoveItem(KeyIndex As Variant) As Boolean

    ' PURPOSE: Remove a node from the collection. Any child nodes will also be removed
    
    ' KeyIndex. A key string or numerical Index to reference the item by
    ' Key cannot be vbNullString and the Index cannot be < 1
    
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 1 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.RemoveItem", "Invalid Key or Index"
    Else
        RemoveItem = DeleteNode(Index)
    End If

End Function

Public Function RemoveChildren(ParentKeyIndex As Variant) As Boolean

    ' PURPOSE: Remove all child nodes, and their children, from ParentKeyIndex
    
    ' ParentKeyIndex. A key or Index to reference the item by
    ' - Passing zero or vbNullString will result in entire collection being removed
    
    Dim Index As Long, sKey As String, kIndex As Long
    Index = ConvertKeyToIndex(ParentKeyIndex, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.RemoveItem", "Invalid Key or Index"
    ElseIf Index = 0 Then
        Call ClearCollection
    Else
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
            ' when deleting nodes, a 2 node swap is actually used...
            ' This can cause an index to change during recursion in DeleteNode method
            ' So we use a key so we can easily identify the node index
            If m_Col(Index).Key = 0 Then    ' no key; create one
                sKey = RandomKey
                kIndex = FindKey(sKey, Index)
                m_Col(Index).Key = StrPtr(m_Keys(kIndex).Key)
            Else                            ' use existing key
                sKey = VBStringFromBSTRpointer(m_Col(Index).Key)
            End If
            Do Until IsArrayEmpty(Not m_Col(Index).Family.Tree) = True
                DeleteNode m_Col(Index).Family.Tree(treelast)
                Index = m_Keys(FindKey(sKey)).Index
            Loop
            If Not kIndex = 0 Then
                ' we created a key for this node; remove it now
                DeleteKey , sKey
                m_Col(Index).Key = 0
            End If
        End If
    End If
    RemoveChildren = True
    
End Function

Public Function MoveItem(KeyIndex As Variant, TargetKeyIndex As Variant, _
                Optional ByVal Relationship As eNodeRelationship = relNextSibling, _
                Optional NewKey As String) As Boolean

    ' PURPOSE: Move a node from current position to a new position and possibly a new parent.
    
    ' KeyIndex. A key or Index to reference the item to move. Cannot be zero or vbNullString
    ' TargetKeyIndex. The key or index to the node that the Relationship parameter applies too
    '       -- If is vbNullString or zero, then the root will be used
    ' Relationship.{Optional}. Default is relNextSibling
    '       -- relFirstChild. The new node will become the first child of the node ref'd by RelativeKeyIndex
    '       -- relLastChild. The new node will become the last child of the node ref'd by RelativeKeyIndex
    '       -- relFirstSibling. The new node will become the first child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relLastSibling. The new node will become the last child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relNextSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its next sibling
    '       -- relPrevSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its previous sibling
    '   The following do not apply for items containing arrays or objects and will reset to relLastChild if that case
    '       -- relChildSortedAsNumeric_Asc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsNumeric_Desc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsText_Asc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedAsText_Desc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedBinaryText_Asc. If item is String, place child in sorted position comparing siblings case-sensitive
    '       -- relChildSortedBinaryText_Desc. If item is String, place child in sorted position comparing siblings case-sensitive
    ' NewKey.(Optional). If provided, the item will be keyed with the new key.
    '       -- If the key results in a duplicate key, an error occurs
    ' Return value. If no errors encounterd, return value is True
    
    Dim bNewTree As Boolean
    Dim moveAfter As Boolean, kIndex As Long
    Dim Index As Long, TgtIndex As Long
    
    If m_ItemCount < 2 Then Exit Function
    
    ' validate the KeyIndex and TargetKeyIndex
    Index = ConvertKeyToIndex(KeyIndex, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.MoveItem", "Invalid Source Key or Index"
    End If
    TgtIndex = ConvertKeyToIndex(TargetKeyIndex, True)
    If TgtIndex < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.MoveItem", "Invalid Destination Key or Index"
    End If
    
    If Not NewKey = vbNullString Then Me.Key(Index) = NewKey
    
    If Relationship > relLastChild And Relationship < relChildSortedAsNumeric_Desc + 1 Then
        ' place in a sorted position.... determine the position among its other siblings
        kIndex = GetChildInsertPosition(Index, TgtIndex, (Relationship - 6 + (Relationship Mod 2 = 1)), (Relationship Mod 2 = 0))
        If kIndex = 0 Then ' this will be the last child
            Relationship = relLastChild
        Else    ' inserted before or after an exisiting node
            TgtIndex = kIndex                   ' cache the existing node's index
            Relationship = relPrevSibling       ' descending order, insert after
        End If
    End If
    
    Select Case Relationship
    Case relFirstChild, relLastChild
        If IsArrayEmpty(Not m_Col(TgtIndex).Family.Tree) Then
            ' target index has no children yet
            bNewTree = True
        Else
            ' set the adjusted target index based on Relationship
            If Relationship = relLastChild Then
                moveAfter = True
                TgtIndex = m_Col(TgtIndex).Family.Tree(treelast)
            Else
                TgtIndex = m_Col(TgtIndex).Family.Tree(treefirst)
            End If
        End If
    Case relFirstSibling ' set the adjusted target index based on Relationship
        TgtIndex = m_Col(m_Col(TgtIndex).Parent).Family.Tree(treefirst)
    Case relLastSibling ' set the adjusted target index based on Relationship
        TgtIndex = m_Col(m_Col(TgtIndex).Parent).Family.Tree(treelast)
        moveAfter = True
    Case relNextSibling ' set the adjusted target index based on Relationship
        If m_Col(m_Col(TgtIndex).Parent).Family.Tree(treelast) = TgtIndex Then
            moveAfter = True
        Else
            TgtIndex = m_Col(TgtIndex).Elder
        End If
    Case relPrevSibling
        ' no adjustment needed; default value is moveBefore not moveAfter
    End Select
        
    If Index = TgtIndex Then    ' if the source & target are same, simply abort
        MoveItem = True
        Exit Function
    End If
    
    ' moving is a rather simple process of updating linked lists,
    ' the physical arrays won't be touched; their linked lists are just updated
    ' 1) The previous and next sibling links around the moved node must now connect to each other
    ' 2) the previous and next sibling links around the new node position must now connect to the moved node
    
    ' Adjust the current node sibling references
    ' update linked lists around this node. We will be removing it from the chain
    With m_Col(Index)
        If .Younger = 0 Then ' this was the first sibling
            If .Elder = 0 Then  ' this was also the last sibling; an only child
                Erase m_Col(.Parent).Family.Tree()
            Else
                m_Col(.Parent).Family.Tree(treefirst) = .Elder ' the next sibling is now the first
                m_Col(.Elder).Younger = 0        ' and it has no previous sibling
            End If
        Else    ' make this node's prev sibling reference this node's next sibling
            m_Col(.Younger).Elder = .Elder
        End If
        
        If .Elder = 0 Then ' this was the last sibling
            If Not .Younger = 0 Then
                m_Col(.Parent).Family.Tree(treelast) = .Younger  ' the prev node is now the last child
                m_Col(.Younger).Elder = 0          ' and it has no next sibling
            End If
        Else    ' make this node's next sibling reference this node's prev sibling
            m_Col(.Elder).Younger = .Younger
        End If
    End With
    
    ' Adjust target position & its sibling's References. We will be inserting into the chain
    If bNewTree Then
        InitializeTree TgtIndex, Index
        m_Col(Index).Elder = 0
        m_Col(Index).Younger = 0
    Else
        ' now adjust the target node's siblings/parent linked lists
        With m_Col(TgtIndex)
            If moveAfter = True Then
                If .Elder = 0 Then  ' was last child; moved node now is
                    m_Col(.Parent).Family.Tree(treelast) = Index
                Else
                    m_Col(.Elder).Younger = Index
                End If
                m_Col(Index).Elder = .Elder
                m_Col(Index).Younger = TgtIndex
                .Elder = Index
            Else
                If .Younger = 0 Then ' was first child; moved node now is
                    m_Col(.Parent).Family.Tree(treefirst) = Index
                Else
                    m_Col(.Younger).Elder = Index
                End If
                m_Col(Index).Elder = TgtIndex
                m_Col(Index).Younger = .Younger
                .Younger = Index
            End If
            m_Col(Index).Parent = .Parent   ' set the new parent
        End With
    End If
    MoveItem = True

End Function

Public Function MoveChildren(ParentKeyIndex As Variant, TargetKeyIndex As Variant, Optional ByVal Relationship As eNodeRelationship = relNextSibling) As Boolean

    ' PURPOSE: Move a node chain from current position to a new position and possibly a new parent.
    
    ' ParentKeyIndex. A key or Index to reference the parent node by
    ' TargetKeyIndex. The key or index to the node that the Relationship parameter applies too
    '       -- If is vbNullString or zero, then the root will be used
    ' Relationship.{Optional}. Default is relNextSibling
    '       -- relFirstChild. The new node will become the first child of the node ref'd by RelativeKeyIndex
    '       -- relLastChild. The new node will become the last child of the node ref'd by RelativeKeyIndex
    '       -- relFirstSibling. The new node will become the first child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relLastSibling. The new node will become the last child of the parent of the node ref'd by RelativeKeyIndex
    '       -- relNextSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its next sibling
    '       -- relPrevSibling. The new node will be inserted between the node ref'd by RelativeKeyIndex and its previous sibling
    '   The following do not apply for items containing arrays or objects and will reset to relLastChild if that case
    '       -- relChildSortedAsNumeric_Asc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsNumeric_Desc. If Item is String or Numeric/Date, place child in sorted position comparing siblings as Numbers
    '       -- relChildSortedAsText_Asc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedAsText_Desc. If item is String, place child in sorted position comparing siblings non case-sensitive
    '       -- relChildSortedBinaryText_Asc. If item is String, place child in sorted position comparing siblings case-sensitive
    '       -- relChildSortedBinaryText_Desc. If item is String, place child in sorted position comparing siblings case-sensitive
    ' Return value. If no errors encounterd, return value is True
    
    Dim moveAfter As Boolean, bNewTree As Boolean
    Dim Index As Long, TgtIndex As Long
    
    If m_ItemCount < 2 Then Exit Function
    
    ' validate ParentKeyIndex and TargetKeyIndex
    Index = ConvertKeyToIndex(ParentKeyIndex, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.MoveChildren", "Invalid Source Key or Index"
    End If
    TgtIndex = ConvertKeyToIndex(TargetKeyIndex, True)
    If TgtIndex < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.MoveChildren", "Invalid Destination Key or Index"
    End If
    
    ' Note: Can't sort on the fly. The children may not be currently sorted and same goes for any existing children of the new parent
    ' So, if the Relationship parameter includes a sorted child item, we will sort post-move.
    
    Select Case Relationship
    Case relFirstChild, relLastChild, relChildSortedAsNumeric_Asc, relChildSortedAsNumeric_Desc, relChildSortedAsText_Asc, _
            relChildSortedAsText_Desc, relChildSortedBinaryText_Asc, relChildSortedBinaryText_Desc
        If IsArrayEmpty(Not m_Col(TgtIndex).Family.Tree) Then
            ' target node has no children, easy...
            ' set target's family.tree to the ParentKeyIndex's family.tree
            m_Col(TgtIndex).Family.Tree = m_Col(Index).Family.Tree
            Erase m_Col(Index).Family.Tree
            ' loop thru each of the child nodes and change the parent ref
            Index = m_Col(TgtIndex).Family.Tree(treefirst)
            Do Until Index = 0
                m_Col(Index).Parent = TgtIndex
                Index = m_Col(Index).Elder
            Loop
            If Relationship > relLastChild Then
                ' sort the children if requested
                SortNodeLevel TgtIndex, (Relationship - 6 + (Relationship Mod 2 = 1)), (Relationship Mod 2 = 0), True, vbNullString
            End If
            MoveChildren = True
            Exit Function
        Else
            If TgtIndex = Index Then
                MoveChildren = True ' no action of target and source parents are same
                Exit Function
            End If
            If Relationship = relLastChild Then
                moveAfter = True
                TgtIndex = m_Col(TgtIndex).Family.Tree(treelast)
            Else
                TgtIndex = m_Col(TgtIndex).Family.Tree(treefirst)
            End If
            Index = m_Col(Index).Family.Tree(treefirst)
        End If
        
    Case relFirstSibling ' target position is parent's first child
        TgtIndex = m_Col(m_Col(TgtIndex).Parent).Family.Tree(treefirst)
        
    Case relLastSibling ' target position is parent's last child
        TgtIndex = m_Col(m_Col(TgtIndex).Parent).Family.Tree(treelast)
        moveAfter = True
        
    Case relNextSibling ' target position is next sibling
        If m_Col(m_Col(TgtIndex).Parent).Family.Tree(treelast) = TgtIndex Then
            moveAfter = True
        Else
            TgtIndex = m_Col(TgtIndex).Elder
        End If
        
    Case relPrevSibling
        ' no adjustment needed; default value is moveBefore not moveAfter
    End Select
        
    If Index = TgtIndex Then    ' if the source & target are same, simply abort
        MoveChildren = True
        Exit Function
    End If
    
    ' moving is a rather simple process of updating linked lists,
    ' the physical arrays won't be touched; neither will there indexes
    With m_Col(TgtIndex)
        If moveAfter = True Then
            If .Elder = 0 Then
                ' target was last child; now new child tree is last
                m_Col(.Parent).Family.Tree(treelast) = m_Col(m_Col(Index).Parent).Family.Tree(treelast)
            Else
                ' make tgtIndex next sibling ref new tree's last child
                m_Col(.Elder).Younger = m_Col(.Parent).Family.Tree(treelast) = m_Col(m_Col(Index).Parent).Family.Tree(treelast)
            End If
            m_Col(Index).Younger = TgtIndex
            .Elder = Index
        Else
            If .Younger = 0 Then
                ' target was first child; now new tree is first
                m_Col(.Parent).Family.Tree(treefirst) = Index
            Else
                ' make tgtIndex prev sibling ref new tree's first child
                m_Col(.Younger).Elder = Index
            End If
            ' new tree's last sibling's next sibling is the target
            m_Col(m_Col(m_Col(Index).Parent).Family.Tree(treelast)).Elder = TgtIndex
            ' make new tree's prev sibling ref target's prev sibling
            m_Col(Index).Younger = .Younger
            ' target's prev sibling is not new tree's last sibling
            .Younger = m_Col(m_Col(Index).Parent).Family.Tree(treelast)
        End If
    End With
    ' clear the old family.tree
    Erase m_Col(m_Col(Index).Parent).Family.Tree
    
    ' loop thru changing the parent of all moved children; changing Parent ref
    Do Until Index = 0
        If m_Col(Index).Parent = m_Col(TgtIndex).Parent Then Exit Do
        m_Col(Index).Parent = m_Col(TgtIndex).Parent
        Index = m_Col(Index).Elder
    Loop
    
    If Relationship > relLastChild And Relationship < relChildSortedAsNumeric_Desc + 1 Then
        ' simply ensure tgtIndex's children are sorted
        SortNodeLevel TgtIndex, (Relationship - 6 + (Relationship Mod 2 = 1)), (Relationship Mod 2 = 0), True, vbNullString
    End If
    MoveChildren = True

End Function

' PURPOSE: Get/Set the value of a collection item
Public Property Get Item(KeyIndex As Variant) As Variant
    ' KeyIndex. A key or Index to reference this item by
    '   - Cannot be vbNullString nor zero
    
    ' Note about referencing items that contain arrays.
    ' If I wanted to reference the 5th array element of an array contained in Item,
    ' my code would look something like the following.....
    ' Debug.Print myCol.Item("KeyOfItem")(5)
    
    ' for Items that contain arrays of objects, then we could use something like
    ' Set NewClassObject = myCol.Item("KeyOfItem")(5)
    ' or if you wanted to view the value of a Public class property named Misc...
    ' Debug.Print myCol.Item("KeyOfItem")(5).Misc
    
    Dim Index As Long, vType As Integer
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 1 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Key Property Get", "Invalid Key or Index"
    Else
        CopyMemory vType, ByVal VarPtr(m_Col(Index).Item), 2&
        If vType = vbObject Then
            Set Item = m_Col(Index).Item
        Else
            Item = m_Col(Index).Item
        End If
    End If
End Property
Public Property Let Item(KeyIndex As Variant, Value As Variant)
    ' KeyIndex. A key or Index to reference this item by
    ' Value. The new value for the item
    '   - Cannot be vbNullString nor zero
    Dim Index As Long
    Select Case VarType(Value)
    Case vbObject, vbDataObject
        Set Me.Item(KeyIndex) = Value
    Case Else
        Index = ConvertKeyToIndex(KeyIndex, False)
        If Index < 1 Then
            Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Item Propety Let", "Invalid Key or Index"
        Else
            m_Col(Index).Item = Value
        End If
    End Select
End Property
Public Property Set Item(KeyIndex As Variant, Value As Variant)
    ' KeyIndex. A key or Index to reference this item by
    ' Value. The new value for the item
    '   - Cannot be vbNullString nor zero
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 1 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Item Property Set", "Invalid Key or Index"
    Else
        Set m_Col(Index).Item = Value
    End If
End Property

Public Function ItemCount() As Long
    ' PURPOSE: Return the total number of items in the collection
    ItemCount = m_ItemCount
End Function

' PURPOSE: Return/Set the Parent index associated with a node
Public Property Let Parent(KeyIndex As Variant, NewParentKeyIndex As Variant)
    ' KeyIndex. A key or Index to reference this item by
    ' NewParentKeyIndex. A key or Index to the new parent node
    '   - vbNullString or zero will make the KeyIndex node a main level node
    ' Setting a parent this way is same as calling MoveItem(KeyIndex, NewParentKeyIndex, relLastChild)
    
    Dim Index As Long, pIndex As Long
    Index = ConvertKeyToIndex(KeyIndex, False) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Property Parent Let", "Invalid Key or Index"
    End If
    pIndex = ConvertKeyToIndex(NewParentKeyIndex, True)
    If pIndex < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Property Parent Let", "Invalid New Parent Key or Index"
    End If
    If Not pIndex = m_Col(Index).Parent Then MoveItem Index, pIndex, relLastChild
    
End Property
Public Property Get Parent(KeyIndex As Variant) As Variant
    ' KeyIndex. A key or Index to reference this item by
    '   - vbNullString or zero will return zero
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, False) ' get the index from passed parameter
    If Index > 0 Then Parent = m_Col(Index).Parent Else Parent = 0&

End Property

Public Function FirstChild(ParentKeyIndex As Variant) As Long
    
    ' PURPOSE: Return the first child associated with ParentKeyIndex
    
    ' ParentKeyIndex. A key or Index to reference the parent node by
    '   - vbNullString or zero will return 1st top level node
    
    Dim Index As Long
    Index = ConvertKeyToIndex(ParentKeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.FirstChild", "Invalid Key or Index"
    Else
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then FirstChild = m_Col(Index).Family.Tree(treefirst)
    End If
End Function

Public Function LastChild(ParentKeyIndex As Variant) As Long
    
    ' PURPOSE: Return the last child associated with ParentKeyIndex
    
    ' ParentKeyIndex. A key or Index to reference the parent node by
    '   - vbNullString or zero will return lasst top level node
    
    Dim Index As Long
    Index = ConvertKeyToIndex(ParentKeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.LastChild", "Invalid Key or Index"
    Else
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then LastChild = m_Col(Index).Family.Tree(treelast)
    End If
    
End Function

Public Function HasChildren(ParentKeyIndex As Variant) As Boolean

    ' PURPOSE: Return whether or not ParentKeyIndex has child nodes
    
    ' ParentKeyIndex. A key or Index to reference the parent node by
    '   - vbNullString or zero will use the root

    Dim Index As Long
    Index = ConvertKeyToIndex(ParentKeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.HasChildren", "Invalid Key or Index"
    Else
        HasChildren = (Not IsArrayEmpty(Not m_Col(Index).Family.Tree))
    End If
    
End Function

Public Function ChildCount(ParentKeyIndex As Variant) As Long
    
    ' PURPOSE: Return the number of children associated with ParentKeyIndex
    
    ' ParentKeyIndex. A key or Index to reference the parent node by
    '   - vbNullString or zero will use the root
    
    Dim Index As Long, Count As Long
    Index = ConvertKeyToIndex(ParentKeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.ChildCount", "Invalid Key or Index"
    Else
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
            Count = 1
            Index = m_Col(Index).Family.Tree(treefirst)
            Do Until m_Col(Index).Elder = 0
                Count = Count + 1
                Index = m_Col(Index).Elder
            Loop
        End If
    End If
    ChildCount = Count
    
End Function

Public Function FirstSibling(KeyIndex As Variant) As Long

    ' PURPOSE: Return the first sibling associated with KeyIndex
    
    ' KeyIndex. A key or Index to reference the node by
    '   - vbNullString or zero will return 1st top level node

    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.FirstSibling", "Invalid Key or Index"
    Else
        If Not IsArrayEmpty(Not m_Col(m_Col(Index).Parent).Family.Tree) Then
            FirstSibling = m_Col(m_Col(Index).Parent).Family.Tree(treefirst)
        End If
    End If
    
End Function

Public Function LastSibling(KeyIndex As Variant) As Long
    
    ' PURPOSE: Return the last sibling associated with KeyIndex
    
    ' KeyIndex. A key or Index to reference the node by
    '   - vbNullString or zero will return last top level node
    
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.LastSibling", "Invalid Key or Index"
    Else
        If Not IsArrayEmpty(Not m_Col(m_Col(Index).Parent).Family.Tree) Then
            LastSibling = m_Col(m_Col(Index).Parent).Family.Tree(treelast)
        End If
    End If
    
End Function

Public Property Get NextSibling(KeyIndex As Variant) As Long
    
    ' PURPOSE: Return the next sibling associated with KeyIndex
    
    ' KeyIndex. A key or Index to reference the node by
    '   - vbNullString or zero will return last top level node
    
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.NextSibling", "Invalid Key or Index"
    Else
        If Index = 0 Then
            NextSibling = m_Col(0).Family.Tree(treelast)
        Else
            NextSibling = m_Col(Index).Elder
        End If
    End If
    
End Property

Public Function PreviousSibling(KeyIndex As Variant) As Long

    ' PURPOSE: Return the previous sibling associated with KeyIndex
    
    ' KeyIndex. A key or Index to reference the node by
    '   - vbNullString or zero will return 1st top level node
    
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, True) ' get the index from passed parameter
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.PreviousSibling", "Invalid Key or Index"
    Else
        If Index = 0 Then
            PreviousSibling = m_Col(0).Family.Tree(treefirst)
        Else
            PreviousSibling = m_Col(Index).Younger
        End If
    End If
End Function

Public Function SiblingOffsetIndex(Optional ParentKey As Variant = 0&, _
                    Optional ByVal Offset As Long = 0&, _
                    Optional ReverseOffset As Boolean = False)
    
    ' PURPOSE: Provide a relatively easy way to get a sibling based off of the
    ' first or last sibling in the family tree.
    
    ' When the collection is displayed in a listbox for example, it may be very difficult
    ' to associate an item with a list item without having to keep a separate list of keys
    ' cross-referenced to the list item indexes.
    ' If keys are numeric, then the list ItemData property can store the keys, but if
    ' the keys are alpha-numeric or keys are not used, it is very difficult to reference
    ' a list item to a collection item.
    ' When the list item displays the tree from 1st siblings to last siblings, this
    ' function is ideal to return the index of any of the siblings
    
    ' All parameters are optional
    ' ParentKey. A key or item index of the node(s) parent.
    '   -- if vbNullString or zero is passed, the root is used
    ' Offset. The number of siblings to count. Return value based off of ReverseOffset parameter
    ' ReverseOffset. Must be a positive value. Default is False
    '   - If False, then offset starts from first sibling
    '       if Offset = 0 :: returns the first sibling index
    '       if Offset > number of siblings :: returns the last sibling
    '       if Offset = 1 to number of siblings-1 :: returns the index of that sibling
    '   - If True, then offset starts from last sibling
    '       if Offset = 0 :: returns the last sibling index
    '       if Offset > number of siblings :: returns the first sibling
    '       if Offset = 1 to number of siblings-1 :: returns the index of that sibling
    
    '   - If the offset is a negative value, then zero is returned
    
    ' Return value of zero indicates the ParentKey has no children
    
    Dim Index As Long, cIndex As Long
    
    If Offset < 0 Then Exit Function
    
    Index = ConvertKeyToIndex(ParentKey, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.SiblingOffsetIndex", "Invalid Key or Index"
    End If
    If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
    
        If ReverseOffset = True Then
            cIndex = m_Col(Index).Family.Tree(treelast)
            Do Until cIndex = 0
                Offset = Offset - 1
                If Offset < 0 Then Exit Do
                cIndex = m_Col(cIndex).Younger
            Loop
            If cIndex = 0 Then cIndex = m_Col(Index).Family.Tree(treefirst)
        Else
            cIndex = m_Col(Index).Family.Tree(treefirst)
            Do Until cIndex = 0
                Offset = Offset - 1
                If Offset < 0 Then Exit Do
                cIndex = m_Col(cIndex).Elder
            Loop
            If cIndex = 0 Then cIndex = m_Col(Index).Family.Tree(treelast)
        End If
        
    End If
    SiblingOffsetIndex = cIndex
    
End Function

' PURPOSE: Set key case-sensitivity
Public Property Get KeysCaseSensitive() As Boolean
    KeysCaseSensitive = (m_CompareType = keyCaseSensitive)
End Property
Public Property Let KeysCaseSensitive(isCaseSensitive As Boolean)
    ' Property is writable only when no items or no keys exist in the collection
    If m_KeyCount = 0 Then m_CompareType = Abs(isCaseSensitive = False)
End Property

Public Function KeyToIndex(ByVal Key As String) As Long

    ' PURPOSE: Return an Item index from its Key value
    
    ' Key. An existing Key within the collection
    '   - vbNullString will return error
    KeyToIndex = FindKey(Key, False)
    If KeyToIndex < 1 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.KeyToIndex", "Invalid Key or Index"
    Else
        KeyToIndex = m_Keys(KeyToIndex).Index
    End If

End Function

Public Function IndexFromKeyIndex(ByVal Index As Long) As Long
    
    ' PURPOSE: Return an Item index from its Key value's Index
    
    ' Index. If zero, then returns the number of keys used in the collection
    '        Otherwise must be a valid number in the range of 1 and number of keys used and
    '        returns the item index refernced by the index into the m_Keys array
    
    ' Keys, internally are always sorted in ascending order. If you want your collection sorted, you could
    ' create keys that can be used as a sort order and also as a key. Then to get your sorted collection,
    ' instead of referencing an item by its key index, you would reference it by the index returned from
    ' this function.  Of course this assumes you are keying all items you are sorting and this method
    ' really works only for flat collections, no parent/child heirarchy.
    ' An example follows
    
    '    Dim kIndex As Long
    '    For kIndex = 1 To myCol.IndexFromKeyIndex(0) ' ascending order
    '        Debug.Print myCol.Item(myCol.IndexFromKeyIndex(kIndex))
    '    Next
    '    For kIndex = myCol.IndexFromKeyIndex(0) To 1 Step -1 ' descending order
    '        Debug.Print myCol.Item(myCol.IndexFromKeyIndex(kIndex))
    '    Next
    
    If Index = 0 Then
        IndexFromKeyIndex = m_KeyCount
    ElseIf Index > 0 Then
        If Index > m_KeyCount Then
            Err.Raise ERROR_INVALID_KEYINDEX, "Collection.IndexFromKeyIndex", "Invalid Index"
        Else
            IndexFromKeyIndex = m_Keys(Index).Index
        End If
    Else
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.IndexFromKeyIndex", "Invalid Index"
    End If
    
End Function

Public Function KeyExists(ByVal Key As String) As Boolean
    
    ' PURPOSE: Return if a key exists in the collection
    
    ' Key. An existing Key within the collection
    If Not Key = vbNullString Then KeyExists = (FindKey(Key) > 0)

End Function

Public Function RandomKey() As String
    
    ' PURPOSE: Somewhat simplistic random key generator to create a unique key for the collection
    On Error GoTo ExitRoutine
    
    Dim sKey As String * 16
    Dim x As Long
    Const maxSeed As Long = 57
    
    Do
        For x = 1 To 16
            Mid$(sKey, x, 1) = Chr$(Int(Rnd * maxSeed) + 65)
        Next
    Loop Until FindKey(sKey) < 0                    ' verify it doesn't exist
    
    RandomKey = sKey

ExitRoutine:
    If Err Then
        Err.Clear   ' Rnd causing blasted "Expression too complex" errors on my PC
        Resume
    End If
End Function

Public Sub ClearKeys(Optional KeyIndex As Variant = 0&, Optional ByVal Level As eSortLevel = brchAllLevels)
    ' PURPOSE: Clear keys. Useful should one want to re-key collection items.
    ' KeyIndex. If zero, then the upper level nodes will be used. Otherwise, only the level of nodes that contain KeyIndex
    ' Level (Optional)
    '   - brchAllLevels. Both sibling/child branches will be sorted, starting with KeyIndex's sibling branch & continuing for all descendants
    '   - brchSiblings. Only the siblings of KeyIndex will be sorted
    '   - brchChildren. Only the immediate children of KeyIndex will be sorted. Any children nodes of these children will not be sorted
    
    Dim KeyID As Long, Index As Long
    Index = ConvertKeyToIndex(KeyIndex, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.ClearKeys", "Invalid Key or Index"
    End If
    
    If Level = brchAllLevels Then   ' delete every key from collection
        For KeyID = 1 To m_KeyCount
            m_Col(m_Keys(KeyID).Index).Key = 0
        Next
        ReDim m_Keys(0 To 0)
        m_KeyCount = 0
        
    Else
        If Level = brchSiblings Then
            ' delete keys for every sibling of KeyIndex else for every child of KeyIndex
            Index = m_Col(Index).Parent
        End If
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
            Index = m_Col(Index).Family.Tree(treefirst)
            Do Until Index = 0
                If Not m_Col(Index).Key = 0 Then
                    DeleteKey m_Col(Index).Key
                    m_Col(Index).Key = 0
                End If
                Index = m_Col(Index).Elder
            Loop
        End If
    End If
End Sub

Public Function ParentRootNode(KeyIndex As Variant) As Long
    
    ' PURPOSE: Return the main level node associated with this node
    
    Dim Index As Long, pIndex As Long
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.ParentRootNode", "Invalid Key or Index"
    End If
    
    If m_Col(Index).Parent = 0 Then
        ParentRootNode = Index
    Else
        pIndex = m_Col(Index).Parent
        Do Until pIndex = 0
            ParentRootNode = pIndex
            pIndex = m_Col(pIndex).Parent
        Loop
    End If
    
End Function

Public Function IsAncestorOf(AncestorKeyIndex As Variant, ChildKeyIndex As Variant) As Boolean
    
    ' PURPOSE: Return if ChildKeyIndex a descendent of AncestorKeyIndex
    
    Dim Index As Long, cIndex As Long
    Index = ConvertKeyToIndex(AncestorKeyIndex, False)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.IsAncestorOf", "Invalid Ancestor Key or Index"
    End If
    cIndex = ConvertKeyToIndex(ChildKeyIndex, False)
    If cIndex < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.IsAncestorOf", "Invalid Child Key or Index"
    End If
    
    Do
        If m_Col(cIndex).Parent = Index Then
            IsAncestorOf = True
            Exit Do
        ElseIf m_Col(cIndex).Parent = 0 Then
            Exit Do
        End If
        cIndex = m_Col(cIndex).Parent
    Loop
    
End Function


' PURPOSE: Get/Set the Key of a collection item
Public Property Get Key(KeyIndex As Variant) As String
    ' KeyIndex. An existing Key or Index within the collection
    Dim Index As Long
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Key Property Get", "Invalid Key or Index"
    Else
        Key = VBStringFromBSTRpointer(m_Col(Index).Key)
    End If
End Property
Public Property Let Key(KeyIndex As Variant, ByVal NewKey As String)
    ' KeyIndex. An existing Key or Index within the collection
    ' Passing vbNullString for NewKey will remove the key for the item at KeyIndex
    Dim Index As Long, kIndex As Long
    Index = ConvertKeyToIndex(KeyIndex, False)
    If Index < 1 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.Key Property Let", "Invalid Key or Index"
    Else
        If NewKey = vbNullString Then
            DeleteKey m_Col(Index).Key
            m_Col(Index).Key = 0
        Else
            kIndex = FindKey(NewKey)
            If kIndex = ERROR_NOT_FOUND Then    ' doesn't exist
                DeleteKey m_Col(Index).Key      ' remove old key
                kIndex = FindKey(NewKey, Index) ' add new key
                m_Col(Index).Key = StrPtr(m_Keys(kIndex).Key)
                
            ElseIf m_Col(kIndex).Item = Index Then
                ' it is possible user has case sensitivity off and is replacing key "van halen" with "Van Halen"
                ' therefore, we will change the key since it results in referencing the same item
                m_Keys(kIndex).Key = NewKey
                m_Col(Index).Key = StrPtr(m_Keys(kIndex).Key)
                
            Else
                Err.Raise ERROR_KEY_DUPLICATED, "Collection.Key Property Let", "Key Already Exists in this Collection"
            End If
        End If
    End If
End Property

' PURPOSE: Get/Set the buffer. Buffer helps prevent redim'ing collection after every new Add/Delete
Public Property Let CollectionBufferLimit(ByVal Value As Long)
    If Value > 0 And Value < 1001 Then
        m_Buffer = Value   ' use some reasonable min/max limits
    ElseIf Value = 0 Then
        m_Buffer = Value        ' trim the collection items
        If UBound(m_Col) > m_ItemCount Then
            ReDim Preserve m_Col(0 To m_ItemCount)
        End If
        If UBound(m_Keys) > m_KeyCount Then
            ReDim Preserve m_Keys(0 To m_KeyCount)
        End If
    End If
End Property
Public Property Get CollectionBufferLimit() As Long
    CollectionBufferLimit = m_Buffer
End Property

Public Function SaveCollection(CollectionArray() As Byte) As Boolean

    ' PURPOSE: Save a collection in parent/child order from 1st Node to the last Node.
    
    ' CollectionArray. An empty/non-initialized byte array
    '   The returned byte array can be used in LoadCollection
    
    Dim ppBagOut As New PropertyBag, ppBytes(0 To 11) As Byte
    
    ' if an error occurs here, no harm no foul. The routine will abort.
    ' The only possible error I could foresee is out of memory if the
    ' collection is that massive.
    
    If SaveProperties(m_Col(0).Family.Tree(treefirst), 0, 1, ppBagOut) = True Then
        ' write the nr of items and case sensitivity flag
        CopyMemory ppBytes(0), m_ItemCount, 4&
        CopyMemory ppBytes(4), m_CompareType, 4&
        CopyMemory ppBytes(8), m_Buffer, 4&
        ppBagOut.WriteProperty "Count", ppBytes
    
        CollectionArray = ppBagOut.Contents()
        SaveCollection = True
        
    End If
    
End Function

Public Function LoadCollection(CollectionBytes() As Byte) As Boolean

    ' PURPOSE: Take a previously Saved Collection and convert its byte content to a new collection
    
    ' CollectionArray. An array that was created via the SaveCollection method

    Dim ppBag As New PropertyBag
    Dim NodeID As Long, sKey As String
    Dim vType As Integer, itemPtr As Long
    Dim ppBytes() As Byte, nSize As Long
    Dim arrayPtr As Long
    Dim tUnSupported As Object
    
    ' erase previous collection if any
    Call ClearCollection
    m_CompareType = keyCaseInsensitive
    
    On Error Resume Next
    ppBag.Contents = CollectionBytes()  ' set collection to property bag
    If Err Then
        Err.Clear
        On Error GoTo 0
        Err.Raise ERROR_INVALID_COLLECTION, "Collection.LoadCollection", "Invalid Collection Items for this Class"
    
    Else
        On Error GoTo ExitRoutine
        ReDim ppBytes(0 To 7)
        ppBytes = ppBag.ReadProperty("Count", ppBytes)  ' get nr items in passed collection & key case-sensitivity
        CopyMemory m_ItemCount, ppBytes(0), 4&          ' nr of items
        CopyMemory m_CompareType, ppBytes(4), 4&        ' key case-sensitivity
        CopyMemory m_Buffer, ppBytes(8), 4&             ' array buffer
        ReDim m_Col(0 To m_ItemCount)                   ' resize collection
    End If
    
    For NodeID = 1 To m_ItemCount                       ' loop thru each collection item
    
        ppBytes = ppBag.ReadProperty("N" & NodeID, ppBytes) ' get item parent, data type, & key
        If Not IsArrayEmpty(Not ppBytes) Then
        
            CopyMemory ByVal VarPtr(m_Col(NodeID).Parent), ppBytes(0), 4&   ' set the item's .Parent
            CopyMemory vType, ppBytes(4), 2&                                ' set item's data type
            
            ' update the item parent's family.tree structure
            With m_Col(m_Col(NodeID).Parent).Family
                If IsArrayEmpty(Not .Tree) Then
                    InitializeTree m_Col(NodeID).Parent, NodeID ' this is the 1st & last chld of the parent
                Else
                    m_Col(NodeID).Younger = .Tree(treelast) ' item's prev sibling is parent's last child
                    m_Col(.Tree(treelast)).Elder = NodeID   ' parent's last child's sibling is the item
                    .Tree(treelast) = NodeID                ' parent's new last child is the item
                End If
            End With
            
            ' check for a key
            If UBound(ppBytes) > 5 Then
                nSize = UBound(ppBytes) - 5                         ' key consists of n bytes
                sKey = String(nSize \ 2, ".")                       ' create blank key of n/2 bytes
                CopyMemory ByVal StrPtr(sKey), ppBytes(6), nSize    ' copy key bytes to key variable
                nSize = FindKey(sKey, NodeID)                       ' add the key to m_Keys()
                m_Col(NodeID).Key = StrPtr(m_Keys(nSize).Key)       ' update the item's .Key pointer
                nSize = 0
            Else
                sKey = vbNullString                                 ' no key
            End If
            Erase ppBytes
            
            If Not ((vType And vbArray) = vbArray) Then ' is the item data an array?
                
                    ' no array items
                    Select Case vType And Not &H4000
                        ' flag when unsupported item was attempted to be saved
                    Case vbError, vbDataObject, vbUserDefinedType ' unsupported
                        Set m_Col(NodeID).Item = tUnSupported
                    
                    Case vbObject
                        ' call routine to deserialize object; flag used to indicate
                        ' object is to be directly assigned to this item
                        SetSerializedItem vType Or &H8000, NodeID, "S" & NodeID, ppBag
                        
                    Case Else
                        ' simply apply to the collection item
                        m_Col(NodeID).Item = ppBag.ReadProperty("S" & NodeID, Empty)
                    End Select
                
            Else    ' arrays
            
                If vType = -1 Then
                    Set m_Col(NodeID).Item = ppBag.ReadProperty("S" & NodeID, Nothing)
                Else
            
                    ' add a varType to our collection's item
                    If (vType And &H4000) = &H4000 Then
                        CopyMemory ByVal VarPtr(m_Col(NodeID).Item), vbObject, 2&
                    Else
                        CopyMemory ByVal VarPtr(m_Col(NodeID).Item), vType, 2&
                    End If
                    ' get memory address of where the variant's storage is
                    CopyMemory itemPtr, SafeOffset(VarPtr(m_Col(NodeID).Item), 8), 4&
                    ' call routine to deserialize array
                    SetSerializedItem vType, itemPtr, "S" & NodeID, ppBag
                End If
            
            End If
        End If
    Next
    LoadCollection = True
    
ExitRoutine:
    ' error? either a corrupt collection being passed or out of memory if collection is that massive
    
    Erase m_Objects ' always clear, error or not

End Function

Public Function SortCollection(KeyIndex As Variant, Optional ByVal Level As eSortLevel = brchAllLevels, _
                            Optional ByVal TextSortAs As eSortOptions = srtTextText, _
                            Optional Ascending As Boolean = True, _
                            Optional ObjectPublicProperty As String = "Name") As Boolean

    ' PURPOSE: Routine sorts a collection with limitations...
    
    ' Commonsense is that not all variable types can be compared with each other. This class will make a good
    ' effort and if it fails the function will return False. If the routine is unsuccessful, you could create
    ' your own sort routine, then call the MoveItem method to place collection items in the correct order.
    ' Tip: Moving each item, one at a time into the Parent's LastChild position results in Ascending order
    '      Moving each item, one at a time into the Parent's FirstChild position results in Descending order
    '      Remember: moving items is fast - items, themselves, are never moved; linked lists are updated
    
    ' Note. Each branch of siblings is handled separately. Some branches may be only strings, while others may be only numbers, etc.
    '   1) Branches that are non-arrayed strings, numbers, dates, booleans can be sorted.
    '   2) Branches that contain only objects can be sorted if a PUBLIC property is provided in the ObjPublicProperty parameter
    '   3) Branches that contain arrays will not be sorted
    '   4) Branches that are a mix of objects and other variable types will not be sorted
    '   5) If a Branch consists of numbers and strings then TextSortAs will default to srtAsNumeric
    '   The sort routine takes 1st 8 items and last 8 items to see what types of items they are.
    '   If routine determines items can't be sorted, it won't. If it attempts to sort and an item
    '   that can't be compared is encountered, the routine aborts & no sorting is performed.
    
    ' If any errors are encountered (usually a result of trying to sort items that cannot be compared), then no sorting is performed and function returns False

    ' KeyIndex. If zero, then the upper level nodes will be used. Otherwise, only the level of nodes that contain KeyIndex
    ' Level (Optional)
    '   - brchAllLevels. Both sibling/child branches will be sorted, starting with KeyIndex's sibling branch & continuing for all descendants
    '   - brchSiblings. Only the siblings of KeyIndex will be sorted
    '   - brchChildren. Only the immediate children of KeyIndex will be sorted. Any children nodes of these children will not be sorted
    ' TextSortAs (Optional).
    '   - srtAsNumeric. Text will be sorted using Val(text)
    '   - srtTextBinary. Text will be sorted as case-sensitive
    '   - srtTextText. Text will be sorted as not case-sensitive
    ' Ascending (Optional). If true, items are sorted in ascending order else descending order
    ' ObjectPropertyName (Optional). Object/classes can only be sorted if you provide a PUBLIC property name that can be sorted. Default is "Name"
    
    Dim Index As Long, sortArray() As Long, rIndex As Long
    
    If m_ItemCount < 2 Then    ' abort search if not needed
        SortCollection = True   ' only one item?
        Exit Function
    End If
    
    Index = ConvertKeyToIndex(KeyIndex, True)
    If Index < 0 Then
        Err.Raise ERROR_INVALID_KEYINDEX, "Collection.SortCollection", "Invalid Key or Index"
    End If
    
    If Level = brchAllLevels Then
        SortCollection = SortRecursiveNodes(m_Col(Index).Parent, TextSortAs, Ascending, ObjectPublicProperty)
    Else
        SortCollection = SortNodeLevel(Index, TextSortAs, Ascending, (Level = brchChildren), ObjectPublicProperty)
    End If
    
End Function

Private Function SortRecursiveNodes(ByVal pNode As Long, ByRef TextSort As Long, ByRef SortOrder As Boolean, objPropName As String) As Boolean

    ' PURPOSE: Recursively navigate parent nodes in order to sort their children. Called by SortCollection
    Dim Index As Long, rtnVal As Boolean
    Index = m_Col(pNode).Family.Tree(treefirst)
    rtnVal = True
    Do Until Index = 0
        If Not (Not m_Col(Index).Family.Tree) Then
            rtnVal = SortRecursiveNodes(Index, TextSort, SortOrder, objPropName)
            If rtnVal = False Then Exit Do
        End If
        Index = m_Col(Index).Elder
    Loop
    If rtnVal Then
        SortRecursiveNodes = SortNodeLevel(pNode, TextSort, SortOrder, True, objPropName)
    End If

End Function

Private Sub InitializeTree(ParentIndex As Long, ChildIndex As Long)

    ' PURPOSE: Helper Function. Initializes a Parent node with its first child
    ' Called by: AddItem, MoveItem, LoadCollection
    With m_Col(ParentIndex).Family
        ReDim .Tree(treefirst To treelast)
        .Tree(treefirst) = ChildIndex
        .Tree(treelast) = ChildIndex
    End With
    m_Col(ChildIndex).Parent = ParentIndex
    ' the .Elder & .Younger members of m_Col(ChildIndex) are zero by default
    
End Sub

Private Function DeleteNode(ByVal Index As Long) As Boolean
    
    ' PURPOSE: Helper Function. Called by: RemoveItem
    
    ' Recursively remove any child nodes from the passed Index
    
    Dim kIndex As Long, sKey As String
    Dim lenUDT As Long, tItem() As Byte
    
    If m_ItemCount = 1 Then ' if we only have one node left, clear the collection
        Call ClearCollection
        
    Else
        ' delete any child nodes first
        If Not IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
            ' when deleting nodes, a 2 node swap is actually used...
            ' This can cause an index to change during recursion.
            ' So we use a key so we can easily identify the node index
            If m_Col(Index).Key = 0 Then    ' no key; create one for now
                sKey = RandomKey            ' it will be deleted with the item
                kIndex = FindKey(sKey, Index)
                m_Col(Index).Key = StrPtr(m_Keys(kIndex).Key)
            Else                            ' use the existing key
                sKey = VBStringFromBSTRpointer(m_Col(Index).Key)
            End If
            Do Until IsArrayEmpty(Not m_Col(Index).Family.Tree)
                DeleteNode m_Col(Index).Family.Tree(treelast)
                Index = m_Keys(FindKey(sKey)).Index
            Loop
            sKey = vbNullString ' recursive routine; clear memory used
        End If
        
        ' now delete this node, clearing memory/reducing reference counts too
        If (VarType(m_Col(Index).Item) And vbArray) = vbArray Then
            Erase m_Col(Index).Item
        Else
            Select Case VarType(m_Col(Index).Item)
            Case vbVariant, vbObject, vbUserDefinedType, vbDataObject, vbError
                Set m_Col(Index).Item = Nothing
            Case vbString
                m_Col(Index).Item = vbNullString
            Case vbNull
                m_Col(Index).Item = Empty
            Case Else
                m_Col(Index).Item = 0
            End Select
        End If
    
        ' we don't shift array contents; which will require fixing linked lists for potentially hundreds of nodes
        ' rather we will swap the last item in the array with the one being deleted,
        ' fix linked lists for just 2 nodes and then delete the item
        ' 1) The previous and next sibling links around the deleted node must now connect to each other
        ' 2) the previous and next sibling links around the node being swapped must now connect to the swapped position
        
        ' update siblings linked lists for departing node
        With m_Col(Index)
        
            If .Younger = 0 Then ' this was the first sibling
                If .Elder = 0 Then  ' this was also the last sibling; an only child
                    Erase m_Col(.Parent).Family.Tree()
                Else
                    m_Col(.Parent).Family.Tree(treefirst) = .Elder ' the next sibling is now the first
                    m_Col(.Elder).Younger = .Younger       ' and it has no previous sibling
                End If
            Else    ' make this node's prev sibling reference this node's next sibling
                m_Col(.Younger).Elder = .Elder
            End If
            
            If .Elder = 0 Then ' this was the last sibling
                If Not .Younger = 0 Then
                    m_Col(.Parent).Family.Tree(treelast) = .Younger  ' the prev node is now the last child
                    m_Col(.Younger).Elder = .Elder          ' and it has no next sibling
                End If
            Else    ' make this node's next sibling reference this node's prev sibling
                m_Col(.Elder).Younger = .Younger
            End If
        End With
        
        ' remove the deleted item's key from the m_Keys() array. Does require shifting array since these are sorted
        DeleteKey m_Col(Index).Key
        
        ' now since we will be moving the last item, we need to update its sibling's properties too
        If Not m_ItemCount = Index Then
        
            With m_Col(m_ItemCount)
                
                ' update last item's key reference as needed
                If Not .Key = 0 Then
                    ' the last item in the array has a key; make its .Index ref point to where the item will be moved to
                    kIndex = FindKey(VBStringFromBSTRpointer(.Key))
                    If kIndex > 0 Then m_Keys(kIndex).Index = Index
                End If
            
                If .Younger = 0 Then ' this item is the first child of its parent
                    m_Col(.Parent).Family.Tree(treefirst) = Index
                Else
                    m_Col(.Younger).Elder = Index   ' its prev sibling's .Elder now points to new Index
                End If
                If .Elder = 0 Then ' this is parent's last child
                    m_Col(.Parent).Family.Tree(treelast) = Index
                Else
                    m_Col(.Elder).Younger = Index   ' its next sibling's .Younger now points to new Index
                End If
                If Not IsArrayEmpty(Not .Family.Tree) Then ' this node has children
                    ' the parent index will change, therefore, each of its children's Parent ref must change too
                    kIndex = .Family.Tree(treefirst)
                    Do Until kIndex = 0
                        m_Col(kIndex).Parent = Index
                        kIndex = m_Col(kIndex).Elder
                    Loop
                End If
            End With
            
            ' last step, swap the collection items
            ' move the last item into the deleted items position
            lenUDT = LenB(m_Col(0)) ' use LenB vs hard coding UDT length in case you changed .Item from Variant to something else
            ReDim tItem(1 To lenUDT)
            CopyMemory tItem(1), ByVal VarPtr(m_Col(Index)), lenUDT
            CopyMemory ByVal VarPtr(m_Col(Index)), ByVal VarPtr(m_Col(m_ItemCount)), lenUDT
            CopyMemory ByVal VarPtr(m_Col(m_ItemCount)), tItem(1), lenUDT
            Erase tItem
        End If
        
        m_Col(m_ItemCount).Elder = 0
        m_Col(m_ItemCount).Parent = 0
        m_Col(m_ItemCount).Younger = 0
        m_Col(m_ItemCount).Key = 0
        m_ItemCount = m_ItemCount - 1
        
        ' tidy up; accounting for possible overflow on m_Keys()
        If m_KeyCount >= MAXINT_4 - m_Buffer Then
            kIndex = MAXINT_4
        Else
            kIndex = m_KeyCount + m_Buffer ' this line could possibly overflow
        End If
        If kIndex < UBound(m_Keys) Then ReDim Preserve m_Keys(0 To m_KeyCount)
        ' tidy up; accounting for possible overflow on m_Col()
        If m_ItemCount >= MAXINT_4 - m_Buffer Then
            kIndex = MAXINT_4
        Else
            kIndex = m_ItemCount + m_Buffer ' this line could possibly overflow
        End If
        If kIndex < UBound(m_Col) Then ReDim Preserve m_Col(0 To m_ItemCount)
    End If
    DeleteNode = True

End Function

Private Function FindKey(ByVal Key As String, Optional ByVal ItemPosition As Long = 0) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Helper function. Searches the m_Keys() array for a match of passed Key parameter
    ' Called by: almost every public property/routine and most private ones too
    
    ' [in] Key {Required}. A key to search or add
    ' [in] ItemPosition {Optional}.
    '       If > 0, then the Key is not expected to exist and if does exist, return value is ERROR_KEY_DUPLICATED else is Key's Index
    '       If = 0 and Key does not exist then return value is ERROR_NOT_FOUND else is the Key's Index
    ' [out] Return value: Either ERROR_NOT_FOUND, ERROR_NOT_FOUND or Key.newIndex member

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    UB = m_KeyCount
    LB = 1
    
    Do Until LB > UB
        newIndex = LB + ((UB - LB) \ 2)
        Select Case StrComp(m_Keys(newIndex).Key, Key, m_CompareType)
        Case 0
            Exit Do
        Case 1  ' new key is lower in sort order
            UB = newIndex - 1
        Case Else ' new key is higher in sort order
            LB = newIndex + 1
        End Select
    Loop

    If LB > UB Then  ' key was not found
        If ItemPosition = 0 Then
            FindKey = ERROR_NOT_FOUND
        Else
            If StrComp(m_Keys(newIndex).Key, Key, m_CompareType) = -1 Then newIndex = newIndex + 1
            
            m_KeyCount = m_KeyCount + 1
            '^^ can never overflow because before this is called, a test in AddItem
            '   guarantees the collection does not exceed MAXINT_4. Know that
            '   the m_Keys() array can never be larger than the m_Col() array.
            ' However, check below could overflow by adding m_Buffer to m_KeyCount
            If m_KeyCount > UBound(m_Keys) Then
                If m_KeyCount >= MAXINT_4 - m_Buffer Then ' check for out of bounds
                    ReDim Preserve m_Keys(0 To MAXINT_4)
                Else
                    ReDim Preserve m_Keys(0 To m_KeyCount + m_Buffer)
                End If
            End If
            If newIndex < m_KeyCount Then
                Dim swapKey(1 To 8) As Byte
                CopyMemory swapKey(1), ByVal VarPtr(m_Keys(m_KeyCount)), 8&
                CopyMemory ByVal VarPtr(m_Keys(newIndex + 1)), ByVal VarPtr(m_Keys(newIndex)), (m_KeyCount - newIndex) * 8&
                CopyMemory ByVal VarPtr(m_Keys(newIndex)), swapKey(1), 8&
            End If
            m_Keys(newIndex).Key = Key
            m_Keys(newIndex).Index = ItemPosition
            FindKey = newIndex
            
        End If
    Else
        If ItemPosition = 0& Then
            FindKey = newIndex
        Else
            FindKey = ERROR_KEY_DUPLICATED
        End If
    End If

End Function

Private Function DeleteKey(Optional KeyPtr As Long = 0&, Optional ByVal KeyValue As String = vbNullString) As Boolean
    
    ' PURPOSE: Helper Function. Remove a key from the m_Keys() array
    ' Called by: DeleteNode, Key property, RemoveChildren
    
    ' KeyPtr.{Optional} If used, it is a pointer to a key. If not used, KeyValue is used
    ' KeyValue.{Optional} If used, it is the key, else KeyPtr is used
    Dim Index As Long
    Dim tItem(1 To 8) As Byte
    
    If Not KeyPtr = 0 Then KeyValue = VBStringFromBSTRpointer(KeyPtr)
    If KeyValue = vbNullString Then Exit Function
    Index = FindKey(KeyValue)
    
    If Index > 0 Then
        m_Keys(Index).Key = vbNullString
        If Index < m_KeyCount Then  ' shift down, unless key is last one in the array
            CopyMemory tItem(1), ByVal VarPtr(m_Keys(Index)), 8&   ' copy item's pointer refs & shift array down one
            CopyMemory ByVal VarPtr(m_Keys(Index)), ByVal VarPtr(m_Keys(Index + 1)), (m_KeyCount - Index) * 8&
            CopyMemory ByVal VarPtr(m_Keys(m_KeyCount)), tItem(1), 8&   ' replace last item's refs with deleted item's refs
        End If
        m_KeyCount = m_KeyCount - 1 ' decrement number of keys used
        DeleteKey = True
    End If

End Function

Private Function IsArrayEmpty(ByVal arrayPtr As Long) As Boolean
    ' PURPOSE: Helper Function. Tests to see if an array has been initialized
    ' Called by majority of routines in this class
    IsArrayEmpty = (arrayPtr = -1)
End Function

Private Function ConvertKeyToIndex(inKey As Variant, IncludeRoot As Boolean) As Long
    
    ' PURPOSE: Helper Function. Converts a Key to a specific item Index
    ' Called by majority of public properties/methods
    
    On Error Resume Next
    Dim Index As Long, rtnVal As Long
   
    rtnVal = ERROR_INVALID_KEYINDEX ' default return value
    
    Select Case VarType(inKey)
    
        Case vbString
            If inKey = vbNullString Then
                If IncludeRoot Then rtnVal = 0
            Else
                rtnVal = FindKey(CStr(inKey))             ' get key index
                If rtnVal > 0 Then rtnVal = m_Keys(rtnVal).Index
            End If
            
        Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency ' numeric parameter, will be the index
            Index = Val(inKey)  ' could error out if a single/double/currency large number is passed
            If Index = 0 And IncludeRoot = True Then
                rtnVal = 0
            Else
                If Index > 0 Then
                    If Not Index > m_ItemCount Then rtnVal = Index
                End If
            End If
            
    End Select
    
    If Err Then Err.Clear           ' if error occurs, then the key is not within the bounds of a Long; invalid key nonetheless
    ConvertKeyToIndex = rtnVal
    
End Function

Private Function SafeOffset(ByVal Ptr As Long, Offset As Long) As Long

    ' ref http://support.microsoft.com/kb/q189323/ ' unsigned math
    ' Purpose: Provide a valid pointer offset
    ' Called by: GetSerializedItem, LoadCollection, SaveProperties, VBStringFromBSTRpointer
    
    ' If a pointer +/- the offset wraps around the high bit of a long, the
    ' pointer needs to change from positive to negative or vice versa.
    
    ' A return of zero indicates the offset exceeds the min/max unsigned long bounds
    
    Const MAXINT_4NEG As Long = -2147483648#

    If Offset = 0 Then
        SafeOffset = Ptr
    Else
    
        If Offset < 0 Then ' subtracting from pointer
            If Ptr < MAXINT_4NEG - Offset Then
                ' wraps around high bit (backwards) & changes to Positive from Negative
                SafeOffset = MAXINT_4 - ((MAXINT_4NEG - Ptr) - Offset - 1)
            ElseIf Ptr > 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr > -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        Else    ' Adding to pointer
            If Ptr > MAXINT_4 - Offset Then
                ' wraps around high bit (forward) & changes to Negative from Positive
                SafeOffset = MAXINT_4NEG + (Offset - (MAXINT_4 - Ptr) - 1)
            ElseIf Ptr < 0 Then ' verify pointer does not wrap around 0 bit
                If Ptr < -Offset Then SafeOffset = Ptr + Offset
            Else
                SafeOffset = Ptr + Offset
            End If
        End If
    End If

End Function

Private Function VBStringFromBSTRpointer(ByVal stringPointer As Long) As String

    ' PURPOSE: Using only a VB StrPtr variable, return the string at that pointer
    ' Called by: DeleteKey, DeleteNode, GetSerializedItem, Key property, RemoveChildren
    
    ' Note. The length of the string is stored at the 4 bytes just before the StrPtr
    Dim strLen As Long
    Dim tString As String
    
    If Not stringPointer = 0 Then ' else pointer is to a vbNullString
        strLen = SafeOffset(stringPointer, -4)
        If Not strLen = 0 Then
            CopyMemory strLen, ByVal strLen, 4&
            If Not strLen = 0 Then
                ' remember, VB strings are 2bytes per character
                tString = String(strLen \ 2, ".")
                CopyMemory ByVal StrPtr(tString), ByVal stringPointer, strLen
            End If
        End If
    End If
    VBStringFromBSTRpointer = tString

End Function

Private Sub ClearCollection()
    ' PURPOSE: Helper Function. Resets the collection
    ' Called by: DeleteNode, LoadCollection, RemoveChildren, Class_Initialize
    ReDim m_Col(0 To 0)     ' never completely erase; m_Col(0) is used to determine what type of base object (Variant, Object, etc)
    ReDim m_Keys(0 To 0)    ' never completley erase; m_Keys(0) is the key for the collection base
    m_KeyCount = 0
    m_ItemCount = 0
End Sub

Private Function SaveProperties(ByVal NodeID As Long, ByVal ParentID As Long, ByRef itemNr As Long, ByRef ppBag As PropertyBag) As Boolean

    ' PURPOSE: Recursively save items in parent/child order from 1st Node to the last Node
    On Error GoTo ExitRoutine
    
    Dim vType As Integer
    Dim itemPtr As Long
    Dim ppBytes() As Byte
    Dim lValue As Long
    '^^ 0-3 is New Parent Index, 4-5 is VarType, 6-xxx is unicode key
    Do Until NodeID = 0
    
        ' get the item's true VarType. Calling VarType on a stdPicture will return vbLong not vbObject
        CopyMemory vType, ByVal VarPtr(m_Col(NodeID).Item), 2&
        vType = vType And Not &H4000 ' remove any ByRef flags
        If vType = vbObject Then
            'flag indicating object may be cachable into property bag w/o serialization
            If Not VarType(m_Col(NodeID).Item) = vType Then vType = -1
        Else
            vType = VarType(m_Col(NodeID).Item)
        End If
        If m_Col(NodeID).Key = 0 Then                   ' does this item have a key?
            ReDim ppBytes(0 To 5)                       ' create gbase item heirarchy, no key
        Else
            itemPtr = SafeOffset(m_Col(NodeID).Key, -4) ' pointer to key's size in bytes
            CopyMemory lValue, ByVal itemPtr, 4&        ' size of key in bytes
            ReDim ppBytes(0 To lValue + 5)              ' create base Item heirarchy, including key bytes
            CopyMemory ppBytes(6), ByVal m_Col(NodeID).Key, lValue ' copy key bytes to array
        End If
        
        CopyMemory ppBytes(0), ParentID, 4&                 ' copy item's ParentIndex to array
        
        If vType = -1 Then ' see if we write directly
            On Error Resume Next
            ppBag.WriteProperty "S" & itemNr, m_Col(NodeID).Item, Nothing
            If Err Then
                Err.Clear
                vType = vbObject
            End If
            On Error GoTo ExitRoutine
        ' check for unsupported types & flag to inform serialization/deserialization routines the item is unsupported
        ElseIf (vType And vbError) = vbError Then
            vType = vType Or &H4000
        ElseIf (vType And vbDataObject) = vbDataObject Then
            vType = vType Or &H4000
        ElseIf (vType And vbUserDefinedType) = vbUserDefinedType Then
            vType = vType Or &H4000
        End If
        CopyMemory ppBytes(4), vType, 2&                    ' copy item's VarType to array
        ppBag.WriteProperty "N" & itemNr, ppBytes()         ' write the array
        Erase ppBytes                                       ' recursive routine, erase now
            
        If ((vType And vbArray) = vbArray) Then
            
            If Not vType = -1 Then
                ' Send to helper function to create the serialized array or retrieve serialized array from user
                CopyMemory itemPtr, SafeOffset(VarPtr(m_Col(NodeID).Item), 8), 4&
                CopyMemory itemPtr, ByVal itemPtr, 4&
                If Not itemPtr = 0 Then                 ' else empty array
                    GetSerializedItem vType, itemPtr, "S" & itemNr, ppBag
                End If
            End If
            
        Else
            
            If (vType And vbObject) = vbObject Then
                
                If Not (vType And &H4000) = &H4000 Then ' else unsupported object
                    ' Send to helper function to retrieve serialized object from user
                    CopyMemory itemPtr, SafeOffset(VarPtr(m_Col(NodeID).Item), 8), 4&
                    CopyMemory itemPtr, ByVal itemPtr, 4&
                    If Not itemPtr = 0 Then ' else object is Nothing
                        GetSerializedItem vType, itemPtr, "S" & itemNr, ppBag
                    End If
                End If
                
            ElseIf Not (vType And &H4000) = &H4000 Then ' exclude unsupported types
                ' simple, write the item
                ppBag.WriteProperty "S" & itemNr, m_Col(NodeID).Item
            End If
            
        End If
        
        itemNr = itemNr + 1 ' increment the nr of items in the collection
        
        If Not IsArrayEmpty(Not m_Col(NodeID).Family.Tree) Then
            ' this node has children, process children now (recursive)
            SaveProperties m_Col(NodeID).Family.Tree(treefirst), itemNr - 1, itemNr, ppBag
        End If
        ' get node's next sibling
        NodeID = m_Col(NodeID).Elder
        
    Loop
    
    SaveProperties = True
        
ExitRoutine:
End Function

Private Sub GetSerializedItem(ByVal vType As Long, _
            ByVal itemPtr As Long, ByVal propName As String, ppBag As PropertyBag)
    
    ' PURPOSE: Serialize a collection item that cannot be saved to a property bag as is
    ' Called by SaveProperties
    
    ' vType: variable type with/without the vbArray flag
    ' itemPtr: Pointer to either the Variant's object data (8 byte reference)
    '          This will be a reference to an object or to an array, nothing else
    ' propName: the property to save the serialized object/array to
    ' ppBag: the property bag to save to
    
    ' Lots of memory peeking and poking. Don't mess with this too much unless you are
    ' very comfortable with memory addresses and understand VB data storage
    
    Dim tObject As Object
    Dim ppBytes() As Byte
    Dim sKey As String      ' object identifier (user-defined)
    
    If vType = vbObject Then
        ' non-arrayed object?
        On Error Resume Next
        CopyMemory tObject, itemPtr, 4&     ' create unreferenced copy of the object
        ppBag.WriteProperty propName, tObject, Nothing
        If Err Then
            Err.Clear
            On Error GoTo 0
            RaiseEvent SerializeObject(True, tObject, ppBytes(), sKey) ' send to user for serialization
            CopyMemory tObject, 0&, 4&          ' remove the unreferenced copy & save passed Variant data pointer
            If Not IsArrayEmpty(Not ppBytes) Then ppBag.WriteProperty propName, ppBytes()
            ppBag.WriteProperty propName & "K", sKey, vbNullString ' save user's object identifier
        Else
            CopyMemory tObject, 0&, 4&          ' remove the unreferenced copy & save passed Variant data pointer
            ppBag.WriteProperty propName & "K", propName & &H8000
            '^^ flag indicating we can deserialize without user input
        End If
        Exit Sub
    End If
    
    ' now we are messing with arrays, the passed itemPtr is to a SafeArray object
    
    Dim otherPtr As Long, arrayPtr As Long
    Dim lValue As Long, nSize As Long
    Dim arrayDims As Long
    Dim arrayElems As Long
    Dim arrayFeatures As Integer    ' cache vs calculating
    Dim tVariant As Variant
    
    ' get number of dimensions from the SafeArray
    CopyMemory arrayDims, ByVal itemPtr, 2& ' bytes 0:1 are dimensions
    CopyMemory arrayFeatures, ByVal SafeOffset(itemPtr, 2), 2& ' bytes 2:3 are array features
    
    arrayElems = 1                      ' calculate nr of array elements
    ' nr Elements & LBound of each dimension start 16 bytes after header
    For arrayDims = 0 To arrayDims - 1
        CopyMemory nSize, ByVal SafeOffset(itemPtr, 16 + (8 * arrayDims)), 4&
        arrayElems = arrayElems * nSize
    Next
    
    Select Case (vType And Not vbArray) And Not &H4000
    
    Case vbString   ' strings are pointers to pointers; a bit more complicated
        nSize = 0   ' Pass #1 :: count bytes needed to store string array
        CopyMemory arrayPtr, ByVal SafeOffset(itemPtr, 12), 4&          ' get pointer to safearray
        For arrayElems = 0 To arrayElems - 1                            ' loop thru each element
            CopyMemory otherPtr, ByVal SafeOffset(arrayPtr, arrayElems * 4), 4& ' Get StrPtr
            If Not otherPtr = 0 Then                                    ' else vbNullString
                CopyMemory lValue, ByVal SafeOffset(otherPtr, -4), 4&   ' get length of unicode string
                nSize = nSize + lValue                                  ' add to overall length of array
            End If
        Next
        
        lValue = 0   ' create & populate array now
        ReDim ppBytes(lValue To nSize + arrayElems * 4 - 1)
        ' Pass #2 :: store the string array with prefix as string lengths
        For arrayElems = 0 To arrayElems - 1
            CopyMemory otherPtr, ByVal SafeOffset(arrayPtr, arrayElems * 4), 4& ' get StrPtr
            If otherPtr = 0 Then                                    ' vbNullString
                lValue = lValue + 4                                 ' skip to next position in array
            Else
                otherPtr = SafeOffset(otherPtr, -4)                 ' move to 4 bytes b4 StrPtr
                CopyMemory nSize, ByVal otherPtr, 4&                ' get length of string
                CopyMemory ppBytes(lValue), ByVal otherPtr, 4 + nSize    ' copy length & string to array
                lValue = lValue + 4 + nSize                         ' move to next position in array
            End If
        Next
        ppBag.WriteProperty propName, ppBytes()                     ' write the serialized data

    Case vbObject
        ' array of objects. no pain, no gain
        ' We store these in the recursive loop we are about to trek on
        CopyMemory arrayPtr, ByVal SafeOffset(itemPtr, 12), 4&      ' get pointer to safeArray
        For arrayElems = 0 To arrayElems - 1
            CopyMemory otherPtr, ByVal SafeOffset(arrayPtr, arrayElems * 4), 4& ' get pointer to object
            If Not otherPtr = 0 Then                                ' else pointer to object that is Nothing
                GetSerializedItem vType And Not vbArray, otherPtr, propName & "." & arrayElems, ppBag
            End If
        Next
        
    Case vbVariant ' variant arrays (this version does support nested variant arrays)
        CopyMemory arrayPtr, ByVal SafeOffset(itemPtr, 12), 4&  ' pointer to SafeArray
        ' Note: Variant Arrays are 16 byte consecutive blocks of memory
        
        ' Pass #1 :: calculate size of array needed; very little is cached here
        '            we will compress 16byte blocks into blocks of 2 to 10 bytes max
        nSize = 0
        For arrayElems = 0 To arrayElems - 1
            otherPtr = SafeOffset(arrayPtr, arrayElems * 16) ' pointer to next element
            CopyMemory vType, ByVal otherPtr, 2&    ' get varType of array element
            If Not (vType And vbArray) = vbArray Then   ' not array within array
                Select Case vType And Not vbArray
                Case vbString   ' pointers to pointetrs
                    ' get the size of the string
                    CopyMemory lValue, ByVal SafeOffset(otherPtr, 8), 4&
                    CopyMemory lValue, ByVal SafeOffset(lValue, -4), 4&
                    nSize = nSize + 4 + lValue
                Case vbLong, vbSingle
                    nSize = nSize + 4
                Case vbInteger, vbBoolean
                    nSize = nSize + 2
                Case vbDouble, vbDate, vbCurrency
                    nSize = nSize + 8
                Case vbByte, vbNull
                    nSize = nSize + 1
                Case Else
                End Select
            End If
        Next
        
        ' now we got a temp array size we can use.....
        ' Pass #2 :: encode values
        ReDim ppBytes(0 To nSize + arrayElems * 2 - 1)
        nSize = 0
        For arrayElems = 0 To arrayElems - 1
        
            otherPtr = SafeOffset(arrayPtr, arrayElems * 16) ' pointer to next array element
            CopyMemory vType, ByVal otherPtr, 2&    ' cache varType of element
            CopyMemory ppBytes(nSize), vType, 2&
            nSize = nSize + 2                       ' set next storage position
            
            If (vType And vbArray) = vbArray Then   ' arrays within arrays (recursive)
                ' get SafeArray pointer
                CopyMemory otherPtr, ByVal SafeOffset(otherPtr, 8), 4&
                If Not otherPtr = 0 Then   ' else points to null array else recurse
                    GetSerializedItem vType, otherPtr, propName & "." & arrayElems, ppBag
                End If
                
            Else
            
                Select Case vType
                
                Case vbError, vbDataObject, vbUserDefinedType ' unsupported
                    CopyMemory ppBytes(nSize - 2), vType Or &H4000, 2&  ' flag as such
                Case vbObject
                    ' get pointer to the object to test for Nothing
                    ' we can cache stdPictures/Fonts & others that support persistence
                    ' test for that now
                    CopyMemory otherPtr, ByVal SafeOffset(otherPtr, 8), 4&
                    If Not otherPtr = 0 Then   ' else Object=Nothing
                        GetSerializedItem vType, otherPtr, propName & "." & arrayElems, ppBag
                    End If
                    
                Case vbString ' get length of string then store length & string
                    CopyMemory otherPtr, ByVal SafeOffset(otherPtr, 8), 4&
                    otherPtr = SafeOffset(otherPtr, -4)
                    CopyMemory lValue, ByVal otherPtr, 4&
                    CopyMemory ppBytes(nSize), ByVal otherPtr, lValue + 4
                    nSize = nSize + 4 + lValue
                    
                Case vbLong, vbSingle ' rest is the same, simply store the value
                    CopyMemory ppBytes(nSize), ByVal SafeOffset(otherPtr, 8), 4&
                    nSize = nSize + 4
                Case vbDouble, vbDate, vbCurrency
                    CopyMemory ppBytes(nSize), ByVal SafeOffset(otherPtr, 8), 8&
                    nSize = nSize + 8
                Case vbInteger, vbBoolean
                    CopyMemory ppBytes(nSize), ByVal SafeOffset(otherPtr, 8), 2&
                    nSize = nSize + 2
                Case vbByte, vbNull
                    CopyMemory ppBytes(nSize), ByVal SafeOffset(otherPtr, 8), 1&
                    nSize = nSize + 1
                Case Else
                End Select
            End If
        Next
        ppBag.WriteProperty propName, ppBytes() ' store our variant array
    
    Case vbError, vbDataObject, vbUserDefinedType ' unsupported
    
    Case Else ' arrayed numbers,dates,booleans
        ' need size of data element, then resize array
        CopyMemory nSize, ByVal SafeOffset(itemPtr, 4), 4&
        ReDim ppBytes(0 To nSize * arrayElems - 1)
        ' get the pointer to the first array element & copy
        CopyMemory arrayPtr, ByVal SafeOffset(itemPtr, 12), 4&
        CopyMemory ppBytes(0), ByVal arrayPtr, nSize * arrayElems
        ppBag.WriteProperty propName, ppBytes() ' store the array
        
    End Select
    
    ' for each array, we need to store the number of elements & dimensions (8 bytes per dimensions)
    ' This information is needed to restore the array dimensions (SafeArray)
    ReDim ppBytes(-1 To arrayDims * 8)
    CopyMemory ppBytes(-1), arrayFeatures, 2&
    CopyMemory ppBytes(1), ByVal SafeOffset(itemPtr, 16), arrayDims * 8
    ppBag.WriteProperty propName & "A", ppBytes()

End Sub

Private Sub SetSerializedItem(ByVal vType As Long, ByVal itemPtr As Long, ByVal propName As String, ppBag As PropertyBag)
            
    ' PURPOSE: Deserialize arrays and objects
    ' vType :: variable type of the array/object to be restored
    ' itemPtr :: points the 8 byte storage area of a variant
    '           Note: the storage area may be in a soft variant
    ' propName :: is the property bag item where data is to restore array/object
    ' ppBag :: the property bag
    
    ' Lots of memory peeking and poking. Don't mess with this too much unless you are
    ' very comfortable with memory addresses and understand VB data storage
    
    ' What is a soft variant? It is a memory only 16-byte block of memory that mimics
    ' a variant. It isn't a true variant, but for our purposes, VB acts as if it is.
    ' The soft variant usually will reside in an array of bytes (ppBytes or ptrArray)
    ' Variant structure is
    '   2 bytes for varType, 6 bytes reserved (used at least for vbUserDefinedTypes), 8 bytes for storage or pointers to objects/arrays
    
    Dim ppBytes() As Byte
    Dim nrDims As Long, nrElements As Long
    Dim lValue As Long, nSize As Long
    Dim safeArray() As Byte, ptrArray() As Byte
    Dim tObject As Object, sKey As String
    
    Dim tVariant As Variant, hostVariant As Variant
    
    ' only objects are called if item to be restored is not an array, else only arrays
    If Not (vType And vbArray) = vbArray Then
    
        ' get the object serialized data
        ' check for special flag indicating deserialization not needed
        sKey = ppBag.ReadProperty(propName & "K", vbNullString)
        If sKey = propName & &H8000 Then
            ' object is a stdPicture,stdFont or other object that allows persistence
            Set tObject = ppBag.ReadProperty(propName, Nothing)
        Else
            ' object is to be deserialized
            ppBytes = ppBag.ReadProperty(propName, ppBytes)
            If Not IsArrayEmpty(Not ppBytes) Then
                ' send serialized data back to user to create an object
                RaiseEvent SerializeObject(False, tObject, ppBytes, sKey)
            End If
        End If
        If Not tObject Is Nothing Then  ' user created object from serialized array
        
            ' because of the way I use soft variants, referenced counts on objects in the soft variants
            ' may not be calculated correctly; so we will cache them until the variant is completely built.
            
            If (vType And &H8000) = &H8000 Then ' this is being set to the collection item. Not in an item contained in arrays
                Set m_Col(itemPtr).Item = tObject
                
            Else ' copy the object's pointer to the passed variant's data section
                CopyMemory ByVal itemPtr, ObjPtr(tObject), 4&
                If IsArrayEmpty(Not m_Objects) Then
                    ReDim m_Objects(0)
                Else
                    ReDim Preserve m_Objects(0 To UBound(m_Objects) + 1)
                End If
                Set m_Objects(UBound(m_Objects)) = tObject  ' add to array to keep object alive until the entire variant array is processed
            End If
        End If
        Exit Sub
    End If
    
    ' arrayed. If the array was not null then there will be an "A" property too
    ppBytes = ppBag.ReadProperty(propName & "A", ppBytes)
    If IsArrayEmpty(Not ppBytes) Then Exit Sub
    
    nrElements = 1                                      ' calculate nr elements in array
    nrDims = UBound(ppBytes) \ 8                        ' 8 bytes for each dimension
    For nrDims = 0 To nrDims - 1
        CopyMemory lValue, ppBytes(nrDims * 8 + 1), 4&  ' get nr elements from the array
        nrElements = nrElements * lValue                ' FYI: 2nd 4 bytes is the LBound
    Next

    ReDim safeArray(-3 To 16 + nrDims * 8)              ' SafeArray structure
    CopyMemory safeArray(-3), vType And Not vbArray, 4& ' 4 bytes before array is data type
    CopyMemory safeArray(1), nrDims, 2&                 ' bytes 1:2 are dimensions
    CopyMemory safeArray(3), ppBytes(-1), 2&            ' bytes 3:4 are array features (128 means offset -4 is data type)
                                                        ' bytes 5:8 is data element size in bytes
                                                        ' bytes 9:12 is lock count
                                                        ' bytes 13:16 is pointer to first data element
    CopyMemory safeArray(17), ppBytes(1), nrDims * 8    ' bytes 17:xx are each dimensions (4 bytes for nrElements, 4 bytes for LBound)
        
    Erase ppBytes ' next get the serialized array
    ppBytes = ppBag.ReadProperty(propName, ppBytes)
        
    Select Case (vType And Not vbArray) And Not &H4000
    
    Case vbString ' string array
        If IsArrayEmpty(Not ppBytes) Then Exit Sub  ' null array, nothing to do
        
        ReDim ptrArray(0 To nrElements * 4 - 1) ' build a pointer array of pointers
        CopyMemory safeArray(13), VarPtr(ptrArray(0)), 4&   ' set array pointer (see SafeArray structure above)
        safeArray(5) = 4                                    ' set data element size (see SafeArray structure above)
        lValue = 0                                          ' pointer into the ppBytes() array
        For nrElements = 0 To nrElements - 1                ' loop thru, populating the pointer array
            CopyMemory nSize, ppBytes(lValue), 4&
            If nSize = 0 Then                               ' test for vbNullString
                lValue = lValue + 4                         ' next location in ppBytes() array
            Else
                CopyMemory ptrArray(nrElements * 4), VarPtr(ppBytes(lValue + 4)), 4& ' set pointer to the 4 bytes after string length
                lValue = lValue + nSize + 4                             ' next location in ppBytes() array
            End If
        Next
        ' copy the pointer to the safearray in the host variant's 8-byte data section
        CopyMemory ByVal SafeOffset(VarPtr(hostVariant), 8), VarPtr(safeArray(1)), 4&
        
    Case vbObject
        ReDim ptrArray(0 To nrElements * 4 - 1)
        CopyMemory safeArray(13), VarPtr(ptrArray(0)), 4&   ' set array pointer (see SafeArray structure above)
        safeArray(5) = 4                                    ' set data element size (see SafeArray structure above)
        lValue = (vType And Not vbArray)
        For nrElements = 0 To nrElements - 1
            SetSerializedItem lValue, VarPtr(ptrArray(nrElements * 4)), propName & "." & nrElements, ppBag
        Next
        ' copy the pointer to the safearray in the host variant's 8-byte data section
        CopyMemory ByVal SafeOffset(VarPtr(hostVariant), 8), VarPtr(safeArray(1)), 4&
        
    Case vbVariant
        If IsArrayEmpty(Not ppBytes) Then Exit Sub  ' a null array, nothing to do
        ReDim ptrArray(0 To nrElements * 16 - 1)
        CopyMemory safeArray(13), VarPtr(ptrArray(0)), 4&   ' set array pointer (see SafeArray structure above)
        safeArray(5) = 16                                    ' set data element size (see SafeArray structure above)
        nSize = 0: lValue = 0
        For nrElements = 0 To nrElements - 1
            CopyMemory lValue, ppBytes(nSize), 2&                       ' get the vartype of this arrays element
            CopyMemory ptrArray(nrElements * 16), ppBytes(nSize), 2&    ' set variant's varType
            nSize = nSize + 2                                           ' set next position
            If (lValue And vbArray) = vbArray Then
                ' an array within an array; process that now
                SetSerializedItem lValue, VarPtr(ptrArray(nrElements * 16 + 8)), propName & "." & nrElements, ppBag
            Else
                Select Case lValue And Not &H4000
                Case vbError, vbDataObject, vbUserDefinedType ' unsupported
                    CopyMemory ptrArray(nrElements * 16), vbObject, 2&   ' set variant's varType
                Case vbObject
                    SetSerializedItem lValue, VarPtr(ptrArray(nrElements * 16 + 8)), propName & "." & nrElements, ppBag
                Case vbString
                    CopyMemory lValue, ppBytes(nSize), 4&       ' get size of string
                    If Not lValue = 0 Then                      ' else vbNullString
                        ' set the variant's pointer to location of BSTR in our array
                        CopyMemory ptrArray(nrElements * 16 + 8), VarPtr(ppBytes(nSize + 4)), 4&
                    End If
                    nSize = nSize + lValue + 4                  ' set next position
                    lValue = 0                                  ' reset for next loop
                Case vbLong, vbSingle   ' copy from our array & set next position
                    CopyMemory ptrArray(nrElements * 16 + 8), ppBytes(nSize), 4&
                    nSize = nSize + 4
                Case vbDouble, vbDate, vbCurrency ' copy from our array & set next position
                    CopyMemory ptrArray(nrElements * 16 + 8), ppBytes(nSize), 8&
                    nSize = nSize + 8
                Case vbInteger, vbBoolean ' copy from our array & set next position
                    CopyMemory ptrArray(nrElements * 16 + 8), ppBytes(nSize), 2&
                    nSize = nSize + 2
                Case vbByte, vbNull ' copy from our array & set next position
                    ptrArray(nrElements * 16 + 8) = ppBytes(nSize)
                    nSize = nSize + 1
                Case Else
                End Select
            End If
        Next
        ' copy the pointer to the safearray in the host variant's 8-byte data section
        CopyMemory ByVal SafeOffset(VarPtr(hostVariant), 8), VarPtr(safeArray(1)), 4&
    
    Case vbError, vbDataObject, vbUserDefinedType ' unsupported, convert to empty objects
        ReDim ptrArray(0 To nrElements * 4 - 1) ' build a pointer array of pointers
        CopyMemory safeArray(-3), vbObject Or vbArray, 4& ' 4 bytes before array is data type
        CopyMemory safeArray(13), VarPtr(ptrArray(0)), 4&   ' set array pointer (see SafeArray structure above)
        safeArray(5) = 4                                    ' set data element size (see SafeArray structure above)
        CopyMemory safeArray(3), 128, 2&                    ' change features
        ' copy the pointer to the safearray in the host variant's 8-byte data section
        CopyMemory ByVal SafeOffset(VarPtr(hostVariant), 8), VarPtr(safeArray(1)), 4&
    
    Case Else ' numbers,dates,booleans
        If IsArrayEmpty(Not ppBytes) Then Exit Sub         ' null array, nothingto do
        CopyMemory safeArray(13), VarPtr(ppBytes(0)), 4&   ' set array pointer (see SafeArray structure above)
        Select Case vType And Not vbArray
            Case vbLong, vbSingle: safeArray(5) = 4
            Case vbDouble, vbDate, vbCurrency: safeArray(5) = 8
            Case vbInteger, vbBoolean: safeArray(5) = 2
            Case Else: safeArray(5) = 1
        End Select
        ' copy the pointer to the safearray in the host variant's 8-byte data section
        CopyMemory ByVal SafeOffset(VarPtr(hostVariant), 8), VarPtr(safeArray(1)), 4&
    
    End Select
    
    ' tell host what varType it is
    CopyMemory ByVal VarPtr(hostVariant), vType, 2
    tVariant = hostVariant ' this forces VB to copy our ppBytes(), safeArray & pointer
                           ' data, so we can release them
    
    ' We now have tVariant as a valid Variant containing the data from our property bag
    ' and local array(s). But we need to transfer that data to our target and also ensure
    ' both tVariant and HostVariant do not reference the data; otherwise VB will destroy
    ' the data when tVariant & HostVariant are destroyed
    ReDim ppBytes(1 To 16) ' create a blank Variant structure
    ' we don't need the actual variant, we only need the pointer to what the
    ' variant references (object or arrays). The arrays were all local in this routine
    ' but tVariant copied them to system memory and we need the pointer to that memory
    CopyMemory ByVal itemPtr, ByVal SafeOffset(VarPtr(tVariant), 8), 4&
    
    ' now we clean up so that tVariant doesn't try to clear its references
    CopyMemory ByVal VarPtr(tVariant), ppBytes(1), 16&
    ' and also so that HostVariant doesn't try to clear its references either
    ' those references will go out of scope & potential Crash if trying to references
    ' bogus memory addresses
    CopyMemory ByVal VarPtr(hostVariant), ppBytes(1), 16&
    
End Sub
    
Private Function SortNodeLevel(ByVal Index As Long, ByVal SortType As eSortOptions, ByVal bAscending As Boolean, ByVal Children As Boolean, ByVal objPropName As String) As Boolean

    ' PURPOSE: Helper Function. Sorts collection items within a specific branch of the collection.
    ' Called by MoveChildren, SortCollection, SortRecursiveNodes
    ' See SortCollection method for restrictions
    
    ' Note: There are several checks in this routine.  This is to allow the class to sort on various types of colleciton items.
    ' Additionally, each branch/node level of the collection is handled separately. Should one branch not be sortable for
    '    whatevever reasons, then another branch may be, or vice versa.
    
    ' Final Note. This IS the class' sort routine. It may not be optimal or may not be fast enough if you have thousands of items in the collection.
    ' Should this prove to be unacceptable, there is nothing stopping you from creating your own sort algorithm, or better yet, when adding items,
    ' add the item to the collection in the proper sorted position among its siblings.
    
    ' The only other actions that can change the sibling positions in relation to each other is the AddItem, MoveItem and MoveChildren methods.
    
    Dim indexArray() As Long
    Dim cIndex As Long, nrItems As Long
    Dim itemVarType As Long
    
    If Children Then    ' sort child nodes
        If IsArrayEmpty(Not m_Col(Index).Family.Tree) Then
            SortNodeLevel = True
            Exit Function ' no children to sort
        End If
    Else                ' sorting on siblings, get passed node's parent
        Index = m_Col(Index).Parent
    End If
    
    ' here we will use the 8 or so items to determine what item types we are dealing with
    cIndex = m_Col(Index).Family.Tree(treefirst)          ' get first node in branch
    If m_Col(Index).Family.Tree(treelast) = cIndex Then
        SortNodeLevel = True    ' see if the branch only has one child
        Exit Function           ' it does; no need to go further
    End If
    
    Do Until cIndex = 0                                                 ' get first 4 items (if that many)
        itemVarType = itemVarType Or VarType(m_Col(cIndex).Item)
        If nrItems = 7 Then Exit Do
        nrItems = nrItems + 1
        cIndex = m_Col(cIndex).Elder
    Loop
    cIndex = m_Col(Index).Family.Tree(treelast)
    Do Until cIndex = 0                                                 ' get last 4 items (if that many)
        itemVarType = itemVarType Or VarType(m_Col(cIndex).Item)
        If nrItems = 15 Then Exit Do
        nrItems = nrItems + 1
        cIndex = m_Col(cIndex).Younger
    Loop
    
    ' some item types or combination of item types won't be sorted
    If itemVarType = vbObject Then  ' only objects; we might be able to do that
        nrItems = 0
    ElseIf itemVarType = vbString Then  ' only strings; we can do that
        nrItems = 0
    Else
        ' the following are excluded from comparisons
        If Not (itemVarType And vbArray) = vbArray Then
            If Not (itemVarType And vbDataObject) = vbDataObject Then
                If Not (itemVarType And vbError) = vbError Then
                    If Not (itemVarType And vbUserDefinedType) = vbUserDefinedType Then
                        ' if byte,integer,long,double,date,etc then next IF will be True
                        If Not (itemVarType And vbObject) = vbObject Then nrItems = 0
                    End If
                End If
            End If
        End If
    End If
    If Not nrItems = 0 Then Exit Function ' the IF nest above resulted in Not sorting
    
    If itemVarType = vbObject Then
        ' objects can be sorted if user supplied a public property to sort on. Test it
        If objPropName = vbNullString Then Exit Function
        On Error Resume Next
        cIndex = Val(CallByName(m_Col(Index).Item, objPropName, VbGet)) ' testing for an error
        If Err Then
            Err.Clear
            Exit Function                                               ' property name doesn't exist or is not public
        End If
        ' here we will use the 1st 8 or so objects & validate again
        itemVarType = 0
        cIndex = m_Col(Index).Family.Tree(treefirst)                        ' get first node in branch
        Do Until cIndex = 0
            If Not m_Col(cIndex).Item Is Nothing Then
                itemVarType = itemVarType Or VarType(CallByName(m_Col(cIndex).Item, objPropName, VbGet))
                If nrItems = 15 Then Exit Do
                nrItems = nrItems + 1
            End If
            cIndex = m_Col(cIndex).Elder
        Loop
        If nrItems = 0 Then Exit Function   ' all objects are set to Nothing, nothing to sort on
    Else
        objPropName = vbNullString          ' flag for ShellSort routines to not handle item as Objects
    End If
    
    ' now determine best string method to use
    If (itemVarType And vbString) = vbString Then
        ' we will be using text comparisons
        If Not itemVarType = vbString Then SortType = srtAsNumeric      ' other variable types exist in branch. Treat stirngs as numeric
        itemVarType = vbString
    End If
    
    ReDim indexArray(0 To 100)
    nrItems = 0
    cIndex = m_Col(Index).Family.Tree(treefirst)                        ' get first node in branch
    Do Until cIndex = 0                                                 ' loop thru branch adding node indexes to array
        If nrItems = UBound(indexArray) Then ReDim Preserve indexArray(0 To nrItems + 100)
        indexArray(nrItems) = cIndex
        nrItems = nrItems + 1
        cIndex = m_Col(cIndex).Elder                                    ' continue til no more branch nodes
    Loop
    
    On Error GoTo ExitSort                                              ' if any error occurs during sorting, no action is taken
    If itemVarType = vbString Then   ' sort as text
        ShellSortString indexArray(), (SortType = srtAsNumeric), ((SortType And srtTextBinary) = srtTextBinary), objPropName, nrItems - 1
    Else                ' sort as numbers (integer,long,boolean,dates,doubles,etc)
        ShellSortNumbers indexArray(), objPropName, nrItems - 1
    End If

ExitSort:
    If Err Then         ' error trying to sort, no sort action will be taken
        Err.Clear
    Else
        If bAscending Then SortType = relLastChild Else SortType = relFirstChild
        ' FYI: the returned sorted array is always ascending order, so moving each items
        ' one at a time as the lastChild position is an ascending order
        ' & moving them into firstChild position is a descending order
        ' Moving is fast; no array items are actually moved; only their linked lists are changed
        For nrItems = 0 To nrItems - 1
            MoveItem indexArray(nrItems), Index, SortType
        Next
        SortNodeLevel = True
    End If
    Erase indexArray()

End Function

Private Sub ShellSortString(vArray() As Long, ByVal bAsNumeric As Boolean, ByVal bBinaryComnpare As Boolean, ByVal objPropName As String, ByVal nrItems As Long)
  
    ' PURPOSE: Standard ShellSort routine modified to handle strings only. Not the fastest, but pretty fast
    ' Called by SortNodeLevel
    ' vArray is always zero bound
    Dim lLoop As Long
    Dim lHold As Long
    Dim lHValue As Long
    Dim lTemp As Long
    Dim sValue As String, sValue2 As String
    Dim lValue As Long, lValue2 As Long
    Dim lCompare As Long

    Do
        lHValue = 3 * lHValue + 1
    Loop Until lHValue > nrItems
  
    If objPropName = vbNullString Then
        ' maintained in separate loops to prevent overloading with IF statements & slowing the sort processs down
  
        If bAsNumeric Then      ' will be comparing Val(node) against Val(otherNode)
            Do
                lHValue = lHValue \ 3
                For lLoop = lHValue To nrItems
                    lTemp = vArray(lLoop)              ' current index
                    lValue = Val(m_Col(lTemp).Item)    ' value at current index
                    lHold = lLoop
                    
                    Do While Val(m_Col(vArray(lHold - lHValue)).Item) > lValue
                        vArray(lHold) = vArray(lHold - lHValue)
                        lHold = lHold - lHValue
                        If lHold < lHValue Then Exit Do
                    Loop
                    vArray(lHold) = lTemp
                    
                Next
            Loop Until lHValue = 0&
            
        Else                ' will be comparing raw string to raw string
            If bBinaryComnpare = True Then lCompare = vbBinaryCompare Else lCompare = vbTextCompare
            Do
                lHValue = lHValue \ 3
                For lLoop = lHValue To nrItems
                    lTemp = vArray(lLoop)                   ' current index
                    sValue = CStr(m_Col(lTemp).Item)        ' value at current index
                    lHold = lLoop
                    
                    Do While StrComp(CStr(m_Col(vArray(lHold - lHValue)).Item), sValue, lCompare) = 1
                        vArray(lHold) = vArray(lHold - lHValue)
                        lHold = lHold - lHValue
                        If lHold < lHValue Then Exit Do
                    Loop
                    vArray(lHold) = lTemp
                
                Next
            Loop Until lHValue = 0&
        End If

    Else    ' this is the same as above section, but applies to objects only
    
        If bAsNumeric Then
            Do
                lHValue = lHValue \ 3
                For lLoop = lHValue To nrItems
                    lTemp = vArray(lLoop)
                    ' use VB's CallByName method to get the property of the class/object
                    If m_Col(lTemp).Item Is Nothing Then
                        lValue = 0&
                    Else
                        lValue = Val(CallByName(m_Col(lTemp).Item, objPropName, VbGet))
                    End If
                    lHold = lLoop
                    
                    Do
                        If m_Col(vArray(lHold - lHValue)).Item Is Nothing Then
                            lValue2 = 0&
                        Else
                            lValue2 = Val(CallByName(m_Col(vArray(lHold - lHValue)).Item, objPropName, VbGet))
                        End If
                        If Not lValue2 > lValue Then Exit Do
                        vArray(lHold) = vArray(lHold - lHValue)
                        lHold = lHold - lHValue
                        If lHold < lHValue Then Exit Do
                    Loop
                    vArray(lHold) = lTemp
                    
                Next
            Loop Until lHValue = 0&
            
        Else
            If bBinaryComnpare = True Then lCompare = vbBinaryCompare Else lCompare = vbTextCompare
            Do
                lHValue = lHValue \ 3
                For lLoop = lHValue To nrItems
                    lTemp = vArray(lLoop)
                    ' use VB's CallByName method to get the property of the class/object
                    If m_Col(lTemp).Item Is Nothing Then
                        sValue = vbNullString
                    Else
                        sValue = CStr(CallByName(m_Col(lTemp).Item, objPropName, VbGet))
                    End If
                    lHold = lLoop
                    
                    Do
                        If m_Col(vArray(lHold - lHValue)).Item Is Nothing Then
                            sValue2 = vbNullString
                        Else
                            sValue2 = CStr(CallByName(m_Col(vArray(lHold - lHValue)).Item, objPropName, VbGet))
                        End If
                        If Not StrComp(sValue2, sValue, lCompare) = 1 Then Exit Do
                        vArray(lHold) = vArray(lHold - lHValue)
                        lHold = lHold - lHValue
                        If lHold < lHValue Then Exit Do
                    Loop
                    vArray(lHold) = lTemp
                
                Next
            Loop Until lHValue = 0&
        End If
    
    End If

End Sub

Private Sub ShellSortNumbers(vArray() As Long, ByVal objPropName As String, ByVal nrItems As Long)
  
    ' PURPOSE: Standard ShellSort routine modified to handle numbers/dates. Not the fastest, but pretty fast
    ' Called by SortNodeLevel
    ' vArray is always zero bound
    Dim lLoop As Long
    Dim lHold As Long
    Dim lHValue As Long
    Dim lTemp As Long
    Dim lValue As Variant
    
    Do
        lHValue = 3 * lHValue + 1
    Loop Until lHValue > nrItems
  
    If objPropName = vbNullString Then
        ' maintained in separate loops to prevent overloading with IF statements & slowing the sort processs down
  
        Do
            lHValue = lHValue \ 3&
            For lLoop = lHValue To nrItems
          
                lTemp = vArray(lLoop)
                lValue = m_Col(lTemp).Item
                lHold = lLoop
                Do While m_Col(vArray(lHold - lHValue)).Item > lValue
                    vArray(lHold) = vArray(lHold - lHValue)
                    lHold = lHold - lHValue
                    If lHold < lHValue Then Exit Do
                Loop
                vArray(lHold) = lTemp
            Next
        Loop Until lHValue = 0&
    
    Else    ' this is the same as above section, but applies to objects only
    
        Do
            lHValue = lHValue \ 3&
            For lLoop = lHValue To nrItems
          
                lTemp = vArray(lLoop)
                ' use VB's CallByName method to get the property of the class/object
                If m_Col(lTemp).Item Is Nothing Then
                    lValue = 0
                Else
                    lValue = CallByName(m_Col(lTemp).Item, objPropName, VbGet)
                End If
                lHold = lLoop
                Do While CallByName(m_Col(vArray(lHold - lHValue)).Item, objPropName, VbGet) > lValue
                    vArray(lHold) = vArray(lHold - lHValue)
                    lHold = lHold - lHValue
                    If lHold < lHValue Then Exit Do
                Loop
                vArray(lHold) = lTemp
            Next
        Loop Until lHValue = 0&
    
    End If

End Sub

Private Function GetChildInsertPosition(ByVal Index As Long, ByVal ParentIndex As Long, ByVal SortType As eSortOptions, ByVal Ascending As Boolean) As Long

    ' PURPOSE: Compare passed Index node with its projected siblings to find its sorted
    '   position among its new siblings.
    ' Called by: AddItem, MoveItem
    
    ' Return Value is the sibling Index to insert the new index Before
    ' If return value is zero, then new index is to be inserted at end of sibling chain

    Dim cIndex As Long, valType As Long, cmpType As Long
    
    If IsArrayEmpty(Not m_Col(ParentIndex).Family.Tree) Then Exit Function
    
    valType = VarType(m_Col(Index).Item)    ' get new item's variable type
    
    Select Case valType
    Case vbObject, vbError, vbDataObject, vbUserDefinedType
        Exit Function ' must sort these using CollectionSort
    Case Else
        If (valType And vbArray) = vbArray Then Exit Function ' won't sort items containing arrays
        If valType = vbString Then              ' new item is string, so sort based on Strings
            If SortType = srtAsNumeric Then     ' will we be sorting on Val(item)?
                valType = vbNull
            ElseIf Not SortType = srtTextBinary Then ' will sorting be case sensitive?
                cmpType = vbTextCompare
            End If
        Else
            valType = vbNull                    ' numbers
        End If
    End Select
    
    On Error GoTo ExitRoutine   ' should error occur in comparisons; item is added to end of family tree
    
    cIndex = m_Col(ParentIndex).Family.Tree(treefirst)
    Do Until cIndex = 0
        If valType = vbNull Then   ' numeric comparison
            Select Case Val(m_Col(Index).Item) - Val(m_Col(cIndex).Item)
            Case 0  ' equal values
                GetChildInsertPosition = cIndex
                Exit Do
            Case Is > 0 ' new value > cindex value
                If Not Ascending Then
                    GetChildInsertPosition = cIndex
                    Exit Do
                End If
            Case Else
                If Ascending = True Then
                    GetChildInsertPosition = cIndex
                    Exit Do
                End If
            End Select
        Else ' case/non-case sensitive comparision
            Select Case StrComp(m_Col(Index).Item, CStr(m_Col(cIndex).Item), cmpType)
            Case 0  ' equal
                GetChildInsertPosition = cIndex
                Exit Do
            Case 1 ' new value > cindex value
                If Not Ascending Then
                    GetChildInsertPosition = cIndex
                    Exit Do
                End If
            Case Else
                If Ascending = True Then
                    GetChildInsertPosition = cIndex
                    Exit Do
                End If
            End Select
        End If
        cIndex = m_Col(cIndex).Elder    ' get next sibling
    Loop
    
ExitRoutine:
If Err Then Err.Clear
End Function


Private Sub Class_Initialize()
    Call ClearCollection
    m_Buffer = 10
    m_CompareType = keyCaseInsensitive
    Randomize Timer
End Sub

Private Sub Class_Terminate()
    Erase m_Col
    Erase m_Keys
End Sub


