VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   5595
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   5595
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Allow multiple Item instances"
      Height          =   465
      Left            =   3720
      TabIndex        =   12
      Top             =   6480
      Width           =   1590
   End
   Begin VB.ListBox lstMainLevel 
      Height          =   1620
      Index           =   1
      ItemData        =   "Form2.frx":0000
      Left            =   2880
      List            =   "Form2.frx":0002
      TabIndex        =   8
      Top             =   825
      Width           =   2430
   End
   Begin VB.ListBox lstGroupItems 
      Height          =   3570
      Index           =   1
      ItemData        =   "Form2.frx":0004
      Left            =   2865
      List            =   "Form2.frx":0006
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   2805
      Width           =   2460
   End
   Begin VB.CommandButton cmdSaveCollection 
      Caption         =   "Re-Load Collection"
      Height          =   510
      Index           =   1
      Left            =   1920
      TabIndex        =   5
      Top             =   6435
      Width           =   1725
   End
   Begin VB.CommandButton cmdSaveCollection 
      Caption         =   "Save Collection"
      Height          =   510
      Index           =   0
      Left            =   165
      TabIndex        =   4
      Top             =   6420
      Width           =   1725
   End
   Begin VB.ListBox lstGroupItems 
      Height          =   3570
      Index           =   0
      ItemData        =   "Form2.frx":0008
      Left            =   165
      List            =   "Form2.frx":000A
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   2790
      Width           =   2460
   End
   Begin VB.ListBox lstMainLevel 
      Height          =   1620
      Index           =   0
      ItemData        =   "Form2.frx":000C
      Left            =   180
      List            =   "Form2.frx":000E
      TabIndex        =   0
      Top             =   810
      Width           =   2430
   End
   Begin VB.Label Label1 
      Caption         =   "EXAMPLE: Collection of Objects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   2955
      TabIndex        =   11
      Top             =   180
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Main Level Nodes (Categories)"
      Height          =   285
      Index           =   4
      Left            =   2925
      TabIndex        =   10
      Top             =   555
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Category Items (Right Click)"
      Height          =   285
      Index           =   3
      Left            =   2940
      TabIndex        =   9
      Top             =   2520
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "EXAMPLE: Collection of Strings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   255
      TabIndex        =   6
      Top             =   165
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Category Items (Right Click)"
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   2505
      Width           =   2340
   End
   Begin VB.Label Label1 
      Caption         =   "Main Level Nodes (Categories)"
      Height          =   285
      Index           =   0
      Left            =   225
      TabIndex        =   2
      Top             =   540
      Width           =   2340
   End
   Begin VB.Menu mnuMain 
      Caption         =   "mnuGeneric"
      Visible         =   0   'False
      Begin VB.Menu mnuGenAdd 
         Caption         =   ""
         Index           =   0
      End
   End
   Begin VB.Menu mnuCats 
      Caption         =   "Categories"
      Visible         =   0   'False
      Begin VB.Menu mnuDelete 
         Caption         =   "Remove Item"
         Index           =   0
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Start Over"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' A sample project for testing the cNodeCollection class.
' Read the remarks in the class module for detailed information.
' Since the collection can contain both keyed and non-keyed items, identifying
' non-keyed items can be tricky. This sample project focuses on a non-keyed collection

' Also included is a RTF document overviewing all the public methods/properties

' Once enough feedback has been recieved that identify bugs and recommended improvements,
' I will add to this project & repost it, some variations of this class.
' Those variations will most likely be.....
'   1) Flat-collection, like VB's Collection object. No heirarchy, about 1/2 of the class will disappear
'   2) Object-based collection, not Variant based
'   3) String-based collection, not Variant based

' TODO. Add TLB to all I_Enum implementation (i.e., "For Each" usage)

' As you are playing, keep in mind that I am allowing items to be added infinitely
' number of times to any category.  If I wanted to prevent more than one instance
' of an item, I could do that easily by either keying each item and then checking
' to see if the key already exists before trying to add/move it to a category.
' The check box on the form, allows toggling the capability

' Many of us a very familiar with VB's collection and TreeView's collection
' that use keys to identify items, this sample project will focus on non-keyed items.
' Keyed items are easy. The main level categories items are keyed just to show how
' one would add a keyed item to the collection. However, when a Generic's child item
' is copied to other categories, the added item is not keyed.  When a child object is
' moved from one category to another, the unkeyed item remains unkeyed.

' Sounds relatively simple, until you realize that having unkeyed items in a collection
' can be rather painful.  Therefore, the collection has several methods and properties
' to allow you to navigate and reference the unkeyed collection items. In practice,
' I would suggest keying, at the very minimum, parent nodes.


' SPECIAL NOTES REGARDING LISTBOXES and COLLECTION ITEMS.
' -------------------------------------------------------
' Multi-Level collection items....
' Think about it. Adding collection items to a listbox is tricky. How do you reference
'   a collection item from the listbox?  Key? only if the keys are numeric so you can
'   add them to the listbox's ItemData property. What about list item text (collection
'   item value) as a key? That will work only if every item is keyed and a collection
'   item value will never be duplicated among other items in the collection.

' So what if the text can't be the key, because the text may be a value in more
' than one collection item?  What if your keys are alphanumeric and you can't add them
' to the listbox's ItemData property? What if your items are not keyed?

' Suggestion forthcoming.

' This sample project focuses on using a listbox to display the collection item values.
' The listbox contains unkeyed child items from first child to last child
' (siblings amongst themselves), in order, within the collection. The collection class
' has a function: SiblingOffsetIndex that will get x number of siblings from the
' first or last sibling in its family tree. This is ideal I think.

' However, I DO NOT set the list box Sorted property to true. If I want the list
' sorted, I will ask the collection to sort it. Otherwise the order in the listbox
' probably won't be the same as the order within the collection. If I
' did set the listbox sorted property to True, I could store the sibling
' indexes into the listbox's ItemData property as I added each item to the listbox.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private WithEvents myCol As cNodeCollection
Attribute myCol.VB_VarHelpID = -1
'^^ WithEvents only applies if you plan on saving your collection or loading a
'   previously saved collection.  And even at that, only if your collection will
'   contain Objects (i.e., Classes). Otherwise, you do not need to declare WithEvents.

Private Sub Check1_Click()

    If Check1.Value = 0 Then
        ' don't allow multiple instances of items within a category
        ' FYI: The easiest way to prevent multiple instances is to key items and
        '      simply test for the key:  myCol.KeyExists(Key)
        
        If MsgBox("Any existing multiple instances will remain. Reset the collection?", vbYesNo + vbQuestion) = vbYes Then
            InitializeCollection
        End If
    Else
        ' allow multiple items.  When you add items to a category, they will not
        ' be keyed; why worry about creating unique keys for items that can be
        ' duplicated throughout the collection?
    End If
    
End Sub

Private Sub Form_Load()
    Randomize Timer
    InitializeCollection
End Sub

Private Sub Form_Unload(Cancel As Integer)
    RemoveExtraMenuItems
End Sub

Private Sub RemoveExtraMenuItems()
    ' popup menu items are dynamically created. Remove any that were added
    Dim mnu As Long
    For mnu = mnuGenAdd.UBound To 1 Step -1
        Unload mnuGenAdd(mnu)
    Next
End Sub


Private Sub lstGroupItems_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Set up the popup menu. Allow items to be copied , moved or sorted
    If Button = vbRightButton Then
        
        Dim mnu As Long, nrMnuItems As Long
        
        If lstGroupItems(Index).SelCount = 0 Then
            MsgBox "First select at least one item in the listing", vbInformation + vbOKOnly
            Exit Sub
        End If
        
        RemoveExtraMenuItems
        
        If lstMainLevel(Index).ListCount > 1 Then ' else all non-Generic categories were deleted
            
            mnuMain.Tag = Index ' identify which list this menu is activated on
            
            If lstMainLevel(Index).ListIndex = 0 Then   ' Generic category
                ' we allow Add Item(s) only and Sorting, no Delete or Move
                mnuGenAdd(0).Caption = "Add Item(s) to " & lstMainLevel(Index).List(1)
                For mnu = 2 To lstMainLevel(Index).ListCount - 1
                    Load mnuGenAdd(mnu - 1)
                    mnuGenAdd(mnu - 1).Caption = "Add Item(s) to " & lstMainLevel(Index).List(mnu)
                Next
                mnu = mnuGenAdd.UBound + 1
                Load mnuGenAdd(mnu)
                mnuGenAdd(mnu).Caption = "-"
                Load mnuGenAdd(mnu + 1)
                mnuGenAdd(mnu + 1).Caption = "Sort Ascending"
                Load mnuGenAdd(mnu + 2)
                mnuGenAdd(mnu + 2).Caption = "Sort Descending"
                PopupMenu mnuMain
                
            ElseIf lstGroupItems(Index).SelCount > 0 Then
                ' if any items are selected, we will allow
                ' Remove, Remove All, and Move
                mnuGenAdd(0).Caption = "Remove Item(s)"
                Load mnuGenAdd(1)
                mnuGenAdd(1).Caption = "Remove All Items"
                nrMnuItems = 1
                For mnu = 1 To lstMainLevel(Index).ListCount - 1
                    If mnu <> lstMainLevel(Index).ListIndex Then
                        nrMnuItems = nrMnuItems + 1
                        Load mnuGenAdd(nrMnuItems)
                        mnuGenAdd(nrMnuItems).Caption = "Move Item(s) to " & lstMainLevel(Index).List(mnu)
                    End If
                Next
                Load mnuGenAdd(nrMnuItems + 1)
                mnuGenAdd(nrMnuItems + 1).Caption = "-"
                Load mnuGenAdd(nrMnuItems + 2)
                mnuGenAdd(nrMnuItems + 2).Caption = "Sort Ascending"
                Load mnuGenAdd(nrMnuItems + 3)
                mnuGenAdd(nrMnuItems + 3).Caption = "Sort Descending"
                PopupMenu mnuMain
                ' refresh the list
                Call lstMainLevel_Click(Index)
            End If
        End If
    End If

End Sub

Private Sub lstMainLevel_Click(Index As Integer)

    ' fills in the category child items under the selected category
    
    If lstMainLevel(Index).ListCount > 0 Then
        
        Dim nodeIndex As Long
        
        lstGroupItems(Index).Clear
        
        ' which is the first child of the selected category?
        If Index = 0 Then
            nodeIndex = myCol.FirstChild("main" & lstMainLevel(Index).Text)
        Else
            nodeIndex = myCol.FirstChild("obj" & lstMainLevel(Index).Text)
        End If
        
        ' loop thru each child, adding the children to the items list
        
        ' I know that listbox(0) are strings, & listbox(1) are objects
        ' alternatively, I could test items like:
        '   If VarType(myCol.Item(nodeIndex))=vbObject Then add to list(1) else add to list(0)
        Do Until nodeIndex = 0
            If Index = 0 Then
                lstGroupItems(Index).AddItem myCol.Item(nodeIndex)
            Else    ' objects
                lstGroupItems(Index).AddItem myCol.Item(nodeIndex).Name ' << testClass property
            End If
            nodeIndex = myCol.NextSibling(nodeIndex)
        Loop
        
        If lstGroupItems(Index).ListCount > 0 Then
            lstGroupItems(Index).ListIndex = 0  ' select first item in the list
            lstGroupItems(Index).Selected(0) = True
        End If
        
    End If
End Sub


Private Sub InitializeCollection()

    ' create a new collection, adding categories and children
    Set myCol = New cNodeCollection
    
    myCol.AddItem "Generic", "mainGeneric"      ' 4 keyed main level nodes, just strings
    myCol.AddItem "Animals", "mainAnimals"
    myCol.AddItem "Minerals", "mainMinerals"
    myCol.AddItem "Plants", "mainPlants"
    
    ' add each as a child to mainGeneric, items are unkeyed here
    myCol.AddItem "Tiger", , "mainGeneric", relLastChild
    myCol.AddItem "Elephant", , "mainGeneric", relLastChild
    myCol.AddItem "Wolf", , "mainGeneric", relLastChild
    myCol.AddItem "Aardvark", , "mainGeneric", relLastChild
    myCol.AddItem "Shark", , "mainGeneric", relLastChild
    myCol.AddItem "Hawk", , "mainGeneric", relLastChild
    myCol.AddItem "Gold", , "mainGeneric", relLastChild
    myCol.AddItem "Silver", , "mainGeneric", relLastChild
    myCol.AddItem "Iron", , "mainGeneric", relLastChild
    myCol.AddItem "Zinc", , "mainGeneric", relLastChild
    myCol.AddItem "Copper", , "mainGeneric", relLastChild
    myCol.AddItem "Onyx", , "mainGeneric", relLastChild
    myCol.AddItem "Sunflower", , "mainGeneric", relLastChild
    myCol.AddItem "Oak", , "mainGeneric", relLastChild
    myCol.AddItem "Spruce", , "mainGeneric", relLastChild
    myCol.AddItem "Poison Ivy", , "mainGeneric", relLastChild
    myCol.AddItem "Clover", , "mainGeneric", relLastChild
    myCol.AddItem "Algae", , "mainGeneric", relLastChild
    
    Dim tClass As classTestClass, x As Long, NodeID As Long
    
    ' add 4 keyed main level nodes that are objects/classes
    Set tClass = New classTestClass
    tClass.Name = "Generic": tClass.Index = 99
    myCol.AddItem tClass, "objGeneric"
    
    Set tClass = New classTestClass
    tClass.Name = "Animals": tClass.Index = 98
    myCol.AddItem tClass, "objAnimals"
    
    Set tClass = New classTestClass
    tClass.Name = "Minerals": tClass.Index = 109
    myCol.AddItem tClass, "objMinerals"
    
    Set tClass = New classTestClass
    tClass.Name = "Plants": tClass.Index = 89
    myCol.AddItem tClass, "A Real Long & Unnecessary Key That We Will change Later *****"
    
    ' now add the items as child classes to the class objGeneric
    
    ' get first child of mainGeneric
    NodeID = myCol.LastChild("mainGeneric")
    Do Until NodeID = 0
        ' loop thru each of the children, create a new class and adding it
        ' as a child to the objGeneric node, unkeyed children
        Set tClass = New classTestClass
        tClass.Name = myCol.Item(NodeID)
        tClass.Index = NodeID + 200     ' n/a. Not used in anything
        myCol.AddItem tClass, "obj" & tClass.Name, "objGeneric", relLastChild
        NodeID = myCol.PreviousSibling(NodeID)
    Loop
    
    ' example of using Key property to change a key
    myCol.Key("A Real Long & Unnecessary Key That We Will change Later *****") = "objPlants"
    
    
    RefreshCategories
    
End Sub

Private Sub RefreshCategories()

    ' Refreshes the categories list boxes
    
    Dim nodeIndex As Long
    lstMainLevel(0).Clear
    lstMainLevel(1).Clear
    
    ' get first node of the collection
    nodeIndex = myCol.FirstSibling(0)
    ' FYI: NodeIndex = myCol.FirstChild(0) will return same thing
    
    ' loop thru each sibling and add it to the categories list boxes
    Do Until nodeIndex = 0
        If Left$(myCol.Key(nodeIndex), 3) = "obj" Then  ' classes vs strings
            lstMainLevel(1).AddItem myCol.Item(nodeIndex).Name ' << testClass property
        Else
            lstMainLevel(0).AddItem myCol.Item(nodeIndex)
        End If
        nodeIndex = myCol.NextSibling(nodeIndex)    ' get next sibling
    Loop
    lstMainLevel(0).ListIndex = 0
    lstMainLevel(1).ListIndex = 0
    Caption = "Collection Count: " & myCol.ItemCount

End Sub

Private Sub lstMainLevel_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    ' Set up main category popup menu
    ' Don't allow deleting the "Generic" category
        
    If Button = vbRightButton Then
        If lstMainLevel(Index).ListIndex = 0 Then
            mnuDelete(0).Enabled = False
        Else
            mnuDelete(0).Enabled = True
        End If
        mnuMain.Tag = Index
        PopupMenu mnuCats
    End If

End Sub

Private Sub mnuDelete_Click(Index As Integer)

    ' Delete menu from category listbox was clicked
    ' Remove child items from the parent node
    If Index = 0 Then
        ' removing a parent level node, removes all of its descendants too
        If Val(mnuMain.Tag) = 0 Then
            myCol.RemoveItem "main" & lstMainLevel(0).Text
        Else
            myCol.RemoveItem "obj" & lstMainLevel(1).Text
        End If
        RefreshCategories
    Else    ' Start Over menu item clicked
        InitializeCollection
    End If
End Sub

Private Sub mnuGenAdd_Click(Index As Integer)
    
    ' Either Add, Remove, or Move menu items were clicked from the Items listboxes
    
    ' Since this is the workhorse function, I will separate it into individual
    ' routines so that you can follow it easier if you wish to
    
    
    Dim sKeyFrom As String, sKeyTo As String, bAsc As Boolean
    Dim LstID As Integer, sMainCat As String
    
    LstID = (Val(mnuMain.Tag))
    
    If mnuGenAdd(Index).Caption = "Remove Item(s)" Then
        Call DeleteCollectionItems(LstID, vbNullString)
        Call lstMainLevel_Click(LstID)
    
    ElseIf mnuGenAdd(Index).Caption = "Remove All Items" Then
        ' remove all children from the highlighted Category
        If LstID = 0 Then sMainCat = "main" Else sMainCat = "obj"
        
        myCol.RemoveChildren sMainCat & lstMainLevel(LstID).Text
        
        Call lstMainLevel_Click(LstID)
    
    ElseIf Left$(mnuGenAdd(Index).Caption, 4) = "Sort" Then
        
        ' sorting vs add/move
        If LstID = 0 Then sMainCat = "main" Else sMainCat = "obj"
        sKeyFrom = sMainCat & lstMainLevel(LstID).Text  ' key for "From" category
        
        If Not Index = mnuGenAdd.UBound Then bAsc = True ' sorting descending order.
        
        If myCol.SortCollection(sMainCat & lstMainLevel(LstID).Text, brchChildren, srtTextText, bAsc, "Name") = True Then
            ' ^^ Note: Last parameter is used when collection contains objects. The parameter is the name of a Public Propety that exists for the
            '    objects in your collection.  If your collection contains multiple types of objects that don't have a common property name you can
            '    sort on, then you will need to sort the indexes yourself and simply move the items in proper sort order :: myCol.MoveItem(...)
            Call lstMainLevel_Click(LstID)
        
        Else
        
            MsgBox "Failed to Sort the Collection", vbInformation + vbOKOnly
            
        End If
        ' About sorting. Individual family tree branches can be sorted or the entire
        ' collection can be sorted. There are 3 options to choose which level to sort.
        
        
    ElseIf Left$(mnuGenAdd(Index).Caption, 8) = "Add Item" Then ' adding
    
        ' pass off to separate routine, include listbox index & category to add to
        Call AppendCollectionItems(LstID, Mid$(mnuGenAdd(Index).Caption, 16))

    Else    ' moving
    
        ' pass off to separate routine, include listbox index & category to move to
        Call MoveCollectionItems(LstID, Mid$(mnuGenAdd(Index).Caption, 17))
        Call lstMainLevel_Click(LstID)
        
    End If
    
    Caption = "Collection Count: " & myCol.ItemCount
    
End Sub

Private Sub cmdSaveCollection_Click(Index As Integer)

    ' The collection class has the ability to cache the entire collection into a byte array.
    ' There are exceptions.... The following types of items must be serialized by you and to
    '   allow you to serialize the items, the WithEvents above are called. See the class for more.
    
    '   Those items that require serialization by the user are:
    '   - Objects (i.e., Classes, Controls, UDTs, etc).
    '       -- Example of serializing a Class is given in classTestClass
    '   - Arrays of Variants and Objects
    
    '   All other items, arrayed or not, can be serialized by the collection class
    '   Exception: Arrays containing more than 10 dimensions will not be saved
    

    Dim fNR As Integer, sFile As String
    Dim colBytes() As Byte, nodeIndex As Long
    
    If Index = 0 Then ' saving
    
        If Not Len(Dir(App.Path & "\SavedCol.dat")) = 0 Then
            Kill App.Path & "\SavedCol.dat" ' remove if exists
        End If
         ' call method to convert collection to byte array
        If myCol.SaveCollection(colBytes()) = True Then
            '^^ Note: If serialization is required, the WithEvents events will be called at this point
            fNR = FreeFile()
            Open App.Path & "\SavedCol.dat" For Binary As #fNR
            Put #fNR, , colBytes
            Close #fNR
        Else
            MsgBox "Failed to save Collection"
        End If
        
    Else    ' else loading
    
        If Len(Dir(App.Path & "\SavedCol.dat")) = 0 Then
            MsgBox "No saved collection found", vbInformation + vbOKOnly
        
        Else
        
            lstGroupItems(0).Clear
            lstGroupItems(1).Clear
            lstMainLevel(0).Clear
            lstMainLevel(1).Clear
        
            fNR = FreeFile()
            Open App.Path & "\SavedCol.dat" For Binary As #fNR
            ReDim colBytes(0 To LOF(fNR) - 1)
            Get #fNR, , colBytes
            Close #fNR
            
            Set myCol = Nothing ' just to show you we are not re-reading same colleciton object
            Set myCol = New cNodeCollection
            
            ' call method to create collection from byte array
            If myCol.LoadCollection(colBytes()) = True Then
                '^^ Note: If any items were serialized by you, the WithEvents events
                '   would be called at this point so you can un-serialize them
            
                nodeIndex = myCol.FirstSibling(0)
                Do Until nodeIndex = 0
                    If Left$(myCol.Key(nodeIndex), 3) = "obj" Then
                        lstMainLevel(1).AddItem myCol.Item(nodeIndex).Name ' << testClass property
                    Else
                        lstMainLevel(0).AddItem myCol.Item(nodeIndex)
                    End If
                    nodeIndex = myCol.NextSibling(nodeIndex)
                Loop
                
                If Not lstMainLevel(0).ListCount = 0 Then lstMainLevel(0).ListIndex = 0
                If Not lstMainLevel(1).ListCount = 0 Then lstMainLevel(1).ListIndex = 0
                Caption = "Collection Count: " & myCol.ItemCount
            
            Else
            
                MsgBox "Failed to Load the Collection", vbInformation + vbOKOnly, "Corrupt or Invalid Collection"
                InitializeCollection
                
            End If
            
        End If
        
    
    End If
End Sub

Private Sub DeleteCollectionItems(LstID As Integer, TargetCategory As String)

    Dim fromKey As String, x As Long
    
    If LstID = 0 Then
        fromKey = "main" & lstMainLevel(LstID).Text ' key for "From" category
    Else
        fromKey = "obj" & lstMainLevel(LstID).Text ' key for "From" category
    End If
    
    If lstGroupItems(LstID).SelCount = lstGroupItems(LstID).ListCount Then
        ' deleting every child; this is rather easy. One function call
        myCol.RemoveChildren fromKey
        
    Else
    
        ' The collection may be keyed or may not be. Doesn't matter, we have no
        ' way of storing alphanumeric keys with the listbox short of creating
        ' a separate cross-reference list.
        
        ' So we will use myCol.SiblingOffsetIndex for navigating
        ' collection items displayed in a list box.
        
        For x = lstGroupItems(LstID).ListCount - 1 To 0 Step -1
            ' just like listboxes and most other collection listings
            ' delete from collection in reverse order
        
            If lstGroupItems(LstID).Selected(x) = True Then
                ' delete the selected item. Listbox items may be unkeyed,
                ' so which collection item is it?
                
                ' the SiblingOffsetIndex function can get sibling offset from
                ' the first sibling or last sibling in the family tree
                myCol.RemoveItem myCol.SiblingOffsetIndex(fromKey, x)
                
            End If
            
        Next

        ' Note: deleting keyed items is far simpler
        ' myCol.DeleteItem keyedItem
    End If
    
End Sub

Private Sub AppendCollectionItems(LstID As Integer, TargetCategory As String)

    Dim toKey As String, fromKey As String, NewKey As String
    Dim x As Long, tClass As classTestClass
    
    If LstID = 0 Then
        fromKey = "main" & lstMainLevel(LstID).Text ' key for "From" category
        toKey = "main" & TargetCategory             ' key for "To" category
    Else
        fromKey = "obj" & lstMainLevel(LstID).Text ' key for "From" category
        toKey = "obj" & TargetCategory              ' key for "To" category
    End If
    
    For x = 0 To lstGroupItems(LstID).ListCount - 1
        If lstGroupItems(LstID).Selected(x) = True Then
            ' add the selected item. If items are unkeyed so which
            ' collection item is it?
            
            If Check1 = 0 Then ' single instances only, we will key items as we add them
            
                ' verify the item doesn't already exist
                
                ' Create key for the selected item
                If LstID = 0 Then
                    NewKey = "String" & lstGroupItems(LstID).List(x)
                Else
                    NewKey = "Class" & lstGroupItems(LstID).List(x)
                End If
                
                If myCol.KeyExists(NewKey) = True Then Exit Sub
                   
            Else    ' allow multiple copies; unkeyed
                    ' newKey will be vbNullString
            End If
                
            If LstID = 0 Then ' strings else objects
                myCol.AddItem lstGroupItems(LstID).List(x), NewKey, toKey, relLastChild
                    
            Else
                Set tClass = New classTestClass
                tClass.Name = lstGroupItems(LstID).List(x)
                tClass.Index = Int(Rnd * vbWhite)
                myCol.AddItem tClass, NewKey, toKey, relLastChild
            End If
            
        End If
    Next

End Sub

Private Sub MoveCollectionItems(LstID As Integer, TargetCategory As String)

    Dim toKey As String, fromKey As String, x As Long, Offset As Long
    
    If LstID = 0 Then
        fromKey = "main" & lstMainLevel(LstID).Text ' key for "From" category
        toKey = "main" & TargetCategory             ' key for "To" category
    Else
        fromKey = "obj" & lstMainLevel(LstID).Text ' key for "From" category
        toKey = "obj" & TargetCategory              ' key for "To" category
    End If
    
    If lstGroupItems(LstID).SelCount = lstGroupItems(LstID).ListCount Then
        ' moving every child; this is rather easy. One function call
        
        myCol.MoveChildren fromKey, toKey, relLastChild
        
    Else
    
        ' The collection may be keyed or may not be. Doesn't matter, we have no
        ' way of storing alphanumeric keys with the listbox short of creating
        ' a separate cross-reference list.
        
        ' So we will use myCol.SiblingOffsetIndex for navigating unkeyed
        ' collection items displayed in a list box.
        
        ' just like listboxes and most other collection listings
        ' move from the collection in reverse order.
        ' This has one draw back...
        
        ' The moved items get into the new category in reverse order.
        ' I don't want this, but if I move from top to bottom, then
        ' the item indexes in the listbox are not the same as the collection
        ' that I'm moving from because any previously moved items no longer
        ' exist in that collection category. Understand?
        
        ' Therefore, I'll just use a simple offset algorithm and move
        ' from top to bottom so the moved items are maintained in same order
        ' at there destination category. Again, if items were keyed we could
        ' simply move items by its key value vs a relative offset
        
        For x = 0 To lstGroupItems(LstID).ListCount - 1
        
        ' If that doesn't matter, then fine.  If it does, look at the alternative
        ' approach after Notes at bottom of this routine
        
            If lstGroupItems(LstID).Selected(x) = True Then
                ' move the selected item. Listbox items may be unkeyed,
                ' so which collection item is it?
                
                ' the SiblingOffsetIndex function can get sibling offset from
                ' the first sibling or last sibling in the family tree
                myCol.MoveItem myCol.SiblingOffsetIndex(fromKey, x - Offset), toKey, relLastChild
                
                ' moved item from collection, adjust offset
                Offset = Offset + 1
                
            End If
            
        Next

        ' Note: moving keyed items is far simpler
        ' myCol.MoveItem keyedItem, toTargetKeyedItem, new relative position
        
    End If
    
End Sub


Private Sub myCol_SerializeObject(ByVal Serialize As Boolean, collectionObject As Object, DataArray() As Byte, ObjectID As String)

    ' With this test form, I am collecting classTestClass objects.
    ' When saving the collection, we are required to serialize objects and arrays
    ' of objects/variants. This function is called when the collection class is
    ' attempting to save an object or array of objects/variants.
    
    'If Serialize=True :: You are to serialize the passed object into the DataArray()
    '    collectionObject. A non-Nothing object within your collection.
    '        -- Test the type object, if needed, by using TypeOf(collectionObject) Is classWhatever
    '    DataArray(). An empty array you are to fill with the contents of the serailized object
    '        -- The serialized object in DataArray can be multidimensional & any LBound.
    '    ObjectID (Optional) :: a name you provide that uniquely distinguishes the kind of object it is. This is usually a Class Type.
    '        -- When you get the array back for deserialization, you may or many not included flags to indicate which class/object
    '            type it was created from (i.e., "cMyClass"). Therefore knowing which class to create from array may be impossible
    '
    'If Serialize=False :: You are to deserialize the DataArray() into the passed collectionObject
    '    collectionObject. An empty object to be set from the deserailized array
    '        -- i.e, deserialize DataArray into a new classWhatever, then Set collectionObject=classWhatever
    '    DataArray(). A populated array you are to use to create a new object
    '        -- The DataArray will be same size, dimensions, and have same LBound/UBound you passed when Serialize=True
    '    ObjectID (Optional) :: If the ObjectID was provided when Serialize=True, the the value you provided else vbNullString
    
    Dim tClass As classTestClass
    
    If Serialize Then
    ' Note that the classTestClass' SaveClass method is Friend. Since it is Friend,
    '   you must reference it specifically as shown below. Whenever an object is stored
    '   in a variant, it is late bound and only public events are exposed, not Friend events
    
    ' If the SaveClass property was public, I could reference it simply as...
    ' Call m_col.Item(Index).SaveClass(DataArray)
    
    ' Last but not least, should you be collection various types of objects,
    ' test for the specific object, like so....
    
        If TypeOf collectionObject Is classTestClass Then
            Set tClass = collectionObject
            Call tClass.SaveClass(DataArray)
            ObjectID = "testClass"
        End If
        
    Else    ' restoring
    
        If ObjectID = "testClass" Then
            Set tClass = New classTestClass
            Call tClass.LoadClass(DataArray)
        
            ' Don't add to the collection in this routine
            ' The collection is being restored and
            ' placeholders for the items are already created.
            
            ' Simply return the restored object
            Set collectionObject = tClass
        
        End If
        
    End If
    
End Sub

