VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Variant Array Test -- Supports Nested Arrays"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add stdFont"
      Height          =   345
      Index           =   12
      Left            =   1470
      TabIndex        =   12
      Top             =   4605
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add stdPicture"
      Height          =   345
      Index           =   11
      Left            =   15
      TabIndex        =   11
      Top             =   4605
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Null Object"
      Height          =   345
      Index           =   10
      Left            =   15
      TabIndex        =   10
      Top             =   4215
      Width           =   1425
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Play Again"
      Enabled         =   0   'False
      Height          =   630
      Left            =   3870
      TabIndex        =   14
      Top             =   2475
      Width           =   1215
   End
   Begin VB.ListBox List2 
      Height          =   4155
      Left            =   5115
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   465
      Width           =   2385
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Array"
      Height          =   345
      Index           =   9
      Left            =   15
      TabIndex        =   9
      Top             =   3825
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add String"
      Height          =   345
      Index           =   8
      Left            =   15
      TabIndex        =   8
      Top             =   3435
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add TestClass"
      Height          =   345
      Index           =   7
      Left            =   15
      TabIndex        =   7
      Top             =   3060
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Date"
      Height          =   345
      Index           =   6
      Left            =   15
      TabIndex        =   6
      Top             =   2685
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Byte"
      Height          =   345
      Index           =   5
      Left            =   15
      TabIndex        =   5
      Top             =   2310
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Currency"
      Height          =   345
      Index           =   4
      Left            =   15
      TabIndex        =   4
      Top             =   1935
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Single"
      Height          =   345
      Index           =   3
      Left            =   15
      TabIndex        =   3
      Top             =   1545
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Double"
      Height          =   345
      Index           =   2
      Left            =   15
      TabIndex        =   2
      Top             =   1170
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Long"
      Height          =   345
      Index           =   1
      Left            =   15
      TabIndex        =   1
      Top             =   795
      Width           =   1425
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add Integer"
      Height          =   345
      Index           =   0
      Left            =   15
      TabIndex        =   0
      Top             =   420
      Width           =   1425
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   435
      Width           =   2385
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test Variant Arrays"
      Enabled         =   0   'False
      Height          =   630
      Left            =   3885
      TabIndex        =   13
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Restored Variant Array"
      Height          =   300
      Index           =   1
      Left            =   5160
      TabIndex        =   18
      Top             =   120
      Width           =   2310
   End
   Begin VB.Label Label1 
      Caption         =   "Soruce Variant Array"
      Height          =   300
      Index           =   0
      Left            =   1530
      TabIndex        =   17
      Top             =   120
      Width           =   2310
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private WithEvents myCol As cNodeCollection
Attribute myCol.VB_VarHelpID = -1

Private Sub cmdAdd_Click(Index As Integer)
    
    Dim vArray() As Variant, tClass As classTestClass
    Dim tLong(1 To 5) As Long
    
    If myCol Is Nothing Then
        Set myCol = New cNodeCollection
        myCol.AddItem vArray, "Test"
        ReDim vArray(0)
        cmdTest.Enabled = True
    Else
        vArray = myCol.Item(1)
        ReDim Preserve vArray(0 To UBound(vArray) + 1)
    End If
    
    Select Case Index
        Case 0: vArray(UBound(vArray)) = CInt(Rnd * 3000)
            List1.AddItem "Integer: " & vArray(UBound(vArray))
        Case 1: vArray(UBound(vArray)) = CLng(Rnd * vbWhite)
            List1.AddItem "Long : " & vArray(UBound(vArray))
        Case 2: vArray(UBound(vArray)) = CDbl(Timer)
            List1.AddItem "Double : " & vArray(UBound(vArray))
        Case 3: vArray(UBound(vArray)) = CSng(Rnd * vbWhite)
            List1.AddItem "Single: " & vArray(UBound(vArray))
        Case 4: vArray(UBound(vArray)) = CCur(Timer)
            List1.AddItem "Currency: " & vArray(UBound(vArray))
        Case 5: vArray(UBound(vArray)) = CByte(Rnd * 255)
            List1.AddItem "Byte: " & vArray(UBound(vArray))
        Case 6: vArray(UBound(vArray)) = Now()
            List1.AddItem "Date: " & vArray(UBound(vArray))
        Case 7
            Set tClass = New classTestClass
            tClass.Name = "LaVolpe"
            Set vArray(UBound(vArray)) = tClass
            List1.AddItem "Class (Name Property): " & vArray(UBound(vArray)).Name
        Case 8: vArray(UBound(vArray)) = "Variant Array Test"
            List1.AddItem "String: " & vArray(UBound(vArray))
        Case 9:
            List1.AddItem "Long Array:(Values) 1,2,3,4,5"
            tLong(1) = 1: tLong(2) = 2: tLong(3) = 3: tLong(4) = 4: tLong(5) = 5
            vArray(UBound(vArray)) = tLong
        Case 10
            List1.AddItem "Object Is Nothing"
            Set vArray(UBound(vArray)) = tClass
        Case 11
            List1.AddItem "Icon: " & Int(ScaleX(Me.Icon.Width, vbHimetric, vbPixels)) & "x" & Int(ScaleY(Me.Icon.Height, vbHimetric, vbPixels))
            Set vArray(UBound(vArray)) = Me.Icon
        Case 12
            List1.AddItem "Font: " & Me.Font.Name
            Set vArray(UBound(vArray)) = Me.Font
    End Select
    
    myCol.Item("Test") = vArray
    
    cmdAdd(Index).Enabled = False
    
End Sub

Private Sub cmdTest_Click()

    Dim p() As Byte, x As Long, I As Integer, sListItem As String
    
    List2.Clear
    cmdReset.Enabled = True
    cmdTest.Enabled = False
    For I = cmdAdd.LBound To cmdAdd.UBound
        cmdAdd(I).Enabled = False
    Next
    
    myCol.SaveCollection p()
    
    Set myCol = Nothing
    Set myCol = New cNodeCollection
    
    myCol.LoadCollection p()
    
    For x = LBound(myCol.Item(1)) To UBound(myCol.Item(1))
    
        sListItem = vbNullString
    
        If (VarType(myCol.Item(1)(x)) And vbArray) = vbArray Then
            ' array element is arrayed
            Select Case VarType(myCol.Item(1)(x)) And Not vbArray
            Case vbString: ' string array
            Case vbVariant: ' variant array & you get the idea....
            Case vbInteger, vbBoolean, vbByte, vbSingle, vbDouble, vbDate, vbCurrency
            Case vbLong
                sListItem = "Long Array:(Values) "
                For I = LBound(myCol.Item(1)(x)) To UBound(myCol.Item(1)(x))
                    sListItem = sListItem & CStr(myCol.Item(1)(x)(I)) & ","
                Next
                sListItem = Left$(sListItem, Len(sListItem) - 1)
            Case vbObject
            Case Else
                ' could be vbUserDefinedType, vbDataObject, vbNull, vbEmpty, vbError
            End Select
            
        Else
            If IsObject(myCol.Item(1)(x)) = True Then
                If myCol.Item(1)(x) Is Nothing Then
                    sListItem = "Object Is Nothing"
                Else
                    If TypeOf myCol.Item(1)(x) Is classTestClass Then
                        sListItem = "Class (Name Property): " & myCol.Item(1)(x).Name
                    ElseIf TypeOf myCol.Item(1)(x) Is StdPicture Then
                        sListItem = "Icon: " & Int(ScaleX(myCol.Item(1)(x).Width, vbHimetric, vbPixels)) & "x" & Int(ScaleY(myCol.Item(1)(x).Height, vbHimetric, vbPixels))
                    ElseIf TypeOf myCol.Item(1)(x) Is StdFont Then
                        sListItem = "Font: " & myCol.Item(1)(x).Name
                    Else
                        sListItem = "Unknown Object"
                    End If
                End If
            Else
                Select Case VarType(myCol.Item(1)(x))
                Case vbString: sListItem = "String: " & myCol.Item(1)(x)
                Case vbDouble: sListItem = "Double: " & myCol.Item(1)(x)
                Case vbSingle: sListItem = "Single: " & myCol.Item(1)(x)
                Case vbLong: sListItem = "Long: " & myCol.Item(1)(x)
                Case vbCurrency: sListItem = "Currency: " & myCol.Item(1)(x)
                Case vbInteger: sListItem = "Integer: " & myCol.Item(1)(x)
                Case vbByte: sListItem = "Byte: " & myCol.Item(1)(x)
                Case vbDate: sListItem = "Date: " & myCol.Item(1)(x)
                Case Else
                    sListItem = "Unknown data type"
                End Select
            End If
        End If
            
        List2.AddItem sListItem
        
    Next
    
End Sub

Private Sub cmdReset_Click()
    ' Play Again
    List2.Clear
    List1.Clear
    Dim x As Long
    For x = cmdAdd.LBound To cmdAdd.UBound
        cmdAdd(x).Enabled = True
    Next
    Set myCol = Nothing
    cmdReset.Enabled = False

End Sub

Private Sub Form_Load()
    Randomize Timer
End Sub

Private Sub List1_Click()
    If List1.ListCount > 0 Then List2.ListIndex = List1.ListIndex
End Sub

Private Sub List2_Click()
    If List2.ListCount > 0 Then List1.ListIndex = List2.ListIndex
End Sub

Private Sub myCol_SerializeObject(ByVal Serialize As Boolean, collectionObject As Object, DataArray() As Byte, ObjectID As String)
    
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
    
    Dim tClass As classTestClass
    
    If Serialize = True Then
    
        If TypeOf collectionObject Is classTestClass Then
            Set tClass = collectionObject
            tClass.SaveClass DataArray()
            ObjectID = "testclass"
        Else
        
        End If
        
    Else
    
        If ObjectID = "testclass" Then
            Set tClass = New classTestClass
            tClass.LoadClass DataArray()
            Set collectionObject = tClass
        End If
        
    End If
    
    
End Sub


Private Function IsArrayEmpty(ByVal arrayPtr As Long) As Boolean
    ' PURPOSE: Helper Function. Tests to see if an array has been initialized
    ' Called by majority of routines in this class
    IsArrayEmpty = (arrayPtr = -1)
End Function


