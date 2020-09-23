VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "classTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' A simple class used for testing the collection class's SaveCollection/LoadCollection methods


Private c_Name As String
Private c_Index As Long


Public Property Get Name() As String
    Name = c_Name
End Property
Public Property Let Name(Value As String)
    c_Name = Value
End Property

Public Property Get Index() As Long
    Index = c_Index
End Property
Public Property Let Index(Value As Long)
    c_Index = Value
End Property


Friend Function SaveClass(classBytes() As Byte) As Boolean
    ' simple method of serializing a class into a byte array
    ' VB5 users cannot use this method; however, one could simply use an INI-type structure instead
    Dim pp As New PropertyBag
    
    pp.WriteProperty "Name", c_Name
    pp.WriteProperty "Index", c_Index
    
    classBytes = pp.Contents

End Function

Friend Function LoadClass(classBytes() As Byte) As Boolean
    ' simple method of un-serializing a class from a byte array
    
    Dim pp As New PropertyBag
    On Error Resume Next
    
    pp.Contents = classBytes
    If Err Then
        Err.Clear
        Exit Function
    Else
        c_Name = pp.ReadProperty("Name", vbNullString)
        c_Index = pp.ReadProperty("Index", 0)
    End If
    
End Function
