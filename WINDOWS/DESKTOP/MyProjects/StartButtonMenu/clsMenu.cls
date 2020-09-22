VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarKey As Variant 'local copy
Private mvarDescription As String 'local copy
Private mvarParentKey As Variant 'local copy
Private mvarIconKey As Variant 'local copy
'local variable(s) to hold property value(s)
Private mvarCallingForm As ErrObject 'local copy
'local variable(s) to hold property value(s)
Private mvarlevel As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarMenuType As String 'local copy
'local variable(s) to hold property value(s)
Private mvarLabelIndex As Integer 'local copy
Public Property Let LabelIndex(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.LabelIndex = 5
    mvarLabelIndex = vData
End Property


Public Property Get LabelIndex() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.LabelIndex
    LabelIndex = mvarLabelIndex
End Property



Public Property Let MenuType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.MenuType = 5
    mvarMenuType = vData
End Property


Public Property Get MenuType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.MenuType
    MenuType = mvarMenuType
End Property



Public Property Let level(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.level = 5
    mvarlevel = vData
End Property


Public Property Get level() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.level
    level = mvarlevel
End Property



Public Property Set CallingForm(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.CallingForm = Form1
    Set mvarCallingForm = vData
End Property


Public Property Get CallingForm() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CallingForm
    Set CallingForm = mvarCallingForm
End Property



Public Property Let Iconkey(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IconKey = 5
    mvarIconKey = vData
End Property


Public Property Set Iconkey(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.IconKey = Form1
    Set mvarIconKey = vData
End Property


Public Property Get Iconkey() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IconKey
    If IsObject(mvarIconKey) Then
        Set Iconkey = mvarIconKey
    Else
        Iconkey = mvarIconKey
    End If
End Property



Public Property Let Parentkey(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.ParentKey = 5
    mvarParentKey = vData
End Property


Public Property Set Parentkey(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.ParentKey = Form1
    Set mvarParentKey = vData
End Property


Public Property Get Parentkey() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.ParentKey
    If IsObject(mvarParentKey) Then
        Set Parentkey = mvarParentKey
    Else
        Parentkey = mvarParentKey
    End If
End Property



Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Key(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Key = 5
    mvarKey = vData
End Property


Public Property Set Key(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Key = Form1
    Set mvarKey = vData
End Property


Public Property Get Key() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Key
    If IsObject(mvarKey) Then
        Set Key = mvarKey
    Else
        Key = mvarKey
    End If
End Property



