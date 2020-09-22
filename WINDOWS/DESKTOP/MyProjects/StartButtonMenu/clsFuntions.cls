VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuntions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Function CreateMenu(ParentKey, lblMenu, colmenu As Collection, imgMenu As Object, imgMenuArrow As Object, imlMain As ImageList, FrmX As Form, stbStart As StatusBar, Optional Top, Optional Left) As Boolean

'This funtion creates the menu and positions it. If submenus are available for that item , it will Recreate this object to use
' another instance of this function.

Dim ii As Integer
Dim ShowCount
Dim CurrentNumber

' Lets see how many items the new menu is going to contain
For ii = 1 To colmenu.Count
    If colmenu(ii).ParentKey = ParentKey Then
        ShowCount = ShowCount + 1
    End If
Next ii

If Not IsMissing(Top) Then
    Top = Top - lblMenu(0).Height
    Top = Top - (ShowCount * lblMenu(0).Height)
End If

For ii = 1 To colmenu.Count
    If colmenu(ii).ParentKey = ParentKey Then ' Only Pick the once that match parentkey.
        CurrentNumber = CurrentNumber + 1
            Load lblMenu(ii)
            colmenu(ii).LabelIndex = ii
            Load imgMenu(ii)
            lblMenu(ii).Caption = colmenu(ii).Description
            Set imgMenu(ii).Picture = imlMain.ListImages(colmenu(ii).Iconkey).Picture
            If Not IsMissing(Left) Then
                imgMenu(ii).Left = Left + imgMenu(ii).Width
            Else
                imgMenu(ii).Left = 100
            End If
            lblMenu(ii).Left = imgMenu(ii).Left + imgMenu(ii).Width

             If Not IsMissing(Top) Then
                lblMenu(ii).Top = Top + (CurrentNumber * lblMenu(ii).Height)
            Else
                lblMenu(ii).Top = ((FrmX.Height - stbStart.Height) - 100) - ((lblMenu(ii).Height * CurrentNumber))
            End If

            imgMenu(ii).Top = lblMenu(ii).Top
    
            'Display arrow if submenu
             If colmenu(ii).MenuType = "S" Then
                Load imgMenuArrow(ii)
                Set imgMenuArrow(ii).Picture = imlMain.ListImages(2).Picture
                imgMenuArrow(ii).Left = lblMenu(ii).Width + lblMenu(ii).Left
                imgMenuArrow(ii).Top = lblMenu(ii).Top
                Dim objFunctions As New clsFuntions ' Create new functions object
                ' Ok, new object to create this items sub menus
                gbCheck = objFunctions.CreateMenu(colmenu(ii).Key, lblMenu, colmenu, imgMenu, imgMenuArrow, imlMain, FrmX, stbStart, lblMenu(ii).Top + lblMenu(ii).Height, imgMenuArrow(ii).Left)
                Set objFunctions = Nothing 'Kill the object
            End If
            
            DoEvents
    End If
Next ii
End Function
