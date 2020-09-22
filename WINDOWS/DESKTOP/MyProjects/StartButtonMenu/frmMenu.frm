VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7815
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   9315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   9315
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imlMain 
      Left            =   1320
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":03DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0790
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":089C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0C64
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":0DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":11C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1598
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbStart 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   7320
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Picture         =   "frmMenu.frx":1934
            Text            =   "Start"
            TextSave        =   "Start"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   13361
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgMenu 
      Height          =   240
      Index           =   0
      Left            =   3735
      Top             =   4545
      Width           =   330
   End
   Begin VB.Image imgMenuArrow 
      Height          =   240
      Index           =   0
      Left            =   3735
      Top             =   4275
      Width           =   330
   End
   Begin VB.Label lblMenu 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   3735
      MouseIcon       =   "frmMenu.frx":1C50
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4815
      Visible         =   0   'False
      Width           =   2235
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This  is an example of a dynamic menu to simulate a start button effect.
' Author : Renier Barnard
' e-mail : Renier_Barnard@Santam.co.za
' May 2000

Dim colmenu As New Collection
Dim colActive As New Collection
Dim gobjFunctions As New clsFuntions
Dim bStarted As Boolean
Dim LastIndex As Integer
Private Sub Form_Load()

'AddMenu Key, Description, ParentKey, MenuType, Iconkey, level, LabelIndex

'Items on level 1
AddMenu 1, "Unload Program", "", "", 1, 1, 1
AddMenu 2, "TWO-Menu", "", "S", 2, 1, 2
AddMenu 3, "Three", "", "S", 3, 1, 3
AddMenu 4, "Four", "", "S", 4, 1, 4

'Items on level2
AddMenu 5, "Five", 2, "", 4, 2, 5
AddMenu 6, "Six", 2, "", 4, 2, 6
AddMenu 7, "Seven", 2, "S", 4, 2, 7

'Items on level 3
AddMenu 8, "Eight", 7, "", 3, 3, 8
AddMenu 9, "Nine", 7, "", 3, 3, 9


End Sub

Private Function LoadMenu()
Dim ii As Integer
Dim bb As Integer
Dim cc As Integer
    
gbCheck = gobjFunctions.CreateMenu("", lblMenu, colmenu, imgMenu, imgMenuArrow, imlMain, Me, stbStart)

End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

'Dim ii As Integer
'
'For ii = 1 To lblMenu.Count - 1 'Un Highlight all labels
'    lblMenu(ii).ForeColor = RGB(0, 0, 0)
'    DoEvents
'Next ii
HideMenu
End Sub

Private Sub lblMenu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

Dim ii As Integer
On Error Resume Next

If LastIndex = Index Then Exit Sub ' Still on the same menu line. No need to bog down the system

ClearHightLight (LastIndex) ' Clears the previous active menu item from it's HIGHLIGHT

' If the previous menu item had a Sub Menu , we need to go and hide that sub menu
If colmenu(LastIndex).MenuType = "S" And colmenu(LastIndex).Key <> colmenu(Index).ParentKey Then
    HideSubMenu (colmenu(LastIndex).Key)
End If

'For ii = 1 To lblMenu.Count  'Unhighlight all labels
'    lblMenu(ii).ForeColor = RGB(0, 0, 0)
'    DoEvents
'Next ii
HighLight Index ' Highlight the selected item

If colmenu(Index).MenuType = "S" Then
    ShowMenu (colmenu(Index).Key)
End If

LastIndex = Index

End Sub

Private Sub lblMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
' Now, a menu item has been clicked. We need to decide what to do , like call a new form , exit , etc. Nice eh ?

Select Case lblMenu(Index).Caption
    Case "Unload Program"
        Unload Me
        Unload mdiFrmMain
    Case Else
        MsgBox "You Clicked : " & lblMenu(Index).Caption
End Select



End Sub

Public Function ShowMenu(ParentKey)
' Activates a menu , so that it becomes visible.
On Error Resume Next

Dim ii As Integer

For ii = 1 To colmenu.Count
    If colmenu(ii).ParentKey = ParentKey Then
        lblMenu(ii).Visible = True
        imgMenu(ii).Visible = True
        imgMenuArrow(ii).Visible = True
    End If
Next ii

End Function

Private Sub stbStart_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'HideMenu
End Sub

Private Sub stbStart_PanelClick(ByVal Panel As MSComctlLib.Panel)
' The Status Panel has been clicked. If it was the start button on there , then lets fire up our menu
' By activating the ROOT menu first.
    
    If Panel.Text = "Start" Then
        If bStarted = False Then
            LoadMenu ' Create the menu. All items will be invisible , but they will be there AND positioned.
            bStarted = True ' Menu is loaded. Set this flag to prevent a reload.
        End If
        ShowMenu ("") ' Show the ROOT menu
    End If

End Sub
Public Function AddMenu(Key As Variant, Description As String, ParentKey As Variant, MenuType As String, Iconkey As Variant, level As Integer, LabelIndex As Integer)

'Adds the menu objects to the menu collection

    Dim objMenu As New clsMenu
    objMenu.Description = Description
    objMenu.Key = Key
    objMenu.ParentKey = ParentKey
    objMenu.MenuType = MenuType
    objMenu.Iconkey = Iconkey
    objMenu.level = level
    objMenu.LabelIndex = LabelIndex
    colmenu.Add objMenu
    Set objMenu = Nothing

End Function
Public Function HideMenu()
' Hide the entire menu
On Error Resume Next

Dim ii As Integer
For ii = 1 To lblMenu.Count
    lblMenu(ii).Visible = False
    imgMenu(ii).Visible = False
    imgMenuArrow(ii).Visible = False
Next ii

End Function
Public Function HideSubMenu(ParentKey)
' Submenu's parent item has lost focus. We need to hide the inactive sub menu
On Error Resume Next

Dim ii As Integer
For ii = 1 To lblMenu.Count - 1
    If colmenu(ii).ParentKey = ParentKey Then
        lblMenu(ii).Visible = False
        imgMenu(ii).Visible = False
        imgMenuArrow(ii).Visible = False
    End If
    DoEvents
Next ii

End Function

Public Function ClearHightLight(Index)
' Menu item has lost focus. Clear the highlight on it.

Dim ii As Integer

    lblMenu(Index).ForeColor = vbBlack
    lblMenu(Index).FontBold = False
    lblMenu(Index).FontUnderline = False
    lblMenu(Index).BackColor = vbWhite
    lblMenu(Index).BackStyle = 0
    lblMenu(Index).Refresh
    DoEvents
    
End Function
Public Function HighLight(Index)
' Menu item has focus (MOUSEOVER). We need to highlight it to create ACTIVE effect
Dim ii As Integer
    lblMenu(Index).ForeColor = &H8000&
    lblMenu(Index).FontBold = True
    lblMenu(Index).FontUnderline = True
     lblMenu(Index).BackStyle = 1
    lblMenu(Index).BackColor = &H80FFFF
    lblMenu(Index).Refresh
    DoEvents
    
End Function
