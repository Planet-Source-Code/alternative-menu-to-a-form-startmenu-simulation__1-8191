VERSION 5.00
Begin VB.MDIForm mdiFrmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Start Menu Project"
   ClientHeight    =   7350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8745
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu mnuSubmenu 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuStart 
      Caption         =   "Start"
      Visible         =   0   'False
      WindowList      =   -1  'True
      Begin VB.Menu mnuTest 
         Caption         =   "Test"
         Begin VB.Menu mnuTest4 
            Caption         =   "Testing 4"
         End
         Begin VB.Menu mnuTest5 
            Caption         =   "Testing 5"
         End
      End
      Begin VB.Menu mnuTest2 
         Caption         =   "Test2"
      End
      Begin VB.Menu mnuTest3 
         Caption         =   "Test3"
      End
      Begin VB.Menu mnuUnload 
         Caption         =   "Unload"
      End
   End
End
Attribute VB_Name = "mdiFrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
frmMenu.Show
End Sub
