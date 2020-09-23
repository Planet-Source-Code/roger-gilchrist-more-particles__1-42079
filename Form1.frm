VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   6450
   ClientLeft      =   165
   ClientTop       =   -135
   ClientWidth     =   13635
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   13635
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()

  'using this empty form to start just allows frmsettings to
  'float above frmMain without further coding
  'Actually This was the orignal from with menus to do settings
  'and as I moved them to the more sophisticated frmsettings
  'this was all that was left. When I tried FrmSettins as startup form
  'it kept disappearing behide frmmain.

    frmMain.Show
    frmSettings.Show , Me

End Sub

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:04:30 PM) 2 + 16 = 18 Lines
