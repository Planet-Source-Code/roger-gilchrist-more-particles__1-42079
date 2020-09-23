VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particles Help"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   3495
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmhelp.frx":0000
      Top             =   120
      Width           =   12135
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   5880
      TabIndex        =   0
      Top             =   3720
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    Me.Hide

End Sub

Private Sub Form_Load()
  'Copyright 2002 Roger Gilchrist
  Dim msg As String

    With Text1
        .Top = 0
        .Left = 0
        .width = Me.width - 120
        .Height = Command1.Top
    End With 'TEXT1
    


    msg = msg & "This is a reworking of 'Geespot_Particles' giving many more options and settings to play with." & vbNewLine
    msg = msg & "See original at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=41231&lngWId=1" & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "Comments are scattered through the code if you want to understand it in detail." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "USAGE:" & vbNewLine
    msg = msg & "Assuming you have a 3 button mouse there are 3 modes of use." & vbNewLine
    msg = msg & "1. Left-Click and Hold button down: Particles follow the mouse until you relasese the button." & vbNewLine
    msg = msg & "2. Right-Click and release: Same as before but you don't need to hold button down." & vbNewLine
    msg = msg & "3. Middle-Click and release: Particles stream from the point at which you clicked." & vbNewLine
    msg = msg & "If you are using mode 2 or 3 then Left-click and release to stop." & vbNewLine
    msg = msg & "Study the code and you will easily see how to adjust the code to operate from something other than the mouse." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "Use method 2 or 3 and you can experiment with the various settings. Unlike the other settings Count stops and restarts to avoid the classes trying to manipulate Particles after they have been unloaded. Also try re-sizing form to smaller sizes to see multiple edge bounces." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "SETTINGS:" & vbNewLine
    msg = msg & "Count:------------ 1 - 500 Number of particles you want. (High numbers are a bit sluggish.)" & vbNewLine
    msg = msg & "Shape:------------ 1-6 the standard VB shapes for Shape control plus the option of randomly choosing between them." & vbNewLine
    msg = msg & "Size:------------- 25 - 200. 25 is minimum visible size. 200 is sluggish." & vbNewLine
    msg = msg & "Life:------------- 10 - 255. Set maximum number of cycles a Particle stays alive (Actually value is randomly set from this value). Particles fade as they age." & vbNewLine
    msg = msg & "Gravity:---------- Where's down and how fast do you go there. Not properly implemented but interesting." & vbNewLine
    msg = msg & "Bounce:----------- 0.0 - 1.0 How much of velocity is kept when Form edge is hit." & vbNewLine
    msg = msg & "Acceleration:----- 0 - 100 interacts with gravity to change velocity." & vbNewLine
    msg = msg & "ColourScheme:----- Fade behaviour of Particles Black2White, White2Black.  Confetti & Christmas fade to white. FireWorks is Christmas fading to black. Random2Black and Random2White" & vbNewLine
    msg = msg & "InitialDirection:- Up, Right, Down, Left and Random. Which way does Particle start moving. " & vbNewLine
    msg = msg & "InitialVelocity:-- Maximum Velocity at start, precise value is random from this value." & vbNewLine
    msg = msg & "InitialSpread:---- Maximum divergence from InitialDirection, precise value is random from this value." & vbNewLine
    msg = msg & "Background: ------ Not part of class, jsut lets you try other background colours than the ones I hard coded into the ColourScheme menu." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "The menus for the numeric values are just sub-sets of possible numbers. ColourScheme can be expanded with programming. Shape cannot be as it depends on VB's settings." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "KNOWN PROBLEMS:" & vbNewLine
    msg = msg & "1. Above 2000 Particles stopping the process leaves artifacts, especially if you start a new process too soon." & vbNewLine
    msg = msg & "2. Acceleration, InitialDirection and InitialSpread behave strangely and interact with each other oddly." & vbNewLine
    msg = msg & "" & vbNewLine
    msg = msg & "Copyright 2002 Roger Gilchrist" & vbNewLine
    msg = msg & "" & vbNewLine

    Text1.Text = msg

End Sub

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:04:13 PM) 1 + 62 = 63 Lines
