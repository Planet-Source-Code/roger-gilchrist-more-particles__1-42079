VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Particles"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5655
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2595
      ScaleWidth      =   5595
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.Timer Timer1 
         Interval        =   1
         Left            =   4920
         Top             =   960
      End
      Begin VB.Frame Frame12 
         Caption         =   "Init Spread"
         Height          =   615
         Left            =   1590
         TabIndex        =   32
         Top             =   1230
         Width           =   1590
         Begin ComCtl2.UpDown UDInitSpread 
            Height          =   255
            Left            =   840
            TabIndex        =   34
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            OrigLeft        =   840
            OrigTop         =   240
            OrigRight       =   1095
            OrigBottom      =   495
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtInitSpread 
            Height          =   285
            Left            =   120
            TabIndex        =   33
            Text            =   "90"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Init Velocity"
         Height          =   615
         Left            =   0
         TabIndex        =   29
         Top             =   1230
         Width           =   1590
         Begin ComCtl2.UpDown UDInitVel 
            Height          =   255
            Left            =   960
            TabIndex        =   31
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   450
            _Version        =   327681
            OrigLeft        =   960
            OrigTop         =   240
            OrigRight       =   1215
            OrigBottom      =   495
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtinitvel 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Text            =   "100"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Acceleration"
         Height          =   615
         Left            =   3180
         TabIndex        =   28
         Top             =   615
         Width           =   1590
         Begin ComCtl2.UpDown UDAccel 
            Height          =   375
            Left            =   840
            TabIndex        =   37
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtAccel 
            Height          =   285
            Left            =   120
            TabIndex        =   36
            Text            =   "Text1"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Max"
         Height          =   255
         Left            =   5040
         TabIndex        =   27
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Help"
         Height          =   255
         Left            =   5040
         TabIndex        =   26
         Top             =   480
         Width           =   495
      End
      Begin VB.Frame Frame9 
         Caption         =   "Init Direction"
         Height          =   615
         Left            =   0
         TabIndex        =   24
         Top             =   1845
         Width           =   1590
         Begin VB.ComboBox cmboDirection 
            Height          =   315
            Left            =   120
            TabIndex        =   25
            Text            =   "Combo1"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "BackGround"
         Height          =   615
         Left            =   3180
         TabIndex        =   22
         Top             =   1230
         Width           =   1590
         Begin VB.ComboBox cmboBackCol 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "ColourScheme"
         Height          =   615
         Left            =   3180
         TabIndex        =   20
         Top             =   1845
         Width           =   1590
         Begin VB.ComboBox cmboColSchm 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Shape"
         Height          =   615
         Left            =   1590
         TabIndex        =   18
         Top             =   1845
         Width           =   1590
         Begin VB.ComboBox cmboShape 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Pop"
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1590
         Begin ComCtl2.UpDown UDPop 
            Height          =   615
            Left            =   1080
            TabIndex        =   3
            Top             =   120
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   1085
            _Version        =   327681
            OrigLeft        =   1680
            OrigTop         =   240
            OrigRight       =   1935
            OrigBottom      =   855
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtPop 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   2
            Text            =   "50"
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Size"
         Height          =   615
         Left            =   1590
         TabIndex        =   4
         Top             =   0
         Width           =   1590
         Begin ComCtl2.UpDown UDSize 
            Height          =   495
            Left            =   1080
            TabIndex        =   6
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   873
            _Version        =   327681
            OrigLeft        =   1080
            OrigTop         =   240
            OrigRight       =   1335
            OrigBottom      =   735
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtSize 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   5
            Text            =   "10"
            Top             =   240
            Width           =   1320
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Life"
         Height          =   615
         Left            =   3180
         TabIndex        =   7
         Top             =   0
         Width           =   1590
         Begin ComCtl2.UpDown UDLife 
            Height          =   495
            Left            =   1080
            TabIndex        =   9
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   873
            _Version        =   327681
            OrigLeft        =   1080
            OrigTop         =   240
            OrigRight       =   1335
            OrigBottom      =   735
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtlife 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   8
            Text            =   "255"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Gravity(±10)"
         Height          =   615
         Left            =   0
         TabIndex        =   10
         Top             =   615
         Width           =   1590
         Begin ComCtl2.UpDown UDGravity 
            Height          =   375
            Left            =   960
            TabIndex        =   13
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            _Version        =   327681
            OrigLeft        =   960
            OrigTop         =   240
            OrigRight       =   1215
            OrigBottom      =   615
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox TxtGravLink 
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Text            =   "10"
            Top             =   240
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox txtgravVis 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   11
            Text            =   "1"
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Bounce(0-1)"
         Height          =   615
         Left            =   1590
         TabIndex        =   14
         Top             =   615
         Width           =   1590
         Begin ComCtl2.UpDown UDbounce 
            Height          =   495
            Left            =   1080
            TabIndex        =   17
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   873
            _Version        =   327681
            OrigLeft        =   1080
            OrigTop         =   240
            OrigRight       =   1335
            OrigBottom      =   735
            Increment       =   0
            Max             =   0
            Enabled         =   -1  'True
         End
         Begin VB.TextBox txtBounceVis 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Text            =   "0.5"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox TxtBounceLink 
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Text            =   "5"
            Top             =   240
            Visible         =   0   'False
            Width           =   1350
         End
      End
      Begin VB.Label lblReciever 
         Caption         =   "Reciever Label"
         Height          =   735
         Left            =   5040
         TabIndex        =   35
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2003 Roger Gilchrist
'Email: rojagilkrist@hotmail.com
'Feel free to use this but please leave copyright notice
'mark any changes as yours if you release a modified version
'and let me know if you used it or modify it

Option Explicit

Private SettingUp As Boolean
Private cUDpop As New ClsUpDown
Private cUDsize As New ClsUpDown
Private cUDlife As New ClsUpDown
Private cUDgrav As New ClsUpDown
Private cUDbonc As New ClsUpDown
Private cUDivel As New ClsUpDown
Private cUDisprd As New ClsUpDown
Private cUDAccel As New ClsUpDown

Private Sub cmboBackCol_Click()

    If SettingUp Then
        Exit Sub '>---> Bottom
    End If

    Select Case cmboBackCol.ListIndex
      Case 0
        PF.BackColor = vbBlack
      Case 1
        PF.BackColor = vbWhite
      Case 2
        PF.BackColor = RGB(78, 78, 78) ' RGB(128, 128, 128)
      Case 3
        PF.BackColor = vbBlue
      Case 4
        PF.BackColor = vbRed
      Case 5
        PF.BackColor = vbGreen
    End Select

End Sub

Private Sub cmboColSchm_click()

    If SettingUp Then
        Exit Sub '>---> Bottom
    End If

    Select Case cmboColSchm.ListIndex
      Case 0
        PF.ColorScheme = Black2White
      Case 1
        PF.ColorScheme = White2Black
      Case 2
        PF.ColorScheme = Confetti
      Case 3
        PF.ColorScheme = Christmas
      Case 4
        PF.ColorScheme = FireWorks
      Case 5
        PF.ColorScheme = Random2Black
      Case 6
        PF.ColorScheme = Random2White
    End Select

End Sub

Private Sub cmboDirection_Click()

    PF.InitialDirection = cmboDirection.ListIndex

End Sub

Private Sub cmboShape_Click()

    If SettingUp Then
        Exit Sub '>---> Bottom
    End If
    PF.Shape = cmboShape.ListIndex

End Sub

Private Sub Command1_Click()

    If frmMain.WindowState = vbNormal Then
        Command1.Caption = "Min"
        frmMain.WindowState = vbMaximized
        frmSettings.Top = frmMain.Top + 100
      Else 'NOT FRMMAIN.WINDOWSTATE...
        Command1.Caption = "Max"
        frmMain.WindowState = vbNormal
        frmSettings.Top = frmMain.Top + frmMain.Height
    End If
    frmSettings.Left = (frmSettings.width - frmSettings.width) / 2

End Sub

Private Sub Command2_Click()

    Form1.Show , frmMain

End Sub

Private Sub Form_Load()

  Dim i As Integer

    SettingUp = True
    PF.Initiate frmMain, Timer1
    With cmboDirection
        .Clear
        .AddItem "Up"
        .AddItem "Right"
        .AddItem "Down"
        .AddItem "Left"
        .AddItem "Rnd"
        .ListIndex = 4
    End With 'CMBODIRECTION

    With cmboBackCol
        .Clear
        .AddItem "Black"
        .AddItem "White"
        .AddItem "Grey"
        .AddItem "Blue"
        .AddItem "Red"
        .AddItem "Green"
        .ListIndex = 2
    End With 'CMBOBACKCOL
    With cmboColSchm
        .Clear
        .AddItem "Black2White"
        .AddItem "White2Black"
        .AddItem "Confetti"
        .AddItem "Christmas"
        .AddItem "FireWorks"
        .AddItem "Rnd2Black"
        .AddItem "Rnd2White"
        .ListIndex = 5
    End With 'CMBOCOLSCHM

    With cmboShape
        .Clear
        .AddItem "Rectangle"
        .AddItem "Square"
        .AddItem "Oval"
        .AddItem "Circle"
        .AddItem "Rnd"
        .ListIndex = 3
    End With 'CMBOSHAPE

    cUDpop.AssignControls UDPop, txtPop, 1, 5000, 150, 10, , , Centre, 2, "Pop", , lblReciever
    Frame2.Caption = cUDpop.Caption
    cUDsize.AssignControls UDSize, txtSize, 1, 300, 30, 5, , , Centre, 2, "Size", , lblReciever
    Frame3.Caption = cUDsize.Caption
    cUDlife.AssignControls UDLife, txtlife, 10, 1000, 250, 50, , , Centre, 2, "Life", , lblReciever
    Frame4.Caption = cUDlife.Caption
    cUDgrav.AssignControls UDGravity, txtgravVis, -10, 10, 1, 0.1, , , Centre, 2, "Gravity", TxtGravLink, lblReciever
    Frame5.Caption = cUDgrav.Caption
    cUDbonc.AssignControls UDbounce, txtBounceVis, 0, 1, 0.7, 0.1, , , Centre, 10, "Bounce", TxtBounceLink, lblReciever
    Frame6.Caption = cUDbonc.Caption
    cUDivel.AssignControls UDInitVel, txtinitvel, 0, 1000, 100, 10, , , Centre, 10, "IntVel", , lblReciever
    Frame11.Caption = cUDivel.Caption
    cUDisprd.AssignControls UDInitSpread, txtInitSpread, 0, 1000, 100, 10, , , Centre, 10, "IntDir", , lblReciever

    Frame12.Caption = cUDisprd.Caption
    cUDAccel.AssignControls UDAccel, txtAccel, 0, 100, 10, 1, , , Centre, 5, "Accel", , lblReciever
    Frame10.Caption = cUDAccel.Caption
SettingUp = False
    With PF
        .ParticleCreate cUDpop.Value
        .Acceleration = cUDAccel.Value
        .Gravity = cUDgrav.Value
        .BounceFactor = cUDbonc.Value
        .MaxLife = cUDlife.Value
        .Size = cUDsize.Value
        .MaxInitialVelocity = cUDivel.Value
        .MaxInitialSpread = cUDisprd.Value
        cmboBackCol_Click
        cmboColSchm_click
        cmboShape_Click
        .InitialDirection = cmboDirection.ListIndex
        frmMain.BackColor = .BackColor
    End With 'PF
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub

Private Sub lblreciever_Change()

    Select Case lblReciever.Tag
      Case cUDpop.ID
        PF.ParticleCreate cUDpop.Value
      Case cUDsize.ID
        PF.Size = cUDsize.Value
      Case cUDlife.ID
        PF.MaxLife = cUDlife.Value
      Case cUDgrav.ID
        PF.Gravity = cUDgrav.Value
      Case cUDbonc.ID
        PF.BounceFactor = cUDbonc.Value
      Case cUDivel.ID
        PF.MaxInitialVelocity = cUDivel.Value
      Case cUDisprd.ID
        PF.MaxInitialSpread = cUDisprd.Value
      Case cUDAccel.ID
        PF.Acceleration = cUDAccel.Value
    End Select

End Sub

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:04:10 PM) 12 + 205 = 217 Lines
Private Sub Picture1_Click()

End Sub
