VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsParticalFountain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2003 Roger Gilchrist
'Email: rojagilkrist@hotmail.com
'Feel free to use this but please leave copyright notice
'mark any changes as yours if you release a modified version
'and let me know if you used it or modify it
Option Explicit

Private Const PARTICLE_COUNT As Long = 501
Private Particle() As CParticle
Private WithEvents Frm As Form
Attribute Frm.VB_VarHelpID = -1
Private WithEvents tim As Timer
Attribute tim.VB_VarHelpID = -1
Private prt(PARTICLE_COUNT) As Shape
Private MX As Long
Private MY As Long
Private SelectedParticle As Long
Private m_Gravity As Single
Private m_Confetti As Boolean
Private m_Acceleration As Single
Private m_Size As Integer
Private m_Shape As Integer
Private m_InitialDirection As Integer
Private m_MaxInitialSpread As Integer
Private m_MaxInitialVelocity As Integer
Private FreezeStart As Boolean
Private myBackcol As Long
'If you add/delete colour schemes don't forget to
'edit Enum ColourSchemes and Property ColorSchemeName to reflect this
Public Enum ColourSchemes
    Black2White
    White2Black
    Confetti
    Christmas
    FireWorks
    Random2Black
    Random2White
End Enum
Rem Mark off
'Force Enums to maintain Case
#If False Then
Dim Christmas
Dim Black2White
Dim White2Black
Dim Confetti
Dim FireWorks
Dim Random2Black
Dim Random2White
#End If
Rem Mark on
Private m_ColorScheme As ColourSchemes
Private m_BounceFactor As Single
Private m_MaxLife As Integer

Public Property Get Acceleration() As Single

    Acceleration = m_Acceleration

End Property

Public Property Let Acceleration(ByVal Value As Single)

    m_Acceleration = Value

End Property

Public Property Get BackColor() As Long

    BackColor = myBackcol

End Property

Public Property Let BackColor(ByVal Value As Long)

  Dim i As Integer

    myBackcol = Value
    'Reset each particle
    Frm.BackColor = Value
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .BackColor = myBackcol
        End With 'Particle(I)
    Next i

End Property

Private Function BounceDetector(i As Long, TBLR As Integer) As Boolean

  'keep the particle fully on screen
  'Returns True and TBLR (Which edge) if at edge
  'Used to trigger BounceReset

    With Particle(i)
        'Top side
        If .Top <= Frm.ScaleTop Then 'NOT .TOP...
            BounceDetector = True
            .Top = Frm.ScaleTop
            TBLR = 1
        End If
        'Bottom side
        If .Top + .Height >= Frm.ScaleHeight Then
            BounceDetector = True
            .Top = Frm.ScaleHeight - .Height
            TBLR = 2
        End If

        'Left side
        If .Left <= Frm.ScaleLeft Then
            BounceDetector = True
            .Left = Frm.ScaleLeft
            .Velocity = -(.Velocity * m_BounceFactor)
            TBLR = 3
        End If
        'Right side
        If .Left + .width >= Frm.ScaleWidth Then 'NOT .LEFT...
            BounceDetector = True
            .Left = Frm.ScaleWidth - .width
            .Velocity = -(.Velocity * m_BounceFactor)
            TBLR = 4
        End If
    End With 'PARTICLE(I)

End Function

Public Property Get BounceFactor() As Single

    BounceFactor = m_BounceFactor

End Property

Public Property Let BounceFactor(ByVal Value As Single)

  Dim i As Integer

    m_BounceFactor = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .BounceFactor = m_BounceFactor
        End With 'Particle(I)
    Next i

End Property

Private Sub BounceReset(i As Long, TBLR As Integer)

  'Shift direction after edge collisions
  'TBLR top=1 bottom =2 Left=3 Right=4

    With Particle(i)
        If .BounceFactor > 0 Then
            .Velocity = -(.Velocity * .BounceFactor)
            If .Velocity = 0 Then
                .DirectionSpeed = 0 'Freeze in position
              Else 'NOT .VELOCITY...
                .DirectionSpeed = .DirectionSpeed * .BounceFactor
            End If
          Else 'NOT .BOUNCEFACTOR...
            .Velocity = 0
            .DirectionSpeed = 0 'Freeze in position
        End If
        If TBLR > 2 Then 'special for sides
            If .Velocity <> 0 Then
                .DirectionSpeed = -.DirectionSpeed '* m_BounceFactor
            End If
        End If
    End With 'PARTICLE(I)

End Sub

Private Sub Class_Initialize()

  'Default values

    m_BounceFactor = 0.5
    m_Gravity = 1
    m_Acceleration = 5
    m_InitialDirection = 4
    m_MaxInitialVelocity = 150
    m_MaxInitialSpread = 120
    m_MaxLife = 255
    m_Shape = 4
    m_Size = 25

    m_MaxInitialSpread = 90
    m_MaxInitialVelocity = 100
    m_ColorScheme = Random2Black

End Sub

Public Property Get ColorScheme() As ColourSchemes

    ColorScheme = m_ColorScheme

End Property

Public Property Let ColorScheme(ByVal Value As ColourSchemes)

  'If you add/delete colour schemes don't forget to edit
  'Enum ColourSchemes and Property ColorSchemeName to reflect this

  Dim i As Integer

    m_ColorScheme = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .ColorScheme = m_ColorScheme
        End With 'Particle(I)
    Next i

End Property

Public Property Get ColorSchemeName() As String

  'If you add/delete colour schemes don't forget to edit
  'Enum ColourSchemes and Property ColorSchemeName to reflect this

    ColorSchemeName = Choose(m_ColorScheme + 1, "Christmas", "Black2White", "White2Black", "Confetti", "FireWorks", "Random2Black", "Random2White")

End Property

Private Sub Frm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Read which button was pressed and do something

    FreezeStart = Button = 4
    If Button = 1 Then 'Left
        If tim.Enabled Then ' if running stop
            StopDisplay
          Else              ' otherwise start mode 1'TIM.ENABLED = FALSE
            StartDisplay
        End If
    End If
    If Button = 2 Or Button = 4 Then 'Right or Middle
        If tim.Enabled Then          'if running stop
            StopDisplay
          Else                      'get initial pos( only needed for Middle button and start'TIM.ENABLED = FALSE
            MX = X
            MY = Y
            StartDisplay
        End If
    End If

End Sub

Private Sub Frm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Not FreezeStart Then 'if not using Middle button then
        MX = X
        MY = Y
    End If

End Sub

Private Sub Frm_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then 'if using Left button stop
        StopDisplay
    End If

End Sub

Public Property Get Gravity() As Single

    Gravity = m_Gravity

End Property

Public Property Let Gravity(ByVal Value As Single)

  'Unlike other properties this one
  'does not have to reset the individual Particles
  'gravity is a constant for all Particles so only
  'exists at ClsParticleFountain level

    m_Gravity = Value

End Property

Public Property Get InitialDirection() As Integer

    InitialDirection = m_InitialDirection

End Property

Public Property Let InitialDirection(ByVal Value As Integer)

  Dim i As Integer

    m_InitialDirection = Value
    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .InitialDirection = m_InitialDirection
        End With 'Particle(I)
    Next i

End Property

Public Property Get InitialDirectionName() As String

  'If you add directions UR,UL, DL, DR, or some such make sure you edit this

    InitialDirectionName = Choose(m_InitialDirection + 1, "Up", "Right", "Down", "Left", "Random")

End Property

Public Sub Initiate(F As Form, T As Timer, _
                    Optional UseConfetti As ColourSchemes = White2Black, _
                    Optional SetGravity As Integer = 2, Optional SetAcceleration As Integer = 5, _
                    Optional SetBounce As Single = 0.5, Optional SetLife As Integer = 255, _
                    Optional SetSize As Integer = 50)

  'setup controls and set default (or user) values

    Randomize Timer
    Set Frm = F
    Set tim = T
    tim.Interval = 20
    ParticleCreate 0
    ColorScheme = UseConfetti
    Gravity = SetGravity
    Acceleration = SetAcceleration
    BounceFactor = SetBounce
    MaxLife = SetLife
    Size = SetSize
    StopDisplay

End Sub

Public Property Get MaxInitialSpread() As Integer

    MaxInitialSpread = m_MaxInitialSpread

End Property

Public Property Let MaxInitialSpread(ByVal Value As Integer)

  Dim i As Integer

    m_MaxInitialSpread = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .MaxInitialSpread = m_MaxInitialSpread
        End With 'Particle(I)
    Next i

End Property

Public Property Get MaxInitialVelocity() As Integer

    MaxInitialVelocity = m_MaxInitialVelocity

End Property

Public Property Let MaxInitialVelocity(ByVal Value As Integer)

  Dim i As Integer

    m_MaxInitialVelocity = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .MaxInitialVelocity = m_MaxInitialVelocity
        End With 'Particle(I)
    Next i

End Property

Public Property Get MaxLife() As Integer

    MaxLife = m_MaxLife

End Property

Public Property Let MaxLife(ByVal Value As Integer)

  Dim i As Integer

    m_MaxLife = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .MaxLife = m_MaxLife
        End With 'Particle(I)
    Next i

End Property

Public Property Get ParticleCount() As Integer

    ParticleCount = UBound(Particle)

End Property

Public Sub ParticleCreate(ByVal Value As Integer)

  'Create or delete particles
  'Changing this value while running can cause problems
  'so the routine automatically stops and restarts if you change the numbers

  Dim Restartdisplay As Boolean
  Dim i As Integer
  Dim StartAt As Integer

    On Error Resume Next
        If tim.Enabled Then 'Restart test
            StopDisplay
            Restartdisplay = True
        End If
        If Value < UBound(Particle) Then ' lower number of particles
            ReDim Preserve Particle(Value + 1) As CParticle
          Else 'NOT VALUE...
            StartAt = UBound(Particle)
            If StartAt < 0 Then 'Set the initial Particle which exists so must not be loaded
                ReDim Preserve Particle(0) As CParticle
                Set Particle(0) = New CParticle

                Particle(0).init m_ColorScheme, Frm
              Else 'add further Particles'NOT STARTAT...
                ReDim Preserve Particle(Value) As CParticle
                For i = StartAt To Value - 1

                    Set Particle(i) = New CParticle
                    Particle(i).init m_ColorScheme, Frm
                    Particle(i).Top = MX 'place at current mouse pos
                    Particle(i).Top = MY
                Next i

                'Update new particles(so new and old have same values)
                For i = 0 To Value - 1
                    With Particle(i)
                        .BounceFactor = m_BounceFactor
                        .MaxLife = m_MaxLife
                        .Shape = m_Shape
                        .Size = m_Size
                        .InitialDirection = m_InitialDirection
                        .MaxInitialSpread = m_MaxInitialSpread
                        .MaxInitialVelocity = m_MaxInitialVelocity
                        .ColorScheme = m_ColorScheme
                        .BackColor = myBackcol
                    End With 'Particle(I)
                Next i

            End If
        End If
        If Restartdisplay Then
            DoEvents
            StartDisplay
        End If
    On Error GoTo 0

End Sub

Public Property Get Shape() As Integer

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal Value As Integer)

  Dim i As Integer

    m_Shape = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .Shape = m_Shape
        End With 'Particle(I)
    Next i

End Property

Public Property Get ShapeName() As String

  'If you add shapes Edit this
  'NOTE as long as you use Shape controls this cannot be changed
  '

    ShapeName = Choose(m_Shape + 1, "Rectangle", "Square", "Oval", "Circle", "Random")

End Property

Public Property Get Size() As Integer

    Size = m_Size

End Property

Public Property Let Size(ByVal Value As Integer)

  Dim i As Integer

    m_Size = Value

    'Reset each particle
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .Size = m_Size
        End With 'Particle(I)
    Next i

End Property

Private Sub StartDisplay()

  Dim i As Integer

    'setup each Particle
    tim.Enabled = True
    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            .Left = MX
            .Top = MY
            .GenerateParticle
            .ColorScheme = m_ColorScheme
            .Visible = True
        End With 'Particle(I)
        DoEvents
        'for large numbers of Particles this DoEVents allows
        'Particles to start displaying while loading others
    Next i

End Sub

Private Sub StopDisplay()

  Dim i As Integer

    tim.Enabled = False
    If IsArray(Particle) Then
        For i = LBound(Particle) To UBound(Particle) - 1
            Particle(i).Visible = False
            ' DoEvents
            'This DoEvents If enabled causes strange things to
            'happen if you use large numbers
        Next i
    End If

End Sub

Private Sub Tim_Timer()

  'Core of class
  'This does the actual moving of the Particles
  'and except for the Bounce stuff is from the inspiration code

  Dim i As Long, DeadCount As Integer, TBLR As Integer

    For i = LBound(Particle) To UBound(Particle) - 1
        With Particle(i)
            '-- Increase Particles Velocity
            .Velocity = .Velocity + m_Acceleration * m_Gravity
            '-- Increase Age of Particle
            .Age = .Age + 1
            'Deal with edges
            If BounceDetector(i, TBLR) Then
                BounceReset i, TBLR
            End If

            '-- If particle is dead Then reset to insertion point
            If .Age >= .Life Then
                .GenerateParticle
                .Left = MX
                .Top = MY
            End If
            '-- Position Particle
            .Top = .Top + .Velocity
            .Left = .Left + .DirectionSpeed
            .ColourAging
            .MyShape

        End With ''NOT .BOUNCEFACTOR...'Particle(I)
    Next i

End Sub

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:04:22 PM) 47 + 534 = 581 Lines
