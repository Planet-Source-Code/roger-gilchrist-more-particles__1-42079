VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CParticle"
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
Private MyCanvas As Form
Public Velocity As Double
Public m_BounceFactor As Double
Public Direction As Long
Public DirectionSpeed As Long
Public Life As Long
Public Age As Long
Private initAge As Long
Private m_Red As Integer
Private m_Blue As Integer
Private m_Green As Integer

Private m_Confetti As Boolean
Private L_ColourScheme As ColourSchemes
Private m_MaxLife As Integer
Private m_Size As Integer
Private m_Shape As ShapeSet
Private m_Exists As Boolean
Private m_InitialDirection As Integer
Private m_MaxInitialVelocity As Integer
Private m_MaxInitialSpread As Integer
Private M_BorderColor As Long
Private M_FillColor As Long
Private m_Height As Long
Private m_Width As Long
Private m_Left As Long
Private m_Top As Long
Private m_PrevLeft As Long
Private m_PrevTop As Long
Private myBackColor As Long

Private m_Visible As Boolean

Public Enum ShapeSet
    sRectangle
    sSquare
    sOval
    sCircle
    sRandom
End Enum

Public Property Get BackColor() As Long

    BackColor = myBackColor

End Property

Public Property Let BackColor(ByVal Value As Long)

    myBackColor = Value

End Property

Public Property Get Blue() As Integer

    Blue = m_Blue

End Property

Public Property Let Blue(ByVal Value As Integer)

    m_Blue = Value

End Property

Public Property Get BorderColor() As Long

    BorderColor = M_BorderColor

End Property

Public Property Let BorderColor(ByVal Value As Long)

    M_BorderColor = Value

End Property

Public Property Get BounceFactor() As Single

    BounceFactor = m_BounceFactor

End Property

Public Property Let BounceFactor(ByVal Value As Single)

    m_BounceFactor = Value

End Property

Private Sub Class_Initialize()

    m_Exists = True
    Randomize Timer
    m_Size = 200 '25
    m_Shape = vbShapeCircle
    GenerateParticle
    m_InitialDirection = 2
    m_MaxInitialVelocity = 150
    m_MaxInitialSpread = 120

End Sub

Private Function ColorAge(Clr As Integer, Optional Up As Boolean = False) As Integer

  'fade a member of RGB colour set

  Dim AgeFactor As Long

    AgeFactor = Percent(Age, Life) '* 2

    If Up Then 'fade to white
        If Clr + AgeFactor < 255 Then
            ColorAge = Clr + AgeFactor
          Else 'NOT CLR...
            ColorAge = 255
        End If

      Else 'fade to black 'UP = FALSE

        If Clr - AgeFactor > 0 Then
            ColorAge = Clr - AgeFactor
          Else 'NOT CLR...
            ColorAge = 0
        End If
    End If

End Function

Private Function ColorMy() As Long

  Dim Fclr As Integer

    Select Case L_ColourScheme
      Case FireWorks, White2Black, Random2Black
        ColorMy = RGB(ColorAge(m_Red), ColorAge(m_Green), ColorAge(m_Blue))
      Case Christmas, Black2White, Random2White, Confetti
        ColorMy = RGB(ColorAge(m_Red, True), ColorAge(m_Green, True), ColorAge(m_Blue, True))
    End Select

End Function

Public Property Get ColorScheme() As ColourSchemes

    ColorScheme = L_ColourScheme

End Property

Public Property Let ColorScheme(Value As ColourSchemes)

    L_ColourScheme = Value
    ColourSet

End Property

Public Property Get Colour() As Long

    Colour = M_BorderColor

End Property

Public Property Let Colour(ByVal Value As Long)

    MyCanvas.FillStyle = vbSolid
    MyCanvas.FillColor = Value
    M_FillColor = Value

    M_BorderColor = M_FillColor

End Property

Public Sub ColourAging()

  'Fade colours to black or white

    Select Case L_ColourScheme
      Case FireWorks, White2Black, Random2Black
        'Fade2Black
        Colour = RGB(ColorAge(m_Red), ColorAge(m_Green), ColorAge(m_Blue))
      Case Christmas, Black2White, Random2White, Confetti
        'Fade2White
        Colour = RGB(ColorAge(m_Red, True), ColorAge(m_Green, True), ColorAge(m_Blue, True))
    End Select

End Sub

Private Sub ColourSet()

  'Set initial colour values depending on scheme

    Select Case L_ColourScheme
      Case Random2Black, Random2White
        Red = Int(Rnd * 256)
        Green = Int(Rnd * 256)
        Blue = Int(Rnd * 256)

      Case Christmas, FireWorks
         If Rnd > 0.5 Then
      Red = 221 + Int(Rnd * 30)
      End If
      If Red = 0 Then
         If Rnd > 0.5 Then
      Green = 221 + Int(Rnd * 30)
      End If
      End If
        m_Blue = 0
      Case Black2White
        Red = 0
        Green = Red
        Blue = Red
      Case White2Black
        Red = 255
        Green = Red
        Blue = Red
      Case Confetti
        If Rnd > 0.5 Then
            Red = 255
        End If
        If Rnd > 0.5 Then
            Green = 255
        End If
        If Rnd > 0.5 Then
            Blue = 255
        End If

    End Select

End Sub

Public Property Get FillColor() As Long

    FillColor = M_FillColor

End Property

Public Property Let FillColor(ByVal Value As Long)

    M_FillColor = Value

End Property

Public Sub GenerateParticle()

  'creates initial movemet, lifespan and direction for a Particle

    Select Case m_InitialDirection
      Case 0 To 3
        SetInitialDirection m_InitialDirection
      Case Else
        SetInitialDirection Int(Rnd * 4)
    End Select
    Life = Int(Rnd * m_MaxLife) + 1               ' 20 for a Fire look
    Age = 0
    '-- Determine direction and speed
    If (Abs(Direction) Mod 2 = 0) Then
        DirectionSpeed = -DirectionSpeed
    End If
    RefreshParticle

End Sub

Public Property Get Green() As Integer

    Green = m_Green

End Property

Public Property Let Green(ByVal Value As Integer)

    m_Green = Value

End Property

Public Property Get Height() As Long

    Height = m_Height

End Property

Public Property Let Height(ByVal Value As Long)

    m_Height = Value

End Property

Public Sub init(ClrSchm As ColourSchemes, Frm As Form)

  'Initialize Particle

    Set MyCanvas = Frm
    BackColor = MyCanvas.BackColor
    MyCanvas.AutoRedraw = True
    RefreshParticle
    L_ColourScheme = ClrSchm
    Colour = RGB(m_Red, m_Green, m_Blue)

End Sub

Public Property Get InitialDirection() As Integer

    InitialDirection = m_InitialDirection

End Property

Public Property Let InitialDirection(ByVal Value As Integer)

    m_InitialDirection = Value
    SetInitialDirection m_InitialDirection

End Property

Public Property Get Left() As Long

    Left = m_Left

End Property

Public Property Let Left(ByVal Value As Long)

    m_Left = Value

End Property

Public Property Get MaxInitialSpread() As Integer

    MaxInitialSpread = m_MaxInitialSpread

End Property

Public Property Let MaxInitialSpread(ByVal Value As Integer)

    m_MaxInitialSpread = Value

End Property

Public Property Get MaxInitialVelocity() As Integer

    MaxInitialVelocity = m_MaxInitialVelocity

End Property

Public Property Let MaxInitialVelocity(ByVal Value As Integer)

    m_MaxInitialVelocity = Value

End Property

Public Property Get MaxLife() As Integer

    MaxLife = m_MaxLife

End Property

Public Property Let MaxLife(ByVal Value As Integer)

    m_MaxLife = Value
    RefreshParticle

End Property

Public Sub MyShape()

    Select Case m_Shape
      Case sRectangle
        'object.Line [Step] (x1, y1) [Step] - (x2, y2), [color], [B][F]
        MyCanvas.FillColor = myBackColor
        MyCanvas.Line (m_PrevLeft, m_PrevTop)-(m_PrevLeft + m_Width, m_PrevTop + m_Height), myBackColor, BF
        If m_Visible Then
            If Age < Life Then
                MyCanvas.FillColor = ColorMy
                MyCanvas.Line (m_Left, m_Top)-(m_Left + m_Width, m_Top + m_Height), ColorMy, BF
            End If
          Else 'M_VISIBLE = FALSE
            MyCanvas.FillColor = myBackColor
            MyCanvas.Line (m_Left, m_Top)-(m_Left + m_Width, m_Top + m_Height), myBackColor, BF
        End If

      Case sSquare

        MyCanvas.FillColor = myBackColor
        MyCanvas.Line (m_PrevLeft, m_PrevTop)-(m_PrevLeft + m_Size, m_PrevTop + m_Size), myBackColor, BF
        If m_Visible Then
            If Age < Life Then
                MyCanvas.FillColor = ColorMy
                MyCanvas.Line (m_Left, m_Top)-(m_Left + m_Size, m_Top + m_Size), ColorMy, BF
            End If
          Else 'M_VISIBLE = FALSE
            MyCanvas.FillColor = myBackColor
            MyCanvas.Line (m_Left, m_Top)-(m_Left + m_Size, m_Top + m_Size), myBackColor, BF
        End If

      Case sOval
        'object.Circle [Step] (x, y), radius, [color, start, end, aspect]
        MyCanvas.FillColor = myBackColor
        MyCanvas.Circle (m_PrevLeft, m_PrevTop), m_Size, myBackColor, , , m_Width / m_Height  ' RGB(128, 128, 128) '
        If m_Visible Then
            If Age < Life Then
                MyCanvas.FillColor = ColorMy
                MyCanvas.Circle (m_Left, m_Top), m_Size, ColorMy, , , m_Width / m_Height
            End If
          Else 'M_VISIBLE = FALSE
            MyCanvas.FillColor = myBackColor
            MyCanvas.Circle (m_Left, m_Top), m_Size, myBackColor, , , m_Width / m_Height
        End If
      Case sCircle

        MyCanvas.FillColor = myBackColor ' RGB(128, 128, 128) '
        MyCanvas.Circle (m_PrevLeft, m_PrevTop), m_Size, myBackColor  ' RGB(128, 128, 128) '

        If m_Visible Then
            If Age < Life Then
                MyCanvas.FillColor = ColorMy
                MyCanvas.Circle (m_Left, m_Top), m_Size, ColorMy
            End If
          Else 'M_VISIBLE = FALSE
            MyCanvas.FillColor = myBackColor
            MyCanvas.Circle (m_Left, m_Top), m_Size, myBackColor
        End If
    End Select
    m_PrevLeft = m_Left
    m_PrevTop = m_Top

End Sub

Private Function Percent(a, b) As Integer

  'calculate percentage given two numbers

  Dim c As Integer

    If a > b Then
        a = b
    End If
    If b = 0 Then
        Percent = 0
        Exit Function '>---> Bottom
    End If
    c = Int(a / b * 100)
    If c > 100 Then
        c = 100
    End If
    Percent = c

End Function

Public Property Get Red() As Integer

    Red = m_Red

End Property

Public Property Let Red(ByVal Value As Integer)

    m_Red = Value

End Property

Private Sub RefreshParticle()

  'Set size, shape and edges for Particle

    m_Size = IIf(m_Size = -999, 25 + Int(Rnd * 175), m_Size)
    m_Height = m_Size
    If m_Shape = sRectangle Or m_Shape = sOval Then
        If Rnd > 0.5 Then
            m_Height = m_Size * 2
            m_Width = m_Size
          Else 'NOT RND...
            m_Width = m_Size * 2
            m_Height = m_Size
        End If

      Else 'NOT M_SHAPE...
        m_Width = m_Size
        m_Height = m_Size
    End If

End Sub

Private Sub SetInitialDirection(Value As Integer)

    Select Case Value
      Case 0 'Up
        Velocity = -m_MaxInitialVelocity + Int(Rnd * m_MaxInitialVelocity) + 1
        DirectionSpeed = Int(Rnd * m_MaxInitialSpread) - 90
        Direction = m_MaxInitialSpread - Int(Rnd * m_MaxInitialSpread) + 1
        Direction = Int(Rnd * m_MaxInitialSpread) - 90
      Case 1 'Right
        Velocity = -m_MaxInitialVelocity / 2 + Int(Rnd * m_MaxInitialVelocity) + 1
        DirectionSpeed = Int(Rnd * m_MaxInitialSpread) '+ 1

        Direction = 1
      Case 2 'Down
        Velocity = m_MaxInitialVelocity + Int(Rnd * m_MaxInitialVelocity) + 1
        DirectionSpeed = Int(Rnd * m_MaxInitialSpread) - 90 '+ 1
        ' Direction = Int(Rnd * m_MaxInitialVelocity) + 1
      Case 3 'Left
        Velocity = -m_MaxInitialVelocity / 2 + Int(Rnd * m_MaxInitialVelocity) + 1
        DirectionSpeed = -Int(Rnd * m_MaxInitialSpread) - 90 '+ 1
        ' Direction = 1
    End Select

End Sub

Public Property Get Shape() As ShapeSet

    Shape = m_Shape

End Property

Public Property Let Shape(ByVal Value As ShapeSet)

    If Value = 4 Then
        m_Shape = Int(Rnd * 4)
      Else 'NOT VALUE...
        m_Shape = Value
    End If
    RefreshParticle

End Property

Public Property Get Size() As Integer

    Size = m_Size

End Property

Public Property Let Size(ByVal Value As Integer)

    m_Size = Value
    RefreshParticle

End Property

Public Property Get Top() As Long

    Top = m_Top

End Property

Public Property Let Top(ByVal Value As Long)

    m_Top = Value

End Property

Public Property Get Visible() As Boolean

    Visible = m_Visible

End Property

Public Property Let Visible(ByVal Value As Boolean)

    m_Visible = Value
    If Value = False Then
        MyShape
        GenerateParticle
    End If

End Property

Public Property Get width() As Long

    width = m_Width

End Property

Public Property Let width(ByVal Value As Long)

    m_Width = Value

End Property

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:04:30 PM) 42 + 532 = 574 Lines
