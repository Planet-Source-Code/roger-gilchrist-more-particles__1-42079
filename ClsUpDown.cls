VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsUpDown"
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
Private WithEvents TxtHidden As TextBox
Attribute TxtHidden.VB_VarHelpID = -1
Private WithEvents UD As UpDown
Attribute UD.VB_VarHelpID = -1
Private WithEvents TxtDisplay As TextBox
Attribute TxtDisplay.VB_VarHelpID = -1
Private Initialised As Boolean
Private m_Div As Double
Private m_AutoSelect As Boolean
Private m_Wrap As Boolean
Private m_Alignment As Align
Private m_AlignmentText As TextAlign
Private TwoBoxes As Boolean
Private m_BeepAtLimit As Boolean
Private StoreStr As String
Private m_Title As String
Private m_RestrictToIncrement As Boolean
Private Reciever As Label
Private RecieverSet As Boolean
Private m_LargeIncrement As Long
Private m_ID As String
Private m_ReverseOrder As Boolean
'These Enums allow you to set Alignment and AlignmentText to the standard possible values
Public Enum Align
    cc2AlignmentLeft = 0
    cc2AlignmentRight = 1
End Enum
Public Enum TextAlign
    [Left Justify]
    [Right Justify]
    Centre
End Enum
'Public Sub AssignReciever(R As Label)
'Set Reciever = R
'RecieverSet = True
'End Sub
'      random Strings
'code helps test SQL functions or other string-manipulation
'  so you can generate random strings. You can generate
'           strings with random characters and set ASCII
'  both upper And lower:

Public Property Get Alignment() As Align

  'Copyright 2003 Roger Gilchrist

    Alignment = m_Alignment

End Property

Public Property Let Alignment(ByVal Value As Align)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        m_Alignment = Value
        PlaceUD
      Else 'INITIALISED = FALSE
        NotInitialized
    End If

End Property

Public Property Get AlignmentText() As TextAlign

    AlignmentText = TxtDisplay.Alignment

End Property

Public Property Let AlignmentText(ByVal Value As TextAlign)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        m_AlignmentText = Value
        TxtDisplay.Alignment = Value
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Sub AssignControls(U As UpDown, T1 As TextBox, _
                          Min As Double, Max As Double, Value As Double, Increment As Double, _
                          Optional Wrap As Boolean = False, _
                          Optional Alignment As Align = cc2AlignmentRight, _
                          Optional AlignmentText As TextAlign = [Left Justify], _
                          Optional LargeIncrement As Long = 0, _
                          Optional Title As String, _
                          Optional T2 As TextBox, _
                          Optional LiveReciever As Label)

  'Copyright 2003 Roger Gilchrist

  'set  minor optional values

    m_ID = RandomString(Asc("A"), Asc("z"), 10, 10)
    m_Title = Title
    m_Alignment = Alignment
    m_AlignmentText = AlignmentText
    m_LargeIncrement = LargeIncrement
    'test for reverse order range to allow TExtBox to cope with
    'KeyUp etc style scrolling
    m_ReverseOrder = (Max < Min)
    m_Wrap = Wrap
    If Increment < 0 Then
        MsgBox "Increment must be a postive number.", vbCritical, "ClsUpDown: Fatal Error"
        End
    End If
    'Max and Min don't need testing because they can be in any order if Max < Min the box will run up to lower values down to higher ones
    'Value doesn't need to be tested it will be forced to Max or Min if it is outside range.
    If Increment <> Int(Increment) Then
        'calculate the ration to use between visible and hidden Min, MAx, Value and Increment properties
        m_Div = 1 / Increment
      Else 'NOT INCREMENT...
        'else Increment is either 1 or a larger whole number
        m_Div = 1
    End If
    '    m_LargeIncrement = LargeIncrement * m_Div
    'test if there is a second box
    TwoBoxes = Not (T2 Is Nothing)
    'Set the TextBox(es)
    If TwoBoxes Then
        Set TxtHidden = T2
        Set TxtDisplay = T1
        T2.Text = Value * m_Div 'Write the value to the TextBox before setting the Class Boxes
        T1.Text = Value         'makes sure that the class can't misfire because of Default Text value in TextBox

        TxtHidden.Visible = False 'Comment this out if you want to see what is in the HiddenBox as you run
        'Only works in IDE see safety system at end of routine
      Else 'TwoBoxes = FALSE
        If m_Div <> 1 Then
            'you have forgotten to include the Optional second TextBox. This is fatal. End user's will never see this msg.
            MsgBox "To use Decimal Fractional increments with this class you need to provide a Second TextBox in AssignControls", vbCritical, "ClsUpDown: Fatal Error"
            End
        End If
        T1.Text = Value
        Set TxtDisplay = T1
        TxtDisplay = Value
    End If
    TxtDisplay.Alignment = m_AlignmentText
    Initialised = True

    'set the UpDown control
    Set UD = U
    With UD
        .Min = Min * m_Div
        .Max = Max * m_Div
        .Increment = Increment * m_Div
        .Value = KeepInRange(Value, Min, Max) * m_Div
        .Wrap = Wrap
        'Link the correct TextBox
        If TwoBoxes Then
            .BuddyControl = TxtHidden
            .SyncBuddy = TxtHidden
          Else 'TwoBoxes = FALSE
            .BuddyControl = TxtDisplay
            .SyncBuddy = TxtDisplay
        End If
    End With 'UD
    PlaceUD
    RecieverSet = Not (LiveReciever Is Nothing)
    If RecieverSet Then
        Set Reciever = LiveReciever
        Reciever.Visible = False
    End If

    ''This is a debug Tool with a safety system to make sure
    ''if you comment out the line 'TxtHidden.Visible = False'
    'the following code will show the HiddenBox just above the DisplayBox
    'Make sure you remove it from final program
    ' Even if you forget to comment this out it will not run in compiled programs
    'and will hid the HiddenBox if you left it visible
    If RunningInIDE Then
        If TwoBoxes Then
            With TxtHidden
                If .Visible Then
                    .Left = TxtDisplay.Left
                    .Top = TxtDisplay.Top - .Height
                    .width = TxtDisplay.width
                End If
            End With 'TxtHidden
        End If
      Else 'RUNNINGINIDE = FALSE
        TxtHidden.Visible = False
    End If

End Sub

Public Property Get AutoSelect() As Boolean

  'Copyright 2003 Roger Gilchrist

    AutoSelect = m_AutoSelect

End Property

Public Property Let AutoSelect(ByVal Value As Boolean)

  'Copyright 2003 Roger Gilchrist

    m_AutoSelect = Value

End Property

Public Property Get BackColor() As Long

    BackColor = TxtDisplay.BackColor

End Property

Public Property Let BackColor(ByVal Value As Long)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        TxtDisplay.BackColor = Value
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Property Get BeepAtLimit() As Boolean

  'Copyright 2003 Roger Gilchrist

    BeepAtLimit = m_BeepAtLimit

End Property

Public Property Let BeepAtLimit(ByVal Value As Boolean)

    m_BeepAtLimit = Value

End Property

Public Property Get Bold() As Boolean

    Bold = TxtDisplay.FontBold

End Property

Public Property Let Bold(ByVal Value As Boolean)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        TxtDisplay.FontBold = Value
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Property Get Caption() As String

    Caption = m_Title & " " & LimitStr

End Property

Private Sub Class_Initialize()

  'Copyright 2003 Roger Gilchrist

    m_RestrictToIncrement = False
    m_AutoSelect = False
    m_Alignment = cc2AlignmentRight
    m_AlignmentText = [Left Justify]
    m_BeepAtLimit = False

End Sub

Public Property Get Divisor() As Double

    Divisor = m_Div

End Property

Public Property Get ForeColor() As Long

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        ForeColor = TxtDisplay.ForeColor
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Property Let ForeColor(ByVal Value As Long)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        TxtDisplay.ForeColor = Value
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Function FractionalUpDownSafety(Value As Single) As Single

  'Copyright 2003 Roger Gilchrist
  
  Dim TVal  As Single

    On Error Resume Next          'Copes with a "-" sign being entered by going to Min
        TVal = Value
        If m_ReverseOrder Then
            If TVal * m_Div < UD.Max Then
                TVal = UD.Max / m_Div                    'Set to Min
              ElseIf TVal * m_Div > UD.Min Then  'NOT CSNG(T.TEXT)...'NOT TVAL...
                TVal = UD.Min / m_Div                     'Set to Max
            End If
          Else 'M_REVERSEORDER = FALSE
            If TVal * m_Div < UD.Min Then
                TVal = UD.Min / m_Div                    'Set to Min
              ElseIf TVal * m_Div > UD.Max Then  'NOT CSNG(T.TEXT)...'NOT TVAL...
                TVal = UD.Max / m_Div                     'Set to Max
            End If
        End If
        'This makes sure that you cannot set a value which is off the
        'set of possible values as only members of that set will return
        'the same value for integer and regular division
        If m_RestrictToIncrement Then
            If TVal / Increment = TVal \ Increment Then
                FractionalUpDownSafety = TVal
              Else 'NOT TVAL...
                FractionalUpDownSafety = CSng(StoreStr)
            End If
          Else 'M_RESTRICTTOINCREMENT = FALSE
            FractionalUpDownSafety = TVal
        End If
    On Error GoTo 0                             ' Turn off error handler

End Function

Public Property Get HiddenIncrement() As Double

    HiddenIncrement = UD.Increment

End Property

Public Property Get HiddenMax() As Variant

    HiddenMax = UD.Max

End Property

Public Property Get HiddenMin() As Variant

    HiddenMin = UD.Min

End Property

Public Property Get HiddenValue() As Variant

    HiddenValue = UD.Value

End Property

Public Property Get ID() As String

    ID = m_ID

End Property

Public Property Get Increment() As Double

    Increment = UD.Increment / m_Div

End Property

Public Property Let Increment(ByVal Value As Double)

  'Copyright 2003 Roger Gilchrist

    If Initialised Then
        UD.Increment = Value * m_Div
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

'VB2theMax:
'KeepInRange - Ensure that a value is in a given range
'Date: 6/5/1999
'Versions: VB4 VB5 VB6 Level: Beginner
'Author: The VB2TheMax Team
'' Keep the first argument in the range [lowLimit, highLimit]
'' If the value is adjusted, the fourth (optional) argument is set to True
''
'' Note that value and limit arguments are variant, so you can use
'' this routine with any type of data.

Public Function KeepInRange(Value As Variant, lowLimit As Variant, highLimit As Variant, Optional OutOfRange As Boolean) As Variant

    If Value < lowLimit Then
        KeepInRange = lowLimit
        OutOfRange = True
      ElseIf Value > highLimit Then 'NOT VALUE...
        KeepInRange = highLimit
        OutOfRange = True
      Else 'NOT VALUE...
        KeepInRange = Value
        OutOfRange = False
    End If

End Function


Public Property Get LargeIncrement() As Variant

    LargeIncrement = m_LargeIncrement

End Property

Public Property Get LimitStr() As Variant

  'Copyright 2003 Roger Gilchrist

    LimitStr = "[" & UD.Min / m_Div & " to " & UD.Max / m_Div & "]"

End Property

Public Property Get LiveUpdate() As Boolean

    LiveUpdate = RecieverSet

End Property

Public Property Let Max(ByVal Value As Variant)

    Max = Value * m_Div

End Property

Public Property Get Max() As Variant

    Max = UD.Max / m_Div

End Property

Public Property Let Min(ByVal Value As Variant)

    Min = Value * m_Div

End Property

Public Property Get Min() As Variant

    Min = UD.Min / m_Div

End Property

Private Sub NotInitialized()

  'Copyright 2003 Roger Gilchrist
  'Force you to program properly

    MsgBox "You have to call AssignControls first.", vbCritical, "ClsUpDown Coding Error"

End Sub

Private Sub PlaceUD()

  'Copyright 2003 Roger Gilchrist
  'This routine simplifies the way the UpDown and TextBox relate
  'To use it simply set the TextBox to cover the area you want
  'to be occupied by the 2 controls. This routine takes care of the rest

    With UD
        If TwoBoxes Then
            'cannot use standard Alignment because the UpDown is not linked
            'to the Display TextBox
            .Top = TxtDisplay.Top ' set vertical position
            .Height = TxtDisplay.Height
            TxtDisplay.width = TxtDisplay.width - .width 'shrink TextBox to make space
            If m_Alignment = cc2AlignmentRight Then ' Set Horz position of UpDown
                .Left = TxtDisplay.width + TxtDisplay.Left
              Else 'NOT M_ALIGNMENT...
                .Left = TxtDisplay.Left                  'Place UpDown then
                TxtDisplay.Left = TxtDisplay.Left + .width 'move TextBox to right
            End If
          Else 'TwoBoxes = FALSE
            .Top = TxtDisplay.Top
            .Height = TxtDisplay.Height
            .Alignment = m_Alignment
        End If
    End With 'UD

End Sub

Private Function RandomString(iLowerBoundAscii As Integer, iUpperBoundAscii As Integer, lLowerBoundLength As Long, lUpperBoundLength As Long) As String

  'don't know where I got this but the style
  '(very long parameter names with inital letter Type indicator) suggests VBPJ
  
  Dim sHoldString As String
  Dim LLength As Long
  Dim LCount As Long

    'Verify boundaries
    If iLowerBoundAscii < 0 Then
        iLowerBoundAscii = 0
    End If
    If iLowerBoundAscii > 255 Then
        iLowerBoundAscii = 255
    End If
    If iUpperBoundAscii < 0 Then
        iUpperBoundAscii = 0
    End If
    If iUpperBoundAscii > 255 Then
        iUpperBoundAscii = 255
    End If
    If lLowerBoundLength < 0 Then
        lLowerBoundLength = 0
    End If

    'Set a random length
    LLength = Int((CDbl(lUpperBoundLength) - _
              CDbl(lLowerBoundLength) + _
              1) * Rnd + lLowerBoundLength)

    'Create the random string
    For LCount = 1 To LLength
        sHoldString = sHoldString & _
                      Chr$(Int((iUpperBoundAscii - iLowerBoundAscii _
                      + 1) * Rnd + iLowerBoundAscii))
    Next LCount
    RandomString = sHoldString
    '      --Eric Lynn, Ballwin, Missouri

End Function

Public Property Let RestrictToIncrement(ByVal Value As Boolean)

    m_RestrictToIncrement = Value

End Property

Public Property Get RestrictToIncrement() As Boolean

    RestrictToIncrement = m_RestrictToIncrement

End Property

Private Function RunningInIDE() As Boolean

  'Don't remember where I got this but it is very useful

    On Error GoTo RunningInIDEErr
    Debug.Print 1 / 0  'Divide by zero (fails within IDE)

Exit Function  'Exit if no error

RunningInIDEErr:
    RunningInIDE = True 'We get error if Debug.Print was evaluated
    On Error GoTo 0          'Turn off error handler

End Function

Public Property Let Title(ByVal Value As String)

    m_Title = Value

End Property

Public Property Get Title() As String

    Title = m_Title

End Property

Private Sub TxtDisplay_Change()

    If RecieverSet Then
        Reciever.Tag = m_ID
        Reciever.Caption = Value

    End If

End Sub

Private Sub TxtDisplay_GotFocus()

    If m_AutoSelect Then
        TxtDisplay.SelStart = 0
        TxtDisplay.SelLength = Len(TxtDisplay.Text)
    End If

End Sub

Private Sub TxtDisplay_KeyDown(KeyCode As Integer, Shift As Integer)

    StoreStr = TxtDisplay.Text

End Sub

Private Sub TxtDisplay_KeyUp(KeyCode As Integer, Shift As Integer)

  'Copyright 2003 Roger Gilchrist
  'Allow focus box to scroll with Up, Down, PgUp, PgDn, Home and End Keys

    If KeyCode = vbKeyUp Then
        KeyCode = 0
        TxtDisplay.Text = FractionalUpDownSafety(TxtDisplay.Text + Increment)
      ElseIf KeyCode = vbKeyDown Then 'NOT KEYCODE...
        KeyCode = 0
        TxtDisplay.Text = FractionalUpDownSafety(TxtDisplay.Text - Increment)
    End If
    If m_LargeIncrement <> 0 Then
        If KeyCode = vbKeyPageUp Then
            KeyCode = 0
            TxtDisplay.Text = FractionalUpDownSafety(TxtDisplay.Text + Increment * m_LargeIncrement)
          ElseIf KeyCode = vbKeyPageDown Then 'NOT KEYCODE...
            KeyCode = 0
            TxtDisplay.Text = FractionalUpDownSafety(TxtDisplay.Text - Increment * m_LargeIncrement)
        End If
      Else 'NOT M_LARGEINCREMENT...
        If KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown Then
            KeyCode = 0
            Exit Sub '>---> Bottom
        End If
    End If
    If KeyCode = vbKeyHome Then
        KeyCode = 0
        TxtDisplay.Text = FractionalUpDownSafety(Max)

      ElseIf KeyCode = vbKeyEnd Then 'NOT KEYCODE...
        KeyCode = 0
        TxtDisplay.Text = FractionalUpDownSafety(Min)
    End If

    TxtDisplay.Text = FractionalUpDownSafety(TxtDisplay.Text)
    If TwoBoxes Then
        TxtHidden.Text = CLng(FractionalUpDownSafety(TxtDisplay.Text) * m_Div)
        UD.Value = TxtHidden.Text
      Else 'TWOBOXES = FALSE
        UD.Value = TxtDisplay.Text
    End If
    TxtDisplay_Change

End Sub

Private Sub TxtHidden_Change()

    TxtDisplay.Text = FractionalUpDownSafety(TxtHidden.Text / m_Div)

End Sub

Private Sub UD_Change()

    If TwoBoxes Then
        TxtHidden.Text = UD.Value

      Else 'TwoBoxes = FALSE
        TxtDisplay.Text = Value
    End If

End Sub

Private Sub UD_DownClick()

    If m_Wrap = False Then
        If UD.Value = UD.Min Then
            Beep
        End If
    End If

End Sub

Private Sub UD_UpClick()

    If m_Wrap = False Then
        If UD.Value = UD.Max Then
            Beep
        End If
    End If

End Sub

Public Function UpDownSafetyTxt(Txt As TextBox) As Integer

  Dim TVal As Single

    On Error Resume Next                         'Copes with a "-" sign being entered by going to Min
        TVal = CSng(Txt.Text)
        If TVal < UD.Min Then
            TVal = UD.Min                    'Set to Min
          ElseIf TVal > UD.Max Then 'NOT CSNG(T.TEXT)...'NOT TVAL...
            TVal = UD.Max                    'Set to Max
        End If
        UpDownSafetyTxt = TVal
    On Error GoTo 0                             ' Turn off error handler

End Function

Public Property Let Value(ByVal Value As Variant)

    If Initialised Then
        UD.Value = Value * m_Div
      Else 'INITIALISED = FALSE
        NotInitialized

    End If

End Property

Public Property Get Value() As Variant

    If Initialised Then
        Value = UD.Value / m_Div
    End If

End Property

Public Property Let Wrap(ByVal Value As Boolean)

    m_Wrap = Value

End Property

Public Property Get Wrap() As Boolean

    Wrap = m_Wrap

End Property

':) Ulli's VB Code Formatter V2.13.6 (2/01/2003 3:02:56 PM) 44 + 699 = 743 Lines
