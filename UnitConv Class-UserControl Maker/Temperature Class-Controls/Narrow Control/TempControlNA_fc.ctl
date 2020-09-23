VERSION 5.00
Begin VB.UserControl TempControlNA
   ClientHeight    =   915
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   915
   ScaleWidth      =   2595
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1560
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   0
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      MousePointer    =   9  'Size W E
      ScaleHeight     =   615
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "TempControlNA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'///////[ Simple Unit Conversion User-Control ]////////////

Option Explicit

'UnitElement classes
'///////[ NOTE: ON 'COMPILE ERROR' ]//////
'//  If you get a Complie error:        //
'//  'User-defined type not defined'    //
'//  here, then you forgot to add the   //
'//  UnitElement.cls class file to      //
'//  your project.                      //
'///////[ END 'COMPILE ERROR' NOTE ]//////
Private Unit_cur_Fahrenheit As New UnitElementCls
Private Unit_cur_Celsius As New UnitElementCls
    
'Default Property Values:
Private Const m_def_ctl_FormatSci = 0
Private Const m_def_ctl_sigDigits = 5

'Property Variables:
Private m_ctl_FormatSci As Boolean
Private m_ctl_sigDigits As Integer

Private UnitCol As New Collection

Private ScriptCtl As Control
Private ScriptControlRegistered As Boolean

'Event Declarations:
Public Event TboxKeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress

'Local Variables:
Private m_picMouseDown As Boolean

Public Sub RegisterScriptControl(sc As Control)
    Set ScriptCtl = sc
    ScriptControlRegistered = True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
    Set Font = Combo1.Font
End Property
Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Set Combo1.Font = New_Font
    Set Text1.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,5
Public Property Get ctl_sigDigits() As Integer
    ctl_sigDigits = m_ctl_sigDigits
End Property
Public Property Let ctl_sigDigits(ByVal New_ctl_sigDigits As Integer)
    m_ctl_sigDigits = New_ctl_sigDigits
    'bounds check
    If m_ctl_sigDigits < 1 Then m_ctl_sigDigits = 1
    If m_ctl_sigDigits > 14 Then m_ctl_sigDigits = 14
    UpdateUnits "°Fahrenheit", Unit_cur_Fahrenheit.curVal
    PropertyChanged "ctl_sigDigits"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ctl_FormatSci() As Boolean
    ctl_FormatSci = m_ctl_FormatSci
End Property
Public Property Let ctl_FormatSci(ByVal New_ctl_FormatSci As Boolean)
    m_ctl_FormatSci = New_ctl_FormatSci
    UpdateUnits "°Fahrenheit", Unit_cur_Fahrenheit.curVal
    PropertyChanged "ctl_FormatSci"
End Property

Public Property Get Enabled() As Boolean
    Enabled = Combo1.Enabled
End Property
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Combo1.Enabled() = New_Enabled
    Text1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = Picture1.BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Picture1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
    ListIndex = Combo1.ListIndex
End Property
Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Dim tmpInt As Integer
    'bounds check
    tmpInt = New_ListIndex
    If New_ListIndex < 0 Then tmpInt = 0
    If New_ListIndex > Combo1.ListCount - 1 Then
        tmpInt = Combo1.ListCount - 1
    End If
    Combo1.ListIndex() = tmpInt
    PropertyChanged "ListIndex"
End Property

'//////////////[ Unit Specific Properties ]//////////////////

'=================[ °Fahrenheit ]
Public Property Get cur_Fahrenheit() As Double
    cur_Fahrenheit = Unit_cur_Fahrenheit.curVal
End Property
Public Property Let cur_Fahrenheit(ByVal New_cur_Fahrenheit As Double)
    Unit_cur_Fahrenheit.curVal = New_cur_Fahrenheit
    UpdateUnits "°Fahrenheit", Unit_cur_Fahrenheit.curVal
    PropertyChanged "cur_Fahrenheit"
End Property
Public Property Get cur_Fahrenheit_err() As Boolean
    cur_Fahrenheit_err = Unit_cur_Fahrenheit.bError
End Property

'=================[ °Celsius ]
Public Property Get cur_Celsius() As Double
    cur_Celsius = Unit_cur_Celsius.curVal
End Property
Public Property Let cur_Celsius(ByVal New_cur_Celsius As Double)
    Unit_cur_Celsius.curVal = New_cur_Celsius
    UpdateUnits "°Celsius", Unit_cur_Celsius.curVal
    PropertyChanged "cur_Celsius"
End Property
Public Property Get cur_Celsius_err() As Boolean
    cur_Celsius_err = Unit_cur_Celsius.bError
End Property

'////////////////////////////////////////////////////////////

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        m_picMouseDown = True
    End If
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If m_picMouseDown Then
        '15*30 < X+30 < uc.wid - 15*30
        If 15 * 30 < X + 30 And X + 30 < UserControl.Width - 15 * 30 Then
            Combo1.Left = X + 30
        End If
        If UserControl.Width - Combo1.Left > 15 * 30 Then
            Combo1.Width = UserControl.Width - Combo1.Left
            PropertyChanged "Combo1_Width"
        End If
        If Combo1.Left > 15 * 4 Then Text1.Width = Combo1.Left - 15 * 3
    End If
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    m_picMouseDown = False
End Sub

Private Sub UserControl_Resize()
    Picture1.Width = UserControl.Width
    Picture1.Height = UserControl.Height
    Text1.Height = Combo1.Height
    Combo1.Left = UserControl.Width - Combo1.Width
    If Combo1.Left > 15 * 4 Then Text1.Width = Combo1.Left - 15 * 3
    UserControl.Height = Combo1.Height
End Sub

'Initialize Properties for User Control
Private Sub UserControl_Initialize()
    'set default values
    Unit_cur_Fahrenheit.defVal = 68#
    Unit_cur_Celsius.defVal = 20#
    'load combobox and unit collection
    LoadUnitColAndComboBox
End Sub

'Initialize Properties for User Control Creation
Private Sub UserControl_InitProperties()
    m_ctl_sigDigits = m_def_ctl_sigDigits
    m_ctl_FormatSci = m_def_ctl_FormatSci
    Combo1.Left = 2280
    Combo1.ListIndex = 0
    'set current values to default values'
    Unit_cur_Fahrenheit.curVal = Unit_cur_Fahrenheit.defVal
    Unit_cur_Celsius.curVal = Unit_cur_Celsius.defVal
    StartUp
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ctl_sigDigits = PropBag.ReadProperty("ctl_sigDigits", m_def_ctl_sigDigits)
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", 0)
    m_ctl_FormatSci = PropBag.ReadProperty("ctl_FormatSci", m_def_ctl_FormatSci)
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    Picture1.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Combo1.Width = PropBag.ReadProperty("Combo1_Width", 1335)
    'set current values to last-saved values
    Unit_cur_Fahrenheit.curVal = PropBag.ReadProperty("cur_Fahrenheit", Unit_cur_Fahrenheit.defVal)
    Unit_cur_Celsius.curVal = PropBag.ReadProperty("cur_Celsius", Unit_cur_Celsius.defVal)
    StartUp
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
    Call PropBag.WriteProperty("ctl_sigDigits", m_ctl_sigDigits, m_def_ctl_sigDigits)
    Call PropBag.WriteProperty("ctl_FormatSci", m_ctl_FormatSci, m_def_ctl_FormatSci)
    Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
    Call PropBag.WriteProperty("BackColor", Picture1.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Combo1_Width", Combo1.Width, 1335)
    'save current values to storage
    Call PropBag.WriteProperty("cur_Fahrenheit", Unit_cur_Fahrenheit.curVal, Unit_cur_Fahrenheit.defVal)
    Call PropBag.WriteProperty("cur_Celsius", Unit_cur_Celsius.curVal, Unit_cur_Celsius.defVal)
End Sub

Private Sub StartUp()
    Set Text1.Font = Combo1.Font
    UserControl_Resize
    Combo1_Click
End Sub

Private Sub LoadUnitColAndComboBox()
    Dim ndx As Integer
    'load unit collection
    Unit_cur_Fahrenheit.uName = "°Fahrenheit": UnitCol.Add Unit_cur_Fahrenheit, "°Fahrenheit"
    Unit_cur_Celsius.uName = "°Celsius": UnitCol.Add Unit_cur_Celsius, "°Celsius"
    'load combobox
    Combo1.Clear
    If UnitCol.Count <> 0 Then
        For ndx = 1 To UnitCol.Count
            Combo1.AddItem " " & UnitCol(ndx).uName
        Next
    End If
End Sub

Private Sub Combo1_Click()
    'Update text display
    If UnitCol(Combo1.ListIndex + 1).bError = True Then
        Text1.Text = "<range>"
    Else
        Text1.Text = getFormat(SigFigs(UnitCol(Combo1.ListIndex + 1).curVal, m_ctl_sigDigits))
    End If
    PropertyChanged "ListIndex"
End Sub

Private Sub UpdateUnits(unitName As String, newVal As Double)
    Dim nameStr As String
    Dim tmpDbl As Double
    
    On Error GoTo numError
    Select Case unitName
        Case Is = "°Fahrenheit"
            nameStr = "°Fahrenheit": UnitCol(nameStr).bError = False
            tmpDbl = newVal: setRange tmpDbl, nameStr
            UnitCol("°Fahrenheit").curVal = tmpDbl
            If UnitCol("°Fahrenheit").curVal < -459.67 Then UnitCol("°Fahrenheit").curVal = -459.67
            nameStr = "°Celsius": UnitCol(nameStr).bError = False
            tmpDbl = (UnitCol("°Fahrenheit").curVal - 32) * 5 / 9
            setRange tmpDbl, nameStr
            UnitCol("°Celsius").curVal = tmpDbl
        Case Is = "°Celsius"
            nameStr = "°Celsius": UnitCol(nameStr).bError = False
            tmpDbl = newVal: setRange tmpDbl, nameStr
            UnitCol("°Celsius").curVal = tmpDbl
            If UnitCol("°Celsius").curVal < -273.15 Then UnitCol("°Celsius").curVal = -273.15
            nameStr = "°Fahrenheit": UnitCol(nameStr).bError = False
            tmpDbl = UnitCol("°Celsius").curVal * 9 / 5 + 32
            setRange tmpDbl, nameStr
            UnitCol("°Fahrenheit").curVal = tmpDbl
    End Select
    Combo1_Click
Exit Sub
numError:
    'Debug.Print Err.Description
    UnitCol(nameStr).bError = True
    If Sgn(tmpDbl) = 1 Then
        tmpDbl = 1E+308
    Else
        tmpDbl = -1E+308
    End If
    Resume Next
End Sub
Private Sub setRange(ByRef dblNum As Double, nameStr As String)
    If dblNum > -1E-308 And dblNum < 0 Then
        dblNum = 0
    End If
    If dblNum < 1E-308 And dblNum > 0 Then
        dblNum = 0
    End If
    If dblNum > 1E+308 Then
        dblNum = 1E+308
        UnitCol(nameStr).bError = True
    End If
    If dblNum < -1E+308 Then
        dblNum = -1E+308
        UnitCol(nameStr).bError = True
    End If
End Sub

'================================================================
'///////////////[ Begin Input Processing ]///////////////////////
'================================================================
'The input processing code below is designed such as to allow
'the use of an external script control to evaluate any entered
'script.
'////////////////////////////////////////////////////////////////
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim expr As String
    Dim result As String
    Dim msgTitle As String
    Dim msgPrompt As String
    Dim KeyAsc As Integer
    Dim tmpDbl As Double
    Dim negFlag As Boolean
    
    'Pass KeyAscii value to external KeyPress() event:
    KeyAsc = KeyAscii
    'Change BackColor to yellow for editing.
    Text1.BackColor = vbYellow
    If KeyAscii = 13 Then '[Enter] key.
        expr = Text1.Text
        If ScriptControlRegistered Then
            'Get any mathmatical expression in the textbox.
            'expr = "Log(10) + (pi * 10) + 10"
            'Calclate result of the expression:
            On Error GoTo scriptError
            result = ScriptCtl.Eval(expr)
            On Error GoTo 0
            'no error occured, continue...
        Else
            'no script control
            result = expr
        End If
        'Supress [Enter] key 'beep'.
        KeyAscii = 0
        'Change BackColor back to normal.
        Text1.BackColor = vbWindowBackground
        'Update the other units.
        If Left(result, 1) = "-" Then
            negFlag = True
        Else
            negFlag = False
        End If
        On Error GoTo numError
        tmpDbl = CDbl(result)
        If tmpDbl > -1E-308 And tmpDbl < 0 Then
            tmpDbl = 0
        End If
        If tmpDbl < 1E-308 And tmpDbl > 0 Then
            tmpDbl = 0
        End If
        If tmpDbl > 1E+308 Then
            tmpDbl = 1E+308
            UnitCol(Combo1.ListIndex + 1).bError = True
        End If
        If tmpDbl < -1E+308 Then
            tmpDbl = -1E+308
            UnitCol(Combo1.ListIndex + 1).bError = True
        End If
        UnitCol(Combo1.ListIndex + 1).curVal = tmpDbl
        UpdateUnits UnitCol(Combo1.ListIndex + 1).uName, tmpDbl
        PropertyChanged "cur_m"
    End If
    'Pass KeyAscii value to external KeyPress() event:
    RaiseEvent TboxKeyPress(KeyAsc)
Exit Sub
scriptError: 'A script control error occured
    'Inform user of the error:
    msgTitle = ScriptCtl.Error.Source & ": " & _
    Str$(ScriptCtl.Error.Number)
    msgPrompt = ScriptCtl.Error.Description
    'Show user the error information:
    MsgBox msgPrompt, , msgTitle
    'Pass KeyAscii value to external KeyPress() event:
    'RaiseEvent TboxKeyPress(KeyAsc)
Exit Sub
numError:
    'Debug.Print Err.Description
    UnitCol(Combo1.ListIndex + 1).bError = True
    If negFlag = False Then
        tmpDbl = 1E+308
    Else
        tmpDbl = -1E+308
    End If
    Resume Next
End Sub
'===============================================================
'/////////////////[ End Input Processing ]//////////////////////
'===============================================================

'toggle sci/no formatting
Private Sub Text1_DblClick()
    m_ctl_FormatSci = Not m_ctl_FormatSci
    Combo1_Click
End Sub

'===============================================================
'Return 'dblNumber' rounded to 'intSF' significant figures
'===============================================================
Private Function SigFigs(dblNumber As Double, intSF As Integer) As Double
'Only works properly for doubles in the range: (+/-)1E(+/-)308
    Dim negFlag As Integer
    Dim tmpDbl As Double
    Dim factor As Double
    Dim dblA As Double
    Dim dblB As Double
    Dim outNum As Double
    
    'dblNumber = 0 ?
    If dblNumber <> 0 Then
        'make sign of tmpDbl <- dblNumber, be positive
        If dblNumber < 0 Then
            tmpDbl = -dblNumber: negFlag = -1
        Else
            tmpDbl = dblNumber: negFlag = 0
        End If
        'get multiplication/division order-of-magnitude factor
        factor = 10 ^ (Int(Log(tmpDbl) / Log(10)) + 1)
        'dblA = tmpDbl's significant digits moved to right of
        'decimal point: 0.########
        dblA = tmpDbl / factor
        'correct dblA for sign if necessary
        If negFlag Then dblA = -dblA
        'round dblA to intSF number of decimal places
        dblB = Round(dblA, intSF)
        'restore dblB to tmpDbl's original order-of-magnitude
        outNum = dblB * factor 'outNum = (positive/negative)
        'Debug.Print tmpDbl, factor, dblA, dblB, outNum
    Else  'dblNumber = 0
        outNum = 0
    End If
    SigFigs = outNum 'return
End Function
'///////////////////////////////////////////////////////////////

Private Function getFormat(Number As Double) As String
    If m_ctl_FormatSci Then
        'Scientific format
        getFormat = FormatSci(Number, 14)
    Else
        'No Formating
        getFormat = Str(Number)
    End If
End Function

Private Function FormatSci(Number As Double, Optional _
                          DecimalPlaces As Long = 1) As String
    Dim Exponent As Long
    Dim Parts() As String
    
    If Abs(Number) < 1000 And Abs(Number) >= 1 Then
        FormatSci = Format(Number, "0.0#############")
        Exit Function
    End If
    If Number = 0 Then
        FormatSci = Format(Number, "0.0")
        Exit Function
    End If
    Parts = Split(Format(Number, "0.0#############E+0"), "E")
    Exponent = 1 * Int(Parts(1) / 1)
    FormatSci = Format(Parts(0) * 10 ^ (Parts(1) - Exponent), _
                               "0.0" & String(DecimalPlaces, "#")) & _
                               "E" & Format(Exponent, "+0;-0")
End Function

'Handy External Program Routines:
'Form_Load()
'   SetupScriptControl ScriptControl1
'End Sub
'Private Sub SetupScriptControl(sc As ScriptControl)
'    Dim myCode As String
'    'Initialize a language engine for the script control:
'    sc.Language = "VBScript"
'    'Add any variables that you want the script control to
'    'know about:
'    'set pi as a known variable
'    sc.ExecuteStatement "pi = 3.1415926535898"
'    'Add any additional functions that you want the script
'    'control to know about...
'    '/// log10(x) /// log to the base 10
'    myCode = _
'    "Function log10(var)" + vbCrLf + _
'    " log10 = log(var) / log(10)" + vbCrLf + _
'    "End Function"
'    sc.AddCode myCode
'    '/// ln(x) /// natural log, to base e
'    myCode = _
'    "Function ln(var)" + vbCrLf + _
'    " ln = log(var)" + vbCrLf + _
'    "End Function"
'    sc.AddCode myCode
'    '/// d2r(x) /// degrees to radians
'    myCode = _
'    "Function d2r(var)" + vbCrLf + _
'    " d2r = pi/180*(var)" + vbCrLf + _
'    "End Function"
'    sc.AddCode myCode
'    '/// r2d(x) /// radians to degrees
'    myCode = _
'    "Function r2d(var)" + vbCrLf + _
'    " r2d = 180/pi*(var)" + vbCrLf + _
'    "End Function"
'    sc.AddCode myCode
'    'NOTE: These functions can be evaluated directly from
'    'within a textbox because they evaluate to a number.
'End Sub
''Register the ScriptControl with unit-conversion controls:
'Private Sub RegScriptControl(sc As ScriptControl)
'    UnitControl1.RegisterScriptControl sc
'    UnitControl2.RegisterScriptControl sc
'End Sub
''Pop up a message box informing user of available math
''scripting inputs for Text1:
'Private Sub ShowScriptControlMathInputHelp()
'    Dim msg As String
'    Dim msgTitle As String
'
'    msgTitle = "   Math Input Functions:"
'    msg = "Standard Functions:" & vbTab & vbTab & "Additional Functions:" & vbCrLf
'    msg = msg & "Trigonometric: (radian mode)" & vbTab & vbTab & "log10(x) - base 10 logarithm" & vbCrLf
'    msg = msg & "atn(x) - inverse tangent" & vbTab & vbTab & "ln(x) - natural logarithm" & vbCrLf
'    msg = msg & "sin(x) - sine" & vbTab & vbTab & vbTab & "d2r(x) - degrees to radians" & vbCrLf
'    msg = msg & "cos(x) - cosine" & vbTab & vbTab & vbTab & "r2d(x) - radians to degrees" & vbCrLf
'    msg = msg & "tan(x) - tangent" & vbTab & vbTab & vbTab & "pi - 3.1415926535898" & vbCrLf & vbCrLf
'    msg = msg & "Standard:" & vbTab & vbTab & vbTab & "Standard Operators:" & vbCrLf
'    msg = msg & "exp(x) - exponential" & vbTab & vbTab & "(+) - addition" & vbTab & "(-) - subtraction" & vbCrLf
'    msg = msg & "log(x) - natural logarithm" & vbTab & vbTab & "(*) - multiplication" & vbTab & "(/) - division" & vbCrLf
'    msg = msg & "sqr(x) - square root" & vbTab & vbTab & vbTab & "(^) - exponetation" & vbTab & "( ) - parentheses" & vbCrLf & vbCrLf
'    msg = msg & "Example Input:    (pi*4^2)+2" & vbTab & vbTab & "Note:  X^(1/3) = cube root of X" & vbCrLf
'    msg = msg & "Resolves To:    52.265482457437"
'    'Show help message:
'    MsgBox msg, , msgTitle
'End Sub
''///////////////////////////////////////////////////////////////
''===============================================================
''Return 'dblNumber' rounded to 'intSF' significant figures
''===============================================================
'Private Function SigFigs(dblNumber As Double, intSF As Integer) As Double
''Only works properly for doubles in the range: (+/-)1E(+/-)308
'    Dim negFlag As Integer
'    Dim tmpDbl As Double
'    Dim factor As Double
'    Dim dblA As Double
'    Dim dblB As Double
'    Dim outNum As Double
'
'    'dblNumber = 0 ?
'    If dblNumber <> 0 Then
'        'make sign of tmpDbl <- dblNumber, be positive
'        If dblNumber < 0 Then
'            tmpDbl = -dblNumber: negFlag = -1
'        Else
'            tmpDbl = dblNumber: negFlag = 0
'        End If
'        'get multiplication/division order-of-magnitude factor
'        factor = 10 ^ (Int(Log(tmpDbl) / Log(10)) + 1)
'        'dblA = tmpDbl's significant digits moved to right of
'        'decimal point: 0.########
'        dblA = tmpDbl / factor
'        'correct dblA for sign if necessary
'        If negFlag Then dblA = -dblA
'        'round dblA to intSF number of decimal places
'        dblB = Round(dblA, intSF)
'        'restore dblB to tmpDbl's original order-of-magnitude
'        outNum = dblB * factor 'outNum = (positive/negative)
'        'Debug.Print tmpDbl, factor, dblA, dblB, outNum
'    Else  'dblNumber = 0
'        outNum = 0
'    End If
'    SigFigs = outNum 'return
'End Function
''///////////////////////////////////////////////////////////////
'Private Function FormatSci(Number As Double, Optional _
'                          DecimalPlaces As Long = 1) As String
'    Dim Exponent As Long
'    Dim Parts() As String
'
'    If Abs(Number) < 1000 And Abs(Number) >= 1 Then
'        FormatSci = Format(Number, "0.0#############")
'        Exit Function
'    End If
'    If Number = 0 Then
'        FormatSci = Format(Number, "0.0")
'        Exit Function
'    End If
'    Parts = Split(Format(Number, "0.0#############E+0"), "E")
'    Exponent = 1 * Int(Parts(1) / 1)
'    FormatSci = Format(Parts(0) * 10 ^ (Parts(1) - Exponent), _
'                               "0.0" & String(DecimalPlaces, "#")) & _
'                               "E" & Format(Exponent, "+0;-0")
'End Function
'Private Function FormatEng(Number As Double, Optional _
'                          DecimalPlaces As Long = 1) As String
'    Dim Exponent As Long
'    Dim Parts() As String
'
'    If Abs(Number) < 1000 And Abs(Number) >= 1 Then
'        FormatEng = Format(Number, "0.0#############")
'        Exit Function
'    End If
'    If Number = 0 Then
'        FormatEng = Format(Number, "0.0")
'        Exit Function
'    End If
'    Parts = Split(Format(Number, "0.0#############E+0"), "E")
'    Exponent = 3 * Int(Parts(1) / 3)
'    FormatEng = Format(Parts(0) * 10 ^ (Parts(1) - Exponent), _
'                               "0.0" & String(DecimalPlaces, "#")) & _
'                               "E" & Format(Exponent, "+0;-0")
'End Function

