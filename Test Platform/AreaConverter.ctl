VERSION 5.00
Begin VB.UserControl AreaConverter
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3600
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   1080
   ScaleWidth      =   3600
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   2250
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "AreaConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'///////[ Simple Unit Conversion User-Control ]////////////

Option Explicit

'UnitElement classes
'///////[ NOTE: ON 'COMPILE ERROR' ]/////////////////////////////////
'//                                                                //
'//  If you get a Complie error: 'User-defined type not defined'   //
'//  here, then you forgot to add the UnitElement.cls class        //
'//  file to your project.                                         //
'//                                                                //
'///////[ END 'COMPILE ERROR' NOTE ]/////////////////////////////////
Private Unit_cur_yd_sqr As New UnitElementCls
Private Unit_cur_ft_sqr As New UnitElementCls
Private Unit_cur_in_sqr As New UnitElementCls
Private Unit_cur_m_sqr As New UnitElementCls
Private Unit_cur_cm_sqr As New UnitElementCls
Private Unit_cur_mm_sqr As New UnitElementCls
    
'Default Property Values:
Private Const m_def_ctl_FormatSci = 0
Private Const m_def_ctl_ListIndex = 0
Private Const m_def_ctl_sigDigits = 5
Private Const m_def_ctl_tboxWidth = 150

'Property Variables:
Private m_ctl_FormatSci As Boolean
Private m_ctl_ListIndex As Integer
Private m_ctl_sigDigits As Integer
Private m_ctl_tboxWidth As Integer

Private UnitCol As New Collection

Private ScriptCtl As Control
Private ScriptControlRegistered As Boolean

'Event Declarations:
Public Event TboxKeyPress(KeyAscii As Integer) 'MappingInfo=Text1,Text1,-1,KeyPress

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
    Set Label1.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get ctl_ListIndex() As Integer
    ctl_ListIndex = m_ctl_ListIndex
End Property

Public Property Let ctl_ListIndex(ByVal New_ctl_ListIndex As Integer)
    m_ctl_ListIndex = New_ctl_ListIndex
    'bounds check
    If m_ctl_ListIndex < 0 Then m_ctl_ListIndex = 0
    If m_ctl_ListIndex > UnitCol.Count - 1 Then
        m_ctl_ListIndex = UnitCol.Count - 1
    End If
    'set list index
    Combo1.ListIndex = m_ctl_ListIndex
    PropertyChanged "ctl_ListIndex"
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
    UpdateUnits "yd^2", Unit_cur_yd_sqr.curVal
    PropertyChanged "ctl_sigDigits"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,3,0,150
Public Property Get ctl_tboxWidth() As Integer
    ctl_tboxWidth = m_ctl_tboxWidth
End Property
Public Property Let ctl_tboxWidth(ByVal New_ctl_tboxWidth As Integer)
    'Text1.Width Bounds check
    m_ctl_tboxWidth = New_ctl_tboxWidth '<-in pixels
    If m_ctl_tboxWidth < 22 Then
        m_ctl_tboxWidth = 22
    End If
    If m_ctl_tboxWidth > ScaleX(UserControl.Width, vbTwips, vbPixels) - 42 Then
        m_ctl_tboxWidth = ScaleX(UserControl.Width, vbTwips, vbPixels) - 42
    End If
    'set Text1 width (in twips)
    Text1.Width = ScaleX(m_ctl_tboxWidth, vbPixels, vbTwips)
    UpdateUnits "yd^2", Unit_cur_yd_sqr.curVal
    PropertyChanged "ctl_tboxWidth"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get ctl_FormatSci() As Boolean
    ctl_FormatSci = m_ctl_FormatSci
End Property
Public Property Let ctl_FormatSci(ByVal New_ctl_FormatSci As Boolean)
    m_ctl_FormatSci = New_ctl_FormatSci
    UpdateUnits "yd^2", Unit_cur_yd_sqr.curVal
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

'//////////////[ Unit Specific Properties ]//////////////////

'=================[ yd^2 ]
Public Property Get cur_yd_sqr() As Double
    cur_yd_sqr = Unit_cur_yd_sqr.curVal
End Property
Public Property Let cur_yd_sqr(ByVal New_cur_yd_sqr As Double)
    Unit_cur_yd_sqr.curVal = New_cur_yd_sqr
    UpdateUnits "yd^2", Unit_cur_yd_sqr.curVal
    PropertyChanged "cur_yd_sqr"
End Property
Public Property Get cur_yd_sqr_err() As Boolean
    cur_yd_sqr_err = Unit_cur_yd_sqr.bError
End Property

'=================[ ft^2 ]
Public Property Get cur_ft_sqr() As Double
    cur_ft_sqr = Unit_cur_ft_sqr.curVal
End Property
Public Property Let cur_ft_sqr(ByVal New_cur_ft_sqr As Double)
    Unit_cur_ft_sqr.curVal = New_cur_ft_sqr
    UpdateUnits "ft^2", Unit_cur_ft_sqr.curVal
    PropertyChanged "cur_ft_sqr"
End Property
Public Property Get cur_ft_sqr_err() As Boolean
    cur_ft_sqr_err = Unit_cur_ft_sqr.bError
End Property

'=================[ in^2 ]
Public Property Get cur_in_sqr() As Double
    cur_in_sqr = Unit_cur_in_sqr.curVal
End Property
Public Property Let cur_in_sqr(ByVal New_cur_in_sqr As Double)
    Unit_cur_in_sqr.curVal = New_cur_in_sqr
    UpdateUnits "in^2", Unit_cur_in_sqr.curVal
    PropertyChanged "cur_in_sqr"
End Property
Public Property Get cur_in_sqr_err() As Boolean
    cur_in_sqr_err = Unit_cur_in_sqr.bError
End Property

'=================[ m^2 ]
Public Property Get cur_m_sqr() As Double
    cur_m_sqr = Unit_cur_m_sqr.curVal
End Property
Public Property Let cur_m_sqr(ByVal New_cur_m_sqr As Double)
    Unit_cur_m_sqr.curVal = New_cur_m_sqr
    UpdateUnits "m^2", Unit_cur_m_sqr.curVal
    PropertyChanged "cur_m_sqr"
End Property
Public Property Get cur_m_sqr_err() As Boolean
    cur_m_sqr_err = Unit_cur_m_sqr.bError
End Property

'=================[ cm^2 ]
Public Property Get cur_cm_sqr() As Double
    cur_cm_sqr = Unit_cur_cm_sqr.curVal
End Property
Public Property Let cur_cm_sqr(ByVal New_cur_cm_sqr As Double)
    Unit_cur_cm_sqr.curVal = New_cur_cm_sqr
    UpdateUnits "cm^2", Unit_cur_cm_sqr.curVal
    PropertyChanged "cur_cm_sqr"
End Property
Public Property Get cur_cm_sqr_err() As Boolean
    cur_cm_sqr_err = Unit_cur_cm_sqr.bError
End Property

'=================[ mm^2 ]
Public Property Get cur_mm_sqr() As Double
    cur_mm_sqr = Unit_cur_mm_sqr.curVal
End Property
Public Property Let cur_mm_sqr(ByVal New_cur_mm_sqr As Double)
    Unit_cur_mm_sqr.curVal = New_cur_mm_sqr
    UpdateUnits "mm^2", Unit_cur_mm_sqr.curVal
    PropertyChanged "cur_mm_sqr"
End Property
Public Property Get cur_mm_sqr_err() As Boolean
    cur_mm_sqr_err = Unit_cur_mm_sqr.bError
End Property

'////////////////////////////////////////////////////////////

'Initialize Properties for User Control
Private Sub UserControl_Initialize()
    'set default values
    Unit_cur_yd_sqr.defVal = 1.1959900462997#
    Unit_cur_ft_sqr.defVal = 10.763910416697#
    Unit_cur_in_sqr.defVal = 1550.0031000062#
    Unit_cur_m_sqr.defVal = 1#
    Unit_cur_cm_sqr.defVal = 10000#
    Unit_cur_mm_sqr.defVal = 1000000#
End Sub

'Initialize Properties for User Control Creation
Private Sub UserControl_InitProperties()
    m_ctl_ListIndex = m_def_ctl_ListIndex
    m_ctl_sigDigits = m_def_ctl_sigDigits
    m_ctl_tboxWidth = m_def_ctl_tboxWidth
    m_ctl_FormatSci = m_def_ctl_FormatSci
    'set current values to default values'
    Unit_cur_yd_sqr.curVal = Unit_cur_yd_sqr.defVal
    Unit_cur_ft_sqr.curVal = Unit_cur_ft_sqr.defVal
    Unit_cur_in_sqr.curVal = Unit_cur_in_sqr.defVal
    Unit_cur_m_sqr.curVal = Unit_cur_m_sqr.defVal
    Unit_cur_cm_sqr.curVal = Unit_cur_cm_sqr.defVal
    Unit_cur_mm_sqr.curVal = Unit_cur_mm_sqr.defVal
    StartUp
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ctl_ListIndex = PropBag.ReadProperty("ctl_ListIndex", m_def_ctl_ListIndex)
    m_ctl_sigDigits = PropBag.ReadProperty("ctl_sigDigits", m_def_ctl_sigDigits)
    m_ctl_tboxWidth = PropBag.ReadProperty("ctl_tboxWidth", m_def_ctl_tboxWidth)
    m_ctl_FormatSci = PropBag.ReadProperty("ctl_FormatSci", m_def_ctl_FormatSci)
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Text1.Enabled = PropBag.ReadProperty("Enabled", True)
    'set current values to last-saved values
    Unit_cur_yd_sqr.curVal = PropBag.ReadProperty("cur_yd_sqr", Unit_cur_yd_sqr.defVal)
    Unit_cur_ft_sqr.curVal = PropBag.ReadProperty("cur_ft_sqr", Unit_cur_ft_sqr.defVal)
    Unit_cur_in_sqr.curVal = PropBag.ReadProperty("cur_in_sqr", Unit_cur_in_sqr.defVal)
    Unit_cur_m_sqr.curVal = PropBag.ReadProperty("cur_m_sqr", Unit_cur_m_sqr.defVal)
    Unit_cur_cm_sqr.curVal = PropBag.ReadProperty("cur_cm_sqr", Unit_cur_cm_sqr.defVal)
    Unit_cur_mm_sqr.curVal = PropBag.ReadProperty("cur_mm_sqr", Unit_cur_mm_sqr.defVal)
    StartUp
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("ctl_ListIndex", m_ctl_ListIndex, m_def_ctl_ListIndex)
    Call PropBag.WriteProperty("ctl_sigDigits", m_ctl_sigDigits, m_def_ctl_sigDigits)
    Call PropBag.WriteProperty("ctl_tboxWidth", m_ctl_tboxWidth, m_def_ctl_tboxWidth)
    Call PropBag.WriteProperty("ctl_FormatSci", m_ctl_FormatSci, m_def_ctl_FormatSci)
    Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
    'save current values to storage
    Call PropBag.WriteProperty("cur_yd_sqr", Unit_cur_yd_sqr.curVal, Unit_cur_yd_sqr.defVal)
    Call PropBag.WriteProperty("cur_ft_sqr", Unit_cur_ft_sqr.curVal, Unit_cur_ft_sqr.defVal)
    Call PropBag.WriteProperty("cur_in_sqr", Unit_cur_in_sqr.curVal, Unit_cur_in_sqr.defVal)
    Call PropBag.WriteProperty("cur_m_sqr", Unit_cur_m_sqr.curVal, Unit_cur_m_sqr.defVal)
    Call PropBag.WriteProperty("cur_cm_sqr", Unit_cur_cm_sqr.curVal, Unit_cur_cm_sqr.defVal)
    Call PropBag.WriteProperty("cur_mm_sqr", Unit_cur_mm_sqr.curVal, Unit_cur_mm_sqr.defVal)
End Sub

Private Sub StartUp()
    Set UserControl.Font = Combo1.Font
    Set Text1.Font = Combo1.Font
    Set Label1.Font = Combo1.Font
    Text1.Width = ScaleX(m_ctl_tboxWidth, vbPixels, vbTwips)
    'load the unit collection
    Unit_cur_yd_sqr.uName = "yd^2": UnitCol.Add Unit_cur_yd_sqr, "yd^2"
    Unit_cur_ft_sqr.uName = "ft^2": UnitCol.Add Unit_cur_ft_sqr, "ft^2"
    Unit_cur_in_sqr.uName = "in^2": UnitCol.Add Unit_cur_in_sqr, "in^2"
    Unit_cur_m_sqr.uName = "m^2": UnitCol.Add Unit_cur_m_sqr, "m^2"
    Unit_cur_cm_sqr.uName = "cm^2": UnitCol.Add Unit_cur_cm_sqr, "cm^2"
    Unit_cur_mm_sqr.uName = "mm^2": UnitCol.Add Unit_cur_mm_sqr, "mm^2"
    ReLoadComboBox m_ctl_ListIndex
End Sub

Private Sub UserControl_Resize()
    Combo1.Width = UserControl.Width
    Combo1.Move 0, 0, UserControl.Width
    Text1.Left = 0: Text1.Top = 0: Text1.Height = Combo1.Height
    UserControl.Height = Combo1.Height
    UpdateUnits "yd^2", Unit_cur_yd_sqr.curVal
End Sub

Private Sub ReLoadComboBox(ByVal lstNdx As Integer)
    Dim ndx As Integer
    Dim numStr As String
    Dim lbl As String
    Dim strt As Integer
    Dim cboStr As String
    
    Combo1.Clear
    If UnitCol.Count <> 0 Then
        For ndx = 1 To UnitCol.Count
            If UnitCol(ndx).bError = True Then
                numStr = "<range>"
            Else
                numStr = getFormat(SigFigs(UnitCol(ndx).curVal, m_ctl_sigDigits))
            End If
            lbl = UnitCol(ndx).uName
            Combo1.AddItem MakeCboString(numStr, lbl)
        Next
        Combo1.ListIndex = lstNdx
    End If
End Sub

Private Function MakeCboString(numStr As String, lbl As String) As String
    Dim tBoxWidth As Long
    Dim boolFlag As Boolean
    Dim leftSpaces As Long
    Dim rightSpaces As Long
    
    tBoxWidth = Text1.Width
    'Use Label1's autosizing capability to build an approximately
    'centered number string with a slightly larger width than
    'Text1.Width. Works fairly well with any selected Font.
    leftSpaces = 0: rightSpaces = 2
    Label1.Caption = ""
    Do While True
    'add spaces, one at a time, to either side of num-string
        Label1.Caption = (String(leftSpaces, " ") & numStr & String(rightSpaces, " "))
        If Label1.Width > tBoxWidth Then Exit Do
        boolFlag = Not boolFlag
        If boolFlag Then
            rightSpaces = rightSpaces + 1
        Else
            leftSpaces = leftSpaces + 1
        End If
    Loop
    MakeCboString = String(leftSpaces, " ") & numStr & String(rightSpaces, " ") & lbl
End Function

Private Sub Combo1_Click()
    'Update text display
    If UnitCol(Combo1.ListIndex + 1).bError = True Then
        Text1.Text = "<range>"
    Else
        Text1.Text = getFormat(SigFigs(UnitCol(Combo1.ListIndex + 1).curVal, m_ctl_sigDigits))
    End If
End Sub

Private Sub UpdateUnits(unitName As String, newVal As Double)
    Dim divisor As Double
    Dim ndx As Integer
    Dim tmpDbl As Double
    
    divisor = UnitCol(unitName).defVal
    On Error GoTo numError
    For ndx = 1 To UnitCol.Count
        UnitCol(ndx).bError = False
        tmpDbl = CDbl(UnitCol(ndx).defVal) / divisor * newVal
        If tmpDbl > -1E-308 And tmpDbl < 0 Then
            tmpDbl = 0
        End If
        If tmpDbl < 1E-308 And tmpDbl > 0 Then
            tmpDbl = 0
        End If
        If tmpDbl > 1E+308 Then
            tmpDbl = 1E+308
            UnitCol(ndx).bError = True
        End If
        If tmpDbl < -1E+308 Then
            tmpDbl = -1E+308
            UnitCol(ndx).bError = True
        End If
        UnitCol(ndx).curVal = tmpDbl
    Next
    ReLoadComboBox Combo1.ListIndex
Exit Sub
numError:
    'Debug.Print Err.Description
    UnitCol(ndx).bError = True
    If Sgn(tmpDbl) = 1 Then
        tmpDbl = 1E+308
    Else
        tmpDbl = -1E+308
    End If
    Resume Next
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
    ReLoadComboBox Combo1.ListIndex
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
