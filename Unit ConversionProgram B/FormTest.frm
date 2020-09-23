VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form FormTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Engineering Unit Converter       [Beta-Development Program]"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   Icon            =   "FormTest.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkOnTop 
      Caption         =   "On Top"
      Height          =   255
      Left            =   4920
      TabIndex        =   8
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton optSciFormat 
      Caption         =   "Scientific Format"
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton optEngFormat 
      Caption         =   "Engineering Format"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.OptionButton optNoFormat 
      Caption         =   "No Format"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.CommandButton cmdScriptHelp 
      Caption         =   "Math Input Help ..."
      Height          =   315
      Left            =   6960
      TabIndex        =   2
      Top             =   90
      Width           =   1575
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Significant Figures"
      Top             =   90
      Width           =   735
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   2040
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3555
      Left            =   2400
      TabIndex        =   7
      Top             =   480
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   6271
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Unit Name                    Toggle Sort"
         Object.Width           =   6235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   4048
      EndProperty
   End
End
Attribute VB_Name = "FormTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Simple unit conversion program

Option Explicit

Private Declare Function SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, ByVal hWndInsertAfter As Long, _
    ByVal x As Long, ByVal y As Long, ByVal cx As Long, _
    ByVal cy As Long, ByVal wFlags As Long) As Long

Private SigDigits As Integer

Private Type UnitElement
    nameString      As String
    defaultValue    As Double
    currentValue    As Double
    numError        As Boolean
End Type
'local array to hold unit data
Private Units() As UnitElement
'IndexOf: key -> nameString, value -> Units()index.
Private IndexOf As Collection

'Constants
Private Const HWND_BOTTOM = 1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1

Private Sub chkOnTop_Click()
    Dim x1 As Long
    Dim y1 As Long
    
    x1 = ScaleX(Me.Left, vbTwips, vbPixels)
    y1 = ScaleY(Me.Top, vbTwips, vbPixels)
    If chkOnTop.Value = 1 Then
        'Dialog with foreground priority
        SetWindowPos hwnd, HWND_TOPMOST, x1, y1, 0, 0, SWP_NOSIZE
    Else
        'Normal priority
        SetWindowPos hwnd, HWND_NOTOPMOST, x1, y1, 0, 0, SWP_NOSIZE
    End If
End Sub


Private Sub Form_Load()
    Dim ndx As Integer
    Dim INIString As String
    
    'subclass ListView1 window procedure:
    'to detect scroll_events'
    OldListView1WindowProc = SetWindowLong( _
        ListView1.hwnd, GWL_WNDPROC, _
        AddressOf NewListView1WindowProc)
        
    'load List1 with unit category choices
    ndx = 0
    Do While True
        ndx = ndx + 1
        INIString = GetINIString("Categories", CStr(ndx))
        If INIString = "None" Then
            Exit Do
        Else
            List1.AddItem INIString
        End If
    Loop
    'Set listbox index
    List1.ListIndex = 36 ' move list down more than needed
    List1.ListIndex = 26 ' set to <- Length
    'load significant digits Combobox
    loadSigFigComboBox
    'load combobox and listview lists
    loadListView
    'set up the script control
    SetupScriptControl ScriptControl1
End Sub

'select from unit category choices
Private Sub List1_Click()
    'loads data from the selected unit
    'category into the listview box
    loadListView
End Sub

Private Sub loadSigFigComboBox()
    Combo2.AddItem "1" 'index 0
    Combo2.AddItem "2" 'index 1
    Combo2.AddItem "3" 'index 2
    Combo2.AddItem "4" 'index 3
    Combo2.AddItem "5" 'index 4
    Combo2.AddItem "6" 'index 5
    Combo2.AddItem "7" 'index 6
    Combo2.AddItem "8" 'index 7
    Combo2.AddItem "9" 'index 8
    Combo2.AddItem "10" 'index 9
    Combo2.AddItem "11" 'index 10
    Combo2.AddItem "12" 'index 11
    Combo2.AddItem "13" 'index 12
    Combo2.AddItem "14" 'index 13
    'set to indicate 14 sig digits
    SigDigits = 14 'default
    Combo2.ListIndex = 13
End Sub

Private Sub Combo2_Click()
    'set significant digits
    Select Case Combo2.ListIndex
        Case Is = 0 '1
            SigDigits = 1
        Case Is = 1 '2
            SigDigits = 2
        Case Is = 2 '3
            SigDigits = 3
        Case Is = 3 '4
            SigDigits = 4
        Case Is = 4 '5
            SigDigits = 5
        Case Is = 5 '6
            SigDigits = 6
        Case Is = 6 '7
            SigDigits = 7
        Case Is = 7 '8
            SigDigits = 8
        Case Is = 8 '9
            SigDigits = 9
        Case Is = 9 '10
            SigDigits = 10
        Case Is = 10 '11
            SigDigits = 11
        Case Is = 11 '12
            SigDigits = 12
        Case Is = 12 '13
            SigDigits = 13
        Case Is = 13 '14
            SigDigits = 14
    End Select
    updateListViewValues
End Sub

Private Sub loadListView()
'load the listview box with data from the
'INI file selected unit heading
    Dim ndx As Integer
    Dim heading As String
    Dim unitName As String
    Dim unitVal As Double
    Dim tmpStr As String
    Dim locn As Integer
    
    'clear the listView
    ListView1.ListItems.Clear
    'set heading to name of selected category in List1
    heading = List1.List(List1.ListIndex)
    'clear the Units() array
    ReDim Units(0)
    'start a new collection
    Set IndexOf = New Collection '<- module level declaration
    'load the listView:
    ndx = 0
    Do While True
        ndx = ndx + 1
        tmpStr = GetINIString(heading, CStr(ndx))
        If tmpStr = "None" Then
            Exit Do
        Else
            'parse unitName & unitVal
            locn = InStr(tmpStr, ";")
            unitName = Left(tmpStr, locn - 1)
            unitVal = CDbl(Right(tmpStr, Len(tmpStr) - locn))
            'add to Units() array
            ReDim Preserve Units(ndx - 1)
            Units(ndx - 1).nameString = unitName
            Units(ndx - 1).defaultValue = unitVal
            Units(ndx - 1).currentValue = unitVal
            Units(ndx - 1).numError = False
            'IndexOf: key -> unitName, value -> Units()index.
            IndexOf.Add ndx - 1, unitName ' <- late-bound
            'add to listView
            ListView1.ListItems.Add , unitName, unitName
            ListView1.ListItems(unitName).SubItems(1) = getFormat(SigFigs_Str(unitVal, SigDigits))
        End If
    Loop
    'Place & Show Text1
    ShowTextBox
End Sub

'A unit value has been changed.
Private Sub UpdateUnits(changedIndex As Integer, newVal As Double)
    Dim divisor As Double
    Dim ndx As Integer
    
    On Error GoTo UpdateUnitsErrHandler
    'check for Temperature:
    If List1.List(List1.ListIndex) = "Temperature" Then
        For ndx = 0 To 3
            Units(ndx).numError = False
        Next
        Select Case changedIndex
            Case Is = 0 '°K
                ndx = 0  '<- for .numError
                Units(IndexOf("°Kelvin")).currentValue = newVal
                If Units(IndexOf("°Kelvin")).currentValue < 0 Then Units(IndexOf("°Kelvin")).currentValue = 0
                ndx = 1  '<- for .numError
                Units(IndexOf("°Celsius")).currentValue = Units(IndexOf("°Kelvin")).currentValue - 273.15
                ndx = 2  '<- for .numError
                Units(IndexOf("°Fahrenheit")).currentValue = Units(IndexOf("°Celsius")).currentValue * 9 / 5 + 32
                ndx = 3  '<- for .numError
                Units(IndexOf("°Rankine")).currentValue = Units(IndexOf("°Kelvin")).currentValue * 9 / 5
            Case Is = 1 '°C
                ndx = 1  '<- for .numError
                Units(IndexOf("°Celsius")).currentValue = newVal
                If Units(IndexOf("°Celsius")).currentValue < -273.15 Then Units(IndexOf("°Celsius")).currentValue = -273.15
                ndx = 0  '<- for .numError
                Units(IndexOf("°Kelvin")).currentValue = Units(IndexOf("°Celsius")).currentValue + 273.15
                ndx = 3  '<- for .numError
                Units(IndexOf("°Rankine")).currentValue = Units(IndexOf("°Kelvin")).currentValue * 9 / 5
                ndx = 2  '<- for .numError
                Units(IndexOf("°Fahrenheit")).currentValue = Units(IndexOf("°Rankine")).currentValue - 459.67
            Case Is = 2 '°F
                ndx = 2  '<- for .numError
                Units(IndexOf("°Fahrenheit")).currentValue = newVal
                If Units(IndexOf("°Fahrenheit")).currentValue < -459.67 Then Units(IndexOf("°Fahrenheit")).currentValue = -459.67
                ndx = 1  '<- for .numError
                Units(IndexOf("°Celsius")).currentValue = (Units(IndexOf("°Fahrenheit")).currentValue - 32) * 5 / 9
                ndx = 3  '<- for .numError
                Units(IndexOf("°Rankine")).currentValue = Units(IndexOf("°Fahrenheit")).currentValue + 459.67
                ndx = 0  '<- for .numError
                Units(IndexOf("°Kelvin")).currentValue = Units(IndexOf("°Rankine")).currentValue * 5 / 9
            Case Is = 3 '°R
                ndx = 3  '<- for .numError
                Units(IndexOf("°Rankine")).currentValue = newVal
                If Units(IndexOf("°Rankine")).currentValue < 0 Then Units(IndexOf("°Rankine")).currentValue = 0
                ndx = 2  '<- for .numError
                Units(IndexOf("°Fahrenheit")).currentValue = Units(IndexOf("°Rankine")).currentValue - 459.67
                ndx = 1  '<- for .numError
                Units(IndexOf("°Celsius")).currentValue = (Units(IndexOf("°Fahrenheit")).currentValue - 32) * 5 / 9
                ndx = 0  '<- for .numError
                Units(IndexOf("°Kelvin")).currentValue = Units(IndexOf("°Rankine")).currentValue * 5 / 9
        End Select
    Else
        'Update all other unit values proportionally
        'using the constant ratio principal:
        '
        '     current_foot      default_foot
        '    --------------- = --------------- = CONSTANT_RATIO
        '     current_meter     default_meter
        '
        'No Units().defaultValue should ever = zero,
        'or division by zero error will occur!
        divisor = Units(changedIndex).defaultValue
        'Set the .currentValue of each unit:
        For ndx = 0 To UBound(Units)
            If divisor <> 0 Then
                Units(ndx).numError = False
                Units(ndx).currentValue = (Units(ndx).defaultValue / divisor) * newVal
            Else
                MsgBox "Division by zero in UpdateUnits()"
                Stop
            End If
        Next
    End If
Exit Sub
UpdateUnitsErrHandler:
    Units(ndx).numError = True
    Resume Next
End Sub

Private Sub updateListViewValues()
'Update the listview values
    Dim ndx As Integer
    For ndx = 0 To UBound(Units())
        If Units(ndx).numError = True Then
            ListView1.ListItems(Units(ndx).nameString).SubItems(1) = "out of range"
        Else
            ListView1.ListItems(Units(ndx).nameString).SubItems(1) = getFormat(SigFigs_Str(Units(ndx).currentValue, SigDigits))
        End If
    Next
    'Place & Show Text1
    ShowTextBox
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    HideTextBox
End Sub

'sort/unsort Unit Names
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim selText As String
    Dim selVal As Double
    Dim selItem As ListItem
    
    'Unit Names' Column Header Button Click:
    If ColumnHeader.Index = 1 Then
        'toggle sort state
        ListView1.Sorted = Not ListView1.Sorted
        'save selection data:
        selText = ListView1.SelectedItem.Text
        selVal = ListView1.SelectedItem.SubItems(1)
        Set selItem = ListView1.ListItems(selText)
        'reload listview from INI file
        loadListView
        'restore selection data:
        Set ListView1.SelectedItem = ListView1.ListItems(selText)
        UpdateUnits IndexOf(selText), selVal
        updateListViewValues
    End If
End Sub

Public Sub HideTextBox()
    Text1.Visible = False
    Text1.BackColor = vbCyan
End Sub

Private Sub ShowTextBox()
    'Place and show Text1
    Dim hei As Single
    Dim offsetX As Long
    Dim offsetY As Long
    Dim ndx As Integer
    
    hei = ListView1.ListItems.Item(1).Height
    Text1.Height = hei - 15
    'Debug.Print hei, Text1.Height
    offsetY = ListView1.Top + 300
    offsetX = ListView1.Left + ListView1.ColumnHeaders(1).Width
    'Text1.Top = offset + 13 * hei
    'Debug.Print ListView1.GetFirstVisible.Index
    'Debug.Print ListView1.SelectedItem.Index
    ndx = ListView1.SelectedItem.Index - ListView1.GetFirstVisible.Index
    Text1.Width = ListView1.ColumnHeaders(2).Width
    Text1.Left = offsetX + 30
    Text1.Top = offsetY + ndx * hei
    'Debug.Print ListView1.SelectedItem.SubItems(1)
    Text1.Text = ListView1.SelectedItem.SubItems(1)
    Text1.Visible = True
    
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ShowTextBox
End Sub

Private Sub optEngFormat_Click()
    updateListViewValues
End Sub
Private Sub optNoFormat_Click()
    updateListViewValues
End Sub
Private Sub optSciFormat_Click()
    updateListViewValues
End Sub

'======[ TextBox Script Control Usage ]==================
'Use a Script Control to evaluate any math operations
'entered into the textbox by the user and place the
'result back into the textbox for processing
Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim expr As String
    Dim result As Double
    Dim msgTitle As String
    Dim msgPrompt As String
    Dim ndx As Integer
    
    Text1.BackColor = vbYellow ' for editing feedback
    If KeyAscii = 13 Then '[Enter] key was pressed
        'Get any mathmatical expression in the textbox.
        'In a real program it will probably be read from a
        'textbox control: expr = "Log(10) + (pi * 10) + 10"
        expr = Text1.Text
        KeyAscii = 0 'suppress 'beep' sound
        On Error GoTo scError_1
        'Calclate result of the expression:
        result = ScriptControl1.Eval(expr)
        'no error occured so...
        'process the result, update all unit values
        UpdateUnits IndexOf(ListView1.SelectedItem.Text), result
        'back to normal color
        Text1.BackColor = vbCyan
        'update the listview values to match
        updateListViewValues
    End If
    Exit Sub
scError_1: 'A script control error occured
    'Inform user of the error:
    msgTitle = ScriptControl1.Error.Source & ": " & _
    Str$(ScriptControl1.Error.Number)
    msgPrompt = ScriptControl1.Error.Description
    'Show user the error information:
    MsgBox msgPrompt, , msgTitle
End Sub

Private Function getFormat(numStr As String) As String
    Dim Number As Double
    
    If numStr = "out of range" Then
        getFormat = "out of range"
        Exit Function
    Else
        Number = CDbl(numStr)
    End If
    If optSciFormat.Value = True Then
        'Engineering format
        getFormat = FormatSci(Number, 14)
        Exit Function
    End If
    If optEngFormat.Value = True Then
        'Engineering format
        getFormat = FormatEng(Number, 14)
        Exit Function
    End If
    'No Formating
    getFormat = Str(Number)
End Function

Private Sub SetupScriptControl(sc As ScriptControl)
    Dim myCode As String
    'Initialize a language engine for the script control:
    sc.Language = "VBScript"
    'Add any variables that you want the script control to
    'know about:
    'set pi as a known variable
    sc.ExecuteStatement "pi = 3.1415926535898"
    'Add any additional functions that you want the script
    'control to know about...
    '/// log10(x) /// log to the base 10
    myCode = _
    "Function log10(var)" + vbCrLf + _
    " log10 = log(var) / log(10)" + vbCrLf + _
    "End Function"
    sc.AddCode myCode
    '/// ln(x) /// natural log, to base e
    myCode = _
    "Function ln(var)" + vbCrLf + _
    " ln = log(var)" + vbCrLf + _
    "End Function"
    sc.AddCode myCode
    '/// d2r(x) /// degrees to radians
    myCode = _
    "Function d2r(var)" + vbCrLf + _
    " d2r = pi/180*(var)" + vbCrLf + _
    "End Function"
    sc.AddCode myCode
    '/// r2d(x) /// radians to degrees
    myCode = _
    "Function r2d(var)" + vbCrLf + _
    " r2d = 180/pi*(var)" + vbCrLf + _
    "End Function"
    sc.AddCode myCode
    'NOTE: These functions can be evaluated directly from
    'within a textbox because they evaluate to a number.
End Sub

Private Sub cmdScriptHelp_Click()
    ShowScriptControlMathInputHelp
End Sub
'Pop up a message box informing user of available math
'scripting inputs for Text1:
Private Sub ShowScriptControlMathInputHelp()
    Dim msg As String
    Dim msgTitle As String
    
    msgTitle = "   Math Input Functions:"
    msg = "Standard Functions:" & vbTab & vbTab & "Additional Functions:" & vbCrLf
    msg = msg & "Trigonometric: (radian mode)" & vbTab & vbTab & "log10(x) - base 10 logarithm" & vbCrLf
    msg = msg & "atn(x) - inverse tangent" & vbTab & vbTab & "ln(x) - natural logarithm" & vbCrLf
    msg = msg & "sin(x) - sine" & vbTab & vbTab & vbTab & "d2r(x) - degrees to radians" & vbCrLf
    msg = msg & "cos(x) - cosine" & vbTab & vbTab & vbTab & "r2d(x) - radians to degrees" & vbCrLf
    msg = msg & "tan(x) - tangent" & vbTab & vbTab & vbTab & "pi - 3.1415926535898" & vbCrLf & vbCrLf
    msg = msg & "Standard:" & vbTab & vbTab & vbTab & "Standard Operators:" & vbCrLf
    msg = msg & "exp(x) - exponential" & vbTab & vbTab & "(+) - addition" & vbTab & "(-) - subtraction" & vbCrLf
    msg = msg & "log(x) - natural logarithm" & vbTab & vbTab & "(*) - multiplication" & vbTab & "(/) - division" & vbCrLf
    msg = msg & "sqr(x) - square root" & vbTab & vbTab & vbTab & "(^) - exponetation" & vbTab & "( ) - parentheses" & vbCrLf & vbCrLf
    msg = msg & "Example Input:    (pi*4^2)+2" & vbTab & vbTab & "Note:  X^(1/3) = cube root of X" & vbCrLf
    msg = msg & "Resolves To:    52.265482457437"
    'Show help message:
    MsgBox msg, , msgTitle
End Sub

'///////////////////////////////////////////////////////////////
'===============================================================
'Return 'dblNumber' rounded to 'intSF' significant figures
'===============================================================
Private Function SigFigs(dblNumber As Double, intSF As Integer) As Double
'Only works properly for doubles in the range: (+/-)1.79769313486231E(+/-)308
    Dim negFlag As Integer, tmpDbl As Double, factor As Double
    Dim dblA As Double, dblB As Double, outNum As Double
    
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
'Wrapper function for calling SigFigs() within its valid range:
'Returns a 'String' which indicates validity...
Private Function SigFigs_Str(dblNumber As Double, intSF As Integer) As String
'Microsoft states that the range of the Double number type in VB6 is:
'-1.79769313486232E308 to -4.94065645841247E-324 for negative values,
'4.94065645841247E-324 to 1.79769313486232E308 for positive values.
'
'However, try to enter: myDbl = 9.9E-320, into any code module; it
'gets converted into: myDbl = 9.9000874113669E-320 right in front
'of your eyes. And there is 'nothing' you can do about it!
'
'This behavior (bug), when it happens, produces an error in the
'calculation of the internal 'factor' variable in the SigFigs()
'function. This error causes the function not to round properly.
'
'I have experimented with this behavior and discovered that it does
'not occur with negative exponents less than 308.
'Specifically, the valid lower limit is: (+/-)1.79769313486231E(-)308
'
'          Abbreviated [valid] SigFigs() Number Line Range:
'
' ---[-1E+308]-------[-1E-308]---[0]---[+1E-308]-------[+1E+308]---
'       [^-----valid-----^]    [valid]    [^-----valid-----^]
'
'Safe-Usage Example:
'Private Sub MySub()
'   On Error GoTo MyErrHandler
'   'The Dbl_Expression evaluation itself could cause an error,
'   'for example, if Dbl_Expression = 23E150 * 15E200.
'   'Also, try Dbl_Expression = -3.7E-321
'   myNumString = SigFigs_Str(Dbl_Expression, 5)
'   If myNumString = "out of range" Then
'      'do some out-of-range error stuff here:
'      Debug.Print "out of range"
'   Else
'      'It WILL have been rounded OK:
'      my_5_SigFigs_Dbl = CDbl(myNumString)
'      Debug.Print my_5_SigFigs_Dbl
'   End If
'Exit Sub
'MyErrHandler:
'   myNumString = "out of range"
'   Resume Next
'End Sub
'
    Dim SFnumInRange As Boolean
    
    SFnumInRange = False
    If dblNumber <> 0 Then
        'negative values
        If dblNumber >= -1E+308 And dblNumber <= -1E-308 Then
            SFnumInRange = True
        End If
        'positive values
        If dblNumber >= 1E-308 And dblNumber <= 1E+308 Then
            SFnumInRange = True
        End If
    Else
        'dblNumber = 0
        SFnumInRange = True
    End If
    'Return a string:
    If SFnumInRange Then
        'SigFigs() WILL round dblNumber properly:
        'Return the SigFigs() string value
        SigFigs_Str = CStr(SigFigs(dblNumber, intSF))
    Else
        'SigFigs() could possibly NOT round dblNumber properly:
        'Return: "out of range"
        SigFigs_Str = "out of range"
    End If
End Function
'///////////////////////////////////////////////////////////////

Private Function FormatEng(Number As Double, Optional _
                          DecimalPlaces As Long = 1) As String
    Dim Exponent As Long
    Dim Parts() As String
    
    If Abs(Number) < 1000 And Abs(Number) >= 1 Then
        FormatEng = Format(Number, "0.0#############")
        Exit Function
    End If
    If Number = 0 Then
        FormatEng = Format(Number, "0.0")
        Exit Function
    End If
    Parts = Split(Format(Number, "0.0#############E+0"), "E")
    Exponent = 3 * Int(Parts(1) / 3)
    FormatEng = Format(Parts(0) * 10 ^ (Parts(1) - Exponent), _
                               "0.0" & String(DecimalPlaces, "#")) & _
                               "E" & Format(Exponent, "+0;-0")
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


