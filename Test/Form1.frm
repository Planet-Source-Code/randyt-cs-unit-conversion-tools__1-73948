VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "User-Control Test Platform:"
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin ProjectTest2.ColorButton ColorButton1 
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   4440
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Register ScriptControl with Unit-Controls"
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   240
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ColorButton1_Click()
    'LengthControl1.RegisterScriptControl ScriptControl1
    'LengthControl2.RegisterScriptControl ScriptControl1
    
End Sub

Private Sub Form_Load()
    SetupScriptControl ScriptControl1
End Sub

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
