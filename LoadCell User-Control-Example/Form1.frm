VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deflecting Beam Load Cell Designer"
   ClientHeight    =   7575
   ClientLeft      =   420
   ClientTop       =   1095
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7575
   ScaleWidth      =   11400
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   360
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Plot"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   7560
      TabIndex        =   10
      Top             =   5760
      Width           =   3735
      Begin VB.CommandButton cmdForceVsDeflection 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Force vs. Deflection: Two beams"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Form1.frx":030A
         MousePointer    =   99  'Custom
         TabIndex        =   11
         Top             =   1080
         Width           =   3495
      End
      Begin VB.CommandButton cmdStressVsLength 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Stress (psi) vs. Length (inches)"
         Height          =   495
         Left            =   120
         MouseIcon       =   "Form1.frx":0614
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Calculate: for TWO beams"
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   7560
      TabIndex        =   9
      Top             =   2040
      Width           =   3735
      Begin LoadCell.ColorButton cmdLength 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "Solve for Length"
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":091E
         MousePointer    =   99
         Caption         =   "Length:     L"
      End
      Begin LoadCell.ColorButton cmdWidth 
         Height          =   375
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Solve for Width"
         Top             =   2436
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":0C38
         MousePointer    =   99
         Caption         =   "Width:      W"
      End
      Begin LoadCell.ColorButton cmdThickness 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Solve for Thickness"
         Top             =   1872
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":0F52
         MousePointer    =   99
         Caption         =   "Thickness: T"
      End
      Begin LoadCell.ColorButton cmdForce 
         Height          =   375
         Left            =   120
         TabIndex        =   21
         ToolTipText     =   "Solve for Force"
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":126C
         MousePointer    =   99
         Caption         =   " Force :      F"
      End
      Begin LoadCell.ColorButton cmdDeflection 
         Height          =   375
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Solve for deflection"
         Top             =   924
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":1586
         MousePointer    =   99
         Caption         =   "deflection: d"
      End
      Begin LoadCell.LengthControl lenDeflection 
         Height          =   360
         Left            =   1440
         TabIndex        =   14
         Top             =   921
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ctl_sigDigits   =   3
         ctl_FormatSci   =   -1  'True
         BackColor       =   0
         Combo1_Width    =   855
         cur_in          =   1
         cur_ft          =   8.33333333333316E-02
         cur_yd          =   2.77777777777772E-02
         cur_mm          =   25.3999999999997
         cur_cm          =   2.53999999999997
         cur_m           =   2.53999999999997E-02
      End
      Begin LoadCell.LengthControl lenThickness 
         Height          =   360
         Left            =   1440
         TabIndex        =   15
         Top             =   1878
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ctl_sigDigits   =   3
         BackColor       =   0
         Combo1_Width    =   840
         cur_in          =   0.1
         cur_ft          =   8.33333333333316E-03
         cur_yd          =   2.77777777777772E-03
         cur_mm          =   2.53999999999997
         cur_cm          =   0.253999999999997
         cur_m           =   2.53999999999997E-03
      End
      Begin LoadCell.LengthControl lenWidth 
         Height          =   360
         Left            =   1440
         TabIndex        =   16
         Top             =   2439
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ctl_sigDigits   =   3
         BackColor       =   0
         Combo1_Width    =   840
         cur_in          =   1
         cur_ft          =   8.33333333333316E-02
         cur_yd          =   2.77777777777772E-02
         cur_mm          =   25.3999999999997
         cur_cm          =   2.53999999999997
         cur_m           =   2.53999999999997E-02
      End
      Begin LoadCell.LengthControl lenLength 
         Height          =   360
         Left            =   1440
         TabIndex        =   17
         Top             =   3000
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ctl_sigDigits   =   3
         BackColor       =   0
         Combo1_Width    =   840
         cur_in          =   3
         cur_ft          =   0.249999999999995
         cur_yd          =   8.33333333333316E-02
         cur_mm          =   76.199999999999
         cur_cm          =   7.6199999999999
         cur_m           =   0.076199999999999
      End
      Begin LoadCell.ForceControl frcForce 
         Height          =   360
         Left            =   1440
         TabIndex        =   18
         Top             =   360
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   3
         ctl_sigDigits   =   3
         BackColor       =   0
         Combo1_Width    =   855
         cur_N           =   4.44822161525477
         cur_gf          =   453.592369999403
         cur_kgf         =   0.453592369999403
         cur_lbf         =   1
         cur_ozf         =   16
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "T => Thickness of ONE beam:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   1485
         Width           =   2580
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   7290
      Begin VB.PictureBox picFunctionPlot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   3015
         Left            =   0
         ScaleHeight     =   2985
         ScaleWidth      =   7260
         TabIndex        =   8
         Top             =   360
         Width           =   7290
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Beam Support / Loading Diagram"
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7290
      Begin VB.PictureBox picBeamSupportAndLoading 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H0000FFFF&
         ForeColor       =   &H80000008&
         Height          =   3420
         Left            =   0
         Picture         =   "Form1.frx":18A0
         ScaleHeight     =   3390
         ScaleWidth      =   7260
         TabIndex        =   1
         Top             =   375
         Width           =   7290
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Beam Material"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   7560
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin LoadCell.ColorButton cmdElasticModulus 
         Height          =   375
         Left            =   120
         TabIndex        =   25
         ToolTipText     =   "Solve for Elastic Modulus"
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Form1.frx":1C47A
         MousePointer    =   99
         Caption         =   "Elastic Modulus"
      End
      Begin VB.TextBox txtYieldStrength 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Text            =   "30000"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox cmbMaterials 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3495
      End
      Begin LoadCell.StressControl strssYoungsModulus 
         Height          =   360
         Left            =   1680
         TabIndex        =   19
         Top             =   840
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   635
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ctl_sigDigits   =   3
         ctl_FormatSci   =   -1  'True
         BackColor       =   0
         Combo1_Width    =   840
         cur_psi         =   30000000
         cur_kpa         =   206842718.795353
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "psi"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3120
         TabIndex        =   4
         Top             =   1350
         Width           =   330
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   " Yield  Strength:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   1350
         Width           =   1875
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Graph As New GraphTool

Private Sub Form_Activate()
    'Set focus to force vs. deflection button:
    cmdForceVsDeflection.SetFocus
End Sub

Private Sub Form_Load()
    Dim d#
    'Set the material properties:
    SetProperties  'Located in Metals.bas
    'Load the materials into the combo box:
    loadMaterials  'In this module
    'calculate and display the initial deflection:
    cmdDeflection_Click
    'Show force vs deflection graph:
    cmdForceVsDeflection_Click
    'Setup ScriptControl1
    SetupScriptControl ScriptControl1
    'Register ScriptControl1 with conversion controls:
    RegScriptControl ScriptControl1
End Sub

'Load material selections into the combobox:
Private Sub loadMaterials()
    Dim n
    
    'Load the materials array "name" items into the combo box:
    For n = LBound(matArray) To UBound(matArray)
        cmbMaterials.AddItem matArray(n).name
    Next n
    'load the first material into the "Beam Material" data box displays:
    cmbMaterials.ListIndex = 0   'loads first material name
    strssYoungsModulus.cur_psi = Str$(matArray(0).E)  'loads first material Elastic Modulus
    txtYieldStrength.Text = Str$(matArray(0).ys) 'loads first material Yield Strength
End Sub

'Setup the ScriptControl:
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

'Register the ScriptControl with unit-conversion controls:
Private Sub RegScriptControl(sc As ScriptControl)
    strssYoungsModulus.RegisterScriptControl sc
    frcForce.RegisterScriptControl sc
    lenDeflection.RegisterScriptControl sc
    lenThickness.RegisterScriptControl sc
    lenWidth.RegisterScriptControl sc
    lenLength.RegisterScriptControl sc
End Sub

'////////////////////[ Material Selection ]//////////////////
Private Sub cmbMaterials_Click()
    strssYoungsModulus.cur_psi = Str$(matArray(cmbMaterials.ListIndex).E)
    txtYieldStrength.Text = Str$(matArray(cmbMaterials.ListIndex).ys)
    cellParameterChanged
End Sub

'////////////////////[ Elastic Modulus ]//////////////////
Private Sub cmdElasticModulus_Click()
    Dim E#

    E# = calculateElasitcModulus()
    strssYoungsModulus.cur_psi = E#
    cellParameterClicked
End Sub
Private Function calculateElasitcModulus() As Single
'For twin beam configuration: Max deflection
    Dim W!, T!, I!, F!, L!, E!, d!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    I! = W! * T! ^ 3 / 12  'Area moment of inertia for beam cross section
    d! = lenDeflection.cur_in
    F! = frcForce.cur_lbf / 2
    L! = lenLength.cur_in
    calculateElasitcModulus = (F! * L! ^ 3) / (12 * d! * I!)   'Youngs Modulus
End Function
Private Sub strssYoungsModulus_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

'////////////////////[ Force ]//////////////////
Private Sub cmdForce_Click()
    Dim F#

    F# = calculateForce()
    frcForce.cur_lbf = F#
    cellParameterClicked
End Sub
Private Function calculateForce() As Single
'For twin beam configuration: Applied force
    Dim W!, T!, I!, F!, L!, E!, d!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    I! = W! * T! ^ 3 / 12  'Area moment of inertia for beam cross section
    d! = lenDeflection.cur_in
    L! = lenLength.cur_in
    E! = strssYoungsModulus.cur_psi
    F! = (d! * 12 * E! * I!) / (L! ^ 3)  'Applied force on one beam
    F! = F! * 2                    'Force on two beams
    calculateForce = F!
End Function
Private Sub frcForce_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

'////////////////////[ Deflection ]//////////////////
Private Sub cmdDeflection_Click()
    Dim d#

    d# = calculateDeflection()
    lenDeflection.cur_in = d#
    cellParameterClicked
End Sub
Private Function calculateDeflection() As Single
'For twin beam configuration: Max deflection
    Dim W!, T!, I!, F!, L!, E!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    I! = W! * T! ^ 3 / 12  'Area moment of inertia for beam cross section
    F! = frcForce.cur_lbf / 2
    L! = lenLength.cur_in
    E! = strssYoungsModulus.cur_psi
    calculateDeflection = (F! * L! ^ 3) / (12 * E! * I!)   'Max deflection
End Function
Private Sub lenDeflection_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

'////////////////////[ Thickness ]//////////////////
Private Sub cmdThickness_Click()
    Dim T#

    T# = calculateThickness()
    lenThickness.cur_in = T#
    cellParameterClicked
End Sub
Private Function calculateThickness() As Single
'For twin beam configuration: Thickness of one beam
    Dim W!, T!, F!, L!, E!, d!
    
    W! = lenWidth.cur_in
    d! = lenDeflection.cur_in
    L! = lenLength.cur_in
    E! = strssYoungsModulus.cur_psi
    F! = frcForce.cur_lbf / 2
    T! = ((F! * L! ^ 3) / (d! * E! * W!)) ^ (1 / 3)
    calculateThickness = T!
End Function
Private Sub lenThickness_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

'////////////////////[ Width ]//////////////////
Private Sub cmdWidth_Click()
    Dim W#

    W# = calculateWidth()
    lenWidth.cur_in = W#
    cellParameterClicked
End Sub
Private Function calculateWidth() As Single
'For twin beam configuration: Width of both beams
    Dim W!, T!, F!, L!, E!, d!
    
    d! = lenDeflection.cur_in
    L! = lenLength.cur_in
    E! = strssYoungsModulus.cur_psi
    F! = frcForce.cur_lbf / 2
    T! = lenThickness.cur_in
    W! = (F! * L! ^ 3) / (d! * E! * T! ^ 3)
    calculateWidth = W!
End Function
Private Sub lenWidth_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

'////////////////////[ Length ]//////////////////
Private Sub cmdLength_Click()
    Dim L#

    L# = calculateLength()
    lenLength.cur_in = L#
    cellParameterClicked
End Sub
Private Function calculateLength() As Single
'For twin beam configuration: Width of both beams
    Dim W!, T!, F!, L!, E!, d!, I!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    I! = W! * T! ^ 3 / 12  'Area moment of inertia for beam cross section
    d! = lenDeflection.cur_in
    F! = frcForce.cur_lbf / 2
    E! = strssYoungsModulus.cur_psi
    L! = ((d! * 12 * E! * I!) / (F!)) ^ (1 / 3)
    calculateLength = L!
End Function
Private Sub lenLength_TboxKeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'Enter Key
        cellParameterChanged
    End If
End Sub

Private Sub cellParameterClicked()
'A parameter-related command button has been clicked:
'Change all backColors back to 'normal' vbButtonFace:

    cmdElasticModulus.BackColor = vbButtonFace
    cmdDeflection.BackColor = vbButtonFace
    cmdForce.BackColor = vbButtonFace
    cmdThickness.BackColor = vbButtonFace
    cmdWidth.BackColor = vbButtonFace
    cmdLength.BackColor = vbButtonFace
    'Enable plot buttons:
    cmdStressVsLength.Enabled = True
    cmdForceVsDeflection.Enabled = True
End Sub

Private Sub cellParameterChanged()
'An entry has been made into one of the parameter-related
'Unit Controls. Change all button backColors to yellow:
    
    cmdElasticModulus.BackColor = vbYellow
    cmdDeflection.BackColor = vbYellow
    cmdForce.BackColor = vbYellow
    cmdThickness.BackColor = vbYellow
    cmdWidth.BackColor = vbYellow
    cmdLength.BackColor = vbYellow
    'Clear graph, disable plot buttons:
    picFunctionPlot.Cls
    Frame3.Caption = ""
    cmdStressVsLength.Enabled = False
    cmdForceVsDeflection.Enabled = False
End Sub

'////////////////////[ Stress Vs. Length ]//////////////////
Private Sub cmdStressVsLength_Click()
'For twin beam configuration:
    Dim W!, T!, Z!, F!, L!, s!, stp!
    Dim sf!, yMax!, yMin!, xMin!, xMax!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    Z! = (W! * T! ^ 2) / 6   'Section Modulus for beam cross section
    F! = frcForce.cur_lbf / 2
    L! = lenLength.cur_in
    'If you need to erase the background "grid",
    'the following code will do it:
    
    'picFunctionPlot.AutoRedraw = True  'enables you to erase the background grid.
    picFunctionPlot.Cls   'clears the background grid.
    'picFunctionPlot.AutoRedraw = False 'disables you to erase the background grid.
    
    'Get ready to setup the graph box:
    sf! = 1.15
    yMax! = (F! * L!) / (2 * Z!) * sf!
    yMin! = -yMax!
    sf! = L! * 0.18
    xMin! = 0 - sf!
    sf! = L! * 0.08
    xMax! = L! + sf!
    'Set the scales on the graph:
    Graph.SetUp picFunctionPlot, xMin!, yMax!, xMax!, yMin!
    'Set graph caption:
    Frame3.Caption = "Surface STRESS due to bending moment about the W-axis: (pointing out at you)"
    'Draw axis grid lines
    picFunctionPlot.DrawWidth = 1
    Graph.drawGridLinesHorizontal (F! * L! / 2 / Z!) / 3, RGB(224, 224, 224)
    Graph.drawGridLinesVertical L! / 6, RGB(224, 224, 224)
    'plot the graph:
    stp! = L! / 100
    For s! = 0! To L! Step stp!
      picFunctionPlot.Line (s!, -((F! / Z!) * (L! / 2 - s!)))-(s!, 0), vbRed
    Next s!
    'Plot the yield Strength:
    picFunctionPlot.DrawWidth = 2
    picFunctionPlot.Line (0, Val(txtYieldStrength.Text))-(L!, Val(txtYieldStrength.Text)), vbBlack
    picFunctionPlot.Line (0, -Val(txtYieldStrength.Text))-(L!, -Val(txtYieldStrength.Text)), vbBlack
    'Draw the axes:
    picFunctionPlot.DrawWidth = 1
    Graph.drawXaxis vbBlack
    Graph.drawYaxis vbBlack
    'Draw axes ticks:
    picFunctionPlot.DrawWidth = 2
    Graph.drawHorizontalAxisTicks L! / 6, 0, 1, vbBlack
    Graph.drawVerticalAxisTicks (F! * L! / 2 / Z!) / 3, 0, 0.7, vbBlack
    'Draw axes labels:
    picFunctionPlot.DrawWidth = 1
    Graph.drawVerticalAxisLabels (F! * L! / 2 / Z!) / 3, 0, -1, "#.##E+#", vbBlue
    Graph.drawHorizontalAxisLabels L! / 6, 0, -1, "#.##E+#", vbBlue
End Sub

'////////////////////[ Force Vs. Deflection ]//////////////////
Private Sub cmdForceVsDeflection_Click()
'For twin beam configuration:
    Dim W!, T!, I!, E!, F!, L!, d!, stp!, s!
    Dim sf!, yMax!, yMin!, xMin!, xMax!
    
    W! = lenWidth.cur_in
    T! = lenThickness.cur_in
    I! = W! * T! ^ 3 / 12  'Area moment of inertia for beam cross section
    E! = strssYoungsModulus.cur_psi
    d! = lenDeflection.cur_in
    F! = frcForce.cur_lbf
    L! = lenLength.cur_in
    'If you need to erase the background "grid",
    'the following code will do it:
       
    'picFunctionPlot.AutoRedraw = True  'enables you to erase the background grid.
    picFunctionPlot.Cls   'clears the background grid.
    'picFunctionPlot.AutoRedraw = False 'disables you to erase the background grid.
    
    'Get ready to setup the graph box:
    sf! = 1.15
    yMax! = F! * sf!
    yMin! = -yMax! / 6
    sf! = d! * 0.19
    xMin! = 0 - sf!
    sf! = d! * 0.1
    xMax! = d! + sf!
    'Set the scales on the graph:
    Graph.SetUp picFunctionPlot, xMin!, yMax!, xMax!, yMin!
    'Set graph caption:
    Frame3.Caption = "Force (pounds) vs. deflection (inches): (For two beams)  "
    'Draw axis grid lines
    Graph.drawGridLinesHorizontal F! / 5, RGB(224, 224, 224)
    Graph.drawGridLinesVertical d! / 5, RGB(224, 224, 224)
    'plot the graph:
    stp! = d! / 100
    For s! = 0! To d! Step stp!
        If s! <> 0 Then
            picFunctionPlot.Line (s!, 2 * ((s! * 12 * E! * I!) / (L! ^ 3)))-(s!, 0), vbRed
        End If
    Next s!
    'Draw the axes:
    picFunctionPlot.DrawWidth = 1
    Graph.drawXaxis vbBlack
    Graph.drawYaxis vbBlack
    'Draw axes ticks:
    picFunctionPlot.DrawWidth = 2
    Graph.drawHorizontalAxisTicks d! / 5, 0, 1, vbBlack
    'Graph.drawVerticalAxisTicks
    Graph.drawVerticalAxisTicks F! / 5, 0, 0.7, vbBlack
    'Draw axes labels:
    picFunctionPlot.DrawWidth = 1
    Graph.drawVerticalAxisLabels F! / 5, 0, -1, "#.##E+#", vbBlue
    Graph.drawHorizontalAxisLabels d! / 5, 0, -1, "#.##E+#", vbBlue
End Sub

