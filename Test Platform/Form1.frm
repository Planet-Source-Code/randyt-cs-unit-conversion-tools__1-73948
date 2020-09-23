VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form Form1 
   Caption         =   "User-Control Test Platform:"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8655
   LinkTopic       =   "Form1"
   ScaleHeight     =   2640
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin ProjectTest2.AreaConverter AreaConverter1 
      Height          =   315
      Left            =   6000
      TabIndex        =   4
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ctl_ListIndex   =   4
      ctl_tboxWidth   =   100
   End
   Begin ProjectTest2.LengthConverter LengthConverter1 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   3
      Combo1_Width    =   750
   End
   Begin ProjectTest2.VolumeConverter VolumeConverter1 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   3
      BackColor       =   255
      Combo1_Width    =   975
   End
   Begin ProjectTest2.ColorButton ColorButton1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1920
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
      Left            =   3480
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "The following unit-conversion controls were created with the 'User Control Maker Tool'."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   210
      TabIndex        =   1
      Top             =   240
      Width           =   8415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ColorButton1_Click()
    LengthConverter1.RegisterScriptControl ScriptControl1
    VolumeConverter1.RegisterScriptControl ScriptControl1
    AreaConverter1.RegisterScriptControl ScriptControl1
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

