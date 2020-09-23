VERSION 5.00
Begin VB.Form FormAreaCalc 
   Caption         =   "Class Tester - Simple Area Calculator"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   330
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
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
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   960
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
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
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "FormAreaCalc.frx":0000
      Left            =   3120
      List            =   "FormAreaCalc.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   1800
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "=  Area"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "----------------------------------------------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   1500
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "X  Length 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Length 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "FormAreaCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'create Length and Area Converter class instances
Private lenCnv1 As New LengthConverter
Private lenCnv2 As New LengthConverter
Private areaCnv1 As New AreaConverter

'arrays to hold combo box lists
Private lenStrings() As String
Private areaStrings() As String

Private Const SigDigits As Integer = 5

'===================================================
'form loading - initialization stuff
'===================================================
Private Sub Form_Load()
    
    'setup unit converter initial values
    lenCnv1.cur_ft = 1
    lenCnv2.cur_ft = 2
    areaCnv1.cur_ft_sqr = lenCnv1.cur_ft * lenCnv2.cur_ft
    
    'setup for length combo boxes
    ReDim lenStrings(5)
    lenStrings(0) = "yd"
    lenStrings(1) = "ft"
    lenStrings(2) = "in"
    lenStrings(3) = "m"
    lenStrings(4) = "cm"
    lenStrings(5) = "mm"
    loadCombo Combo1, lenStrings, 2 ' "in" index
    loadCombo Combo2, lenStrings, 2 ' "in" index
    
    'setup for area combo box
    ReDim areaStrings(5)
    areaStrings(0) = "yd^2"
    areaStrings(1) = "ft^2"
    areaStrings(2) = "in^2"
    areaStrings(3) = "m^2"
    areaStrings(4) = "cm^2"
    areaStrings(5) = "mm^2"
    loadCombo Combo3, areaStrings, 2 ' "in^2" index
    
End Sub

'load strings into combo box and set
'combo box listindex to ndx
Private Sub loadCombo(cmb As ComboBox, strArray() As String, ndx As Integer)
    Dim n As Integer
    'load the combo box list
    For n = 0 To UBound(strArray)
        cmb.AddItem strArray(n)
    Next
    'set initial list item to be displayed
    cmb.ListIndex = ndx
End Sub
'///////////////////////////////////////////////////

'===================================================
' text box key presses
'===================================================
Private Sub Text1_KeyPress(KeyAscii As Integer)
    txtKeyPress KeyAscii, Text1, Combo1, lenCnv1
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
    txtKeyPress KeyAscii, Text2, Combo2, lenCnv2
End Sub
Private Sub txtKeyPress(KeyAscii As Integer, tbox As TextBox, cbo As ComboBox, lenCnv As LengthConverter)
    tbox.BackColor = vbYellow
    If KeyAscii = 13 Then '[Enter] key pressed
        tbox.Text = Val(tbox.Text)
        Select Case cbo.ListIndex
            Case Is = 0 'yd
                lenCnv.cur_yd = Val(tbox.Text)
            Case Is = 1 'ft
                lenCnv.cur_ft = Val(tbox.Text)
            Case Is = 2 'in
                lenCnv.cur_in = Val(tbox.Text)
            Case Is = 3 'm
                lenCnv.cur_m = Val(tbox.Text)
            Case Is = 4 'cm
                lenCnv.cur_cm = Val(tbox.Text)
            Case Is = 5 'mm
                lenCnv.cur_mm = Val(tbox.Text)
        End Select
        tbox.BackColor = vbWindowBackground 'normal
        KeyAscii = 0 'suppress 'beep' sound
        'calculate and show the area
        calculateAndShowArea
    End If
End Sub
Private Sub calculateAndShowArea()
    'Calculate and show the area.
    'It doesn't matter which area converter unit you
    'pick to update here, all the other area unit values
    'will get updated automatically.
    areaCnv1.cur_m_sqr = lenCnv1.cur_m * lenCnv2.cur_m
    'This assignment would work just as well:
    'areaCnv1.cur_ft_sqr = lenCnv1.cur_ft * lenCnv2.cur_ft
    'Show the area.
    Combo3_Click
End Sub
'///////////////////////////////////////////////////

'===================================================
'click events for LENGTH combo boxes
'===================================================
Private Sub Combo1_Click()
    cboClick Combo1, Text1, lenCnv1
End Sub
Private Sub Combo2_Click()
    cboClick Combo2, Text2, lenCnv2
End Sub
Private Sub cboClick(cbo As ComboBox, tbox As TextBox, lenCnv As LengthConverter)
    Select Case cbo.ListIndex
        Case Is = 0 'yd
            tbox.Text = SigFigs(lenCnv.cur_yd, SigDigits)
        Case Is = 1 'ft
            tbox.Text = SigFigs(lenCnv.cur_ft, SigDigits)
        Case Is = 2 'in
            tbox.Text = SigFigs(lenCnv.cur_in, SigDigits)
        Case Is = 3 'm
            tbox.Text = SigFigs(lenCnv.cur_m, SigDigits)
        Case Is = 4 'cm
            tbox.Text = SigFigs(lenCnv.cur_cm, SigDigits)
        Case Is = 5 'mm
            tbox.Text = SigFigs(lenCnv.cur_mm, SigDigits)
    End Select
End Sub
'///////////////////////////////////////////////////

'===================================================
'click event for AREA combo box
'===================================================
Private Sub Combo3_Click()
    Select Case Combo3.ListIndex
        Case Is = 0 'yd^2
            Label1.Caption = SigFigs(areaCnv1.cur_yd_sqr, SigDigits)
        Case Is = 1 'ft^2
            Label1.Caption = SigFigs(areaCnv1.cur_ft_sqr, SigDigits)
        Case Is = 2 'in^2
            Label1.Caption = SigFigs(areaCnv1.cur_in_sqr, SigDigits)
        Case Is = 3 'm^2
            Label1.Caption = SigFigs(areaCnv1.cur_m_sqr, SigDigits)
        Case Is = 4 'cm^2
            Label1.Caption = SigFigs(areaCnv1.cur_cm_sqr, SigDigits)
        Case Is = 5 'mm^2
            Label1.Caption = SigFigs(areaCnv1.cur_mm_sqr, SigDigits)
    End Select
End Sub

'///////////////////////////////////////////////////
'===============================================================
'Return 'dblNumber' rounded to 'intSF' significant figures
'===============================================================
Private Function SigFigs(dblNumber As Double, intSF As Integer) As Double
    Dim negFlag As Integer, tmpNum As Double, factor As Double
    Dim numA As Double, numB As Double, outNum As Double
    'dblNumber = 0 ?
    If dblNumber <> 0 Then
        'make sign of tmpNum <- dblNumber, be positive
        If dblNumber < 0 Then
            tmpNum = -dblNumber: negFlag = -1
        Else
            tmpNum = dblNumber: negFlag = 0
        End If
        'get multiplication/division order-of-magnitude factor
        factor = 10 ^ (-Int(Log(tmpNum) / Log(10)) - 1)
        'numA = tmpNum's significant digits moved to right of
        'decimal point: 0.########
        numA = tmpNum * factor
        'round numA to intSF number of decimal places
        numB = Round(numA, intSF)
        'restore numB to tmpNum's original order-of-magnitude
        outNum = numB / factor 'outNum = (positive)
        'Debug.Print tmpNum, factor, numA, numB, outNum
        'correct outNum for sign if necessary
        If negFlag Then outNum = -outNum
    Else  'dblNumber = 0
        outNum = 0
    End If
    SigFigs = outNum 'return
End Function
'///////////////////////////////////////////////////////////////
