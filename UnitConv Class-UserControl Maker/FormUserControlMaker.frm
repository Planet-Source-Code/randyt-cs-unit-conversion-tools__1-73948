VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FormUserControlMaker 
   Caption         =   "Unit Conversion Class / User-Control Maker Tool"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8820
   Icon            =   "FormUserControlMaker.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5775
   ScaleWidth      =   8820
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optClass 
      Caption         =   "Make Class"
      Height          =   255
      Left            =   1560
      TabIndex        =   11
      Top             =   2595
      Width           =   1215
   End
   Begin VB.OptionButton optWideBeta 
      Caption         =   "Wide_Beta"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   2595
      Width           =   1215
   End
   Begin VB.OptionButton optNarrow 
      Caption         =   "Narrow"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   2595
      Value           =   -1  'True
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSetPropName 
      Caption         =   "Edit Control Property Name..."
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   5280
      Width           =   2295
   End
   Begin VB.CommandButton cmdMakeUnitControl 
      Caption         =   "Make UnitControl.ctl..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdDeleteItem 
      Caption         =   "Delete Item"
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSetSymbol 
      Caption         =   "Edit Control Symbol..."
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   5280
      Width           =   1695
   End
   Begin VB.CommandButton cmdClearList 
      Caption         =   "Clear List"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   2520
      TabIndex        =   1
      Top             =   120
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
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
   Begin VB.ListBox List1 
      Height          =   2385
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Unit Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Default Value"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Control Symbol"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Control Property Name"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Make User Control Type:"
      Height          =   255
      Left            =   3240
      TabIndex        =   9
      Top             =   2595
      Width           =   1815
   End
End
Attribute VB_Name = "FormUserControlMaker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Unit Conversion User-Control Maker
'creates a simple Unit Conversion User-Control

Option Explicit

Private m_ControlFileTitle As String
Private m_ClassFileTitle As String

Private Sub Form_Load()
    Dim ndx As Integer
    Dim INIString As String
    
    'load List1 with unit category choices
    ndx = 0
    Do While True
        ndx = ndx + 1
        INIString = GetINIString("Categories", CStr(ndx))
        If INIString = "None" Then
            Exit Do
        Else
            If INIString <> "Temperature" Then
                List1.AddItem INIString
            End If
        End If
    Loop
    'Set listbox index
    List1.ListIndex = 30 ' move list down more than needed
    List1.ListIndex = 26 ' set to <- Length
End Sub

'select from unit category choices
Private Sub List1_Click()
    'loads data from the selected unit
    'category into the listview box
    loadListView1
End Sub

Private Sub loadListView1()
'load the listview box with data from the
'INI file selected unit heading
    Dim ndx As Integer
    Dim heading As String
    Dim unitName As String
    Dim unitVal As Double
    Dim tmpStr As String
    Dim locn As Integer
    
    'clear listView1
    ListView1.ListItems.Clear
    'clear listView2
    ListView2.ListItems.Clear
    'set heading to name of selected category in List1
    heading = List1.List(List1.ListIndex)
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
            'add to listView
            ListView1.ListItems.Add , unitName, unitName
            ListView1.ListItems(unitName).SubItems(1) = unitVal
        End If
    Loop
End Sub

'Add selected item to ListView2
Private Sub ListView1_Click()
    Dim lstItmX As ListItem
    
    'Debug.Print ListView1.SelectedItem.Index - 1
    'Debug.Print ListView1.SelectedItem.Text
    'Debug.Print ListView1.SelectedItem.SubItems(1)
    Set lstItmX = ListView2.FindItem(ListView1.SelectedItem.Text)
    If lstItmX Is Nothing Then  'Not duplicate, load it.
        ListView2.ListItems.Add , ListView1.SelectedItem.Text, ListView1.SelectedItem.Text
        ListView2.ListItems(ListView1.SelectedItem.Text).SubItems(1) = ListView1.SelectedItem.SubItems(1)
        ListView2.ListItems(ListView1.SelectedItem.Text).SubItems(2) = ListView1.SelectedItem.Text
        ListView2.ListItems(ListView1.SelectedItem.Text).SubItems(3) = ParsePropNameFrom(ListView1.SelectedItem.Text)
    Else
        MsgBox "Duplicate items not allowed"
    End If
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Dim selText As String
    Dim selVal As Double
    Dim selItem As ListItem
    
    If ColumnHeader.Index = 1 Then
        'toggle sort state
        ListView1.Sorted = Not ListView1.Sorted
        'save selection data:
        selText = ListView1.SelectedItem.Text
        selVal = ListView1.SelectedItem.SubItems(1)
        Set selItem = ListView1.ListItems(selText)
        'reload listview from INI file
        loadListView1
        'restore selection data:
        Set ListView1.SelectedItem = ListView1.ListItems(selText)
    End If
End Sub

Private Sub cmdSetSymbol_Click()
    Dim lstItmX As ListItem
    Dim promptStr As String
    Dim strSymbol As String
    Dim tmpStr As String
    
    Set lstItmX = ListView2.SelectedItem
    If lstItmX Is Nothing Then Exit Sub
    If optClass.Value = True Then
        tmpStr = " Class "
    Else
        tmpStr = " Control "
    End If
    promptStr = "Set" & tmpStr & "Symbol name for unit: '" & ListView2.SelectedItem.Text & "'" & vbCrLf & vbCrLf
    promptStr = promptStr & "This action also re-parses the" & tmpStr & "Property" & vbCrLf
    promptStr = promptStr & "Name to follow the" & tmpStr & "Symbol name."
    strSymbol = InputBox(promptStr, , ListView2.ListItems(ListView2.SelectedItem.Text).SubItems(2))
    If strSymbol <> "" Then
        ListView2.ListItems(ListView2.SelectedItem.Text).SubItems(2) = strSymbol
        ListView2.ListItems(ListView2.SelectedItem.Text).SubItems(3) = ParsePropNameFrom(strSymbol)
    End If
End Sub
Private Sub cmdSetPropName_Click()
    Dim lstItmX As ListItem
    Dim promptStr As String
    Dim strPropName As String
    Dim tmpStr As String
    
    Set lstItmX = ListView2.SelectedItem
    If lstItmX Is Nothing Then Exit Sub
    If optClass.Value = True Then
        tmpStr = " Class "
    Else
        tmpStr = " Control "
    End If
    promptStr = "Set" & tmpStr & "Property Name for unit: '" & ListView2.SelectedItem.Text & "'"
    strPropName = InputBox(promptStr, , ListView2.ListItems(ListView2.SelectedItem.Text).SubItems(3))
    If strPropName <> "" Then
        ListView2.ListItems(ListView2.SelectedItem.Text).SubItems(3) = strPropName
    End If
End Sub
Private Sub cmdDeleteItem_Click()
    Dim lstItmX As ListItem
    Set lstItmX = ListView2.SelectedItem
    If lstItmX Is Nothing Then Exit Sub
    ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub
Private Sub cmdClearList_Click()
    ListView2.ListItems.Clear
End Sub

'Parse a valid PropertyName from inputStr:
Private Function ParsePropNameFrom(inputStr As String) As String
    Dim tmpStr As String
    Dim tmpChar As String * 1
    Dim ctr As Integer
    
    tmpStr = ""
    For ctr = 1 To Len(inputStr)
        tmpChar = Mid(inputStr, ctr, 1)
        Select Case tmpChar
            Case Is = " " 'Replace spaces with underscores.
                tmpStr = tmpStr & "_"
            Case Is = "(" 'Remove
            Case Is = ")" 'Remove
            Case Is = "'" 'Remove
            Case Is = "°" 'Remove
            Case Is = "@" 'Remove
            Case Is = "." 'Replace with underscore
                tmpStr = tmpStr & "_"
            Case Is = "-" 'Replace hyphen with underscore
                tmpStr = tmpStr & "_"
            Case Is = "/" 'Replace with _per_
                tmpStr = tmpStr & "_per_"
            Case Else
                tmpStr = tmpStr & tmpChar
        End Select
    Next
    'replace ^2 with _sqr
    tmpStr = Replace(tmpStr, "^2", "_sqr")
    'replace ^3 with _cube
    tmpStr = Replace(tmpStr, "^3", "_cube")
    'replace double underscores with single underscores
    tmpStr = Replace(tmpStr, "__", "_")
    ParsePropNameFrom = "cur_" & tmpStr
End Function

'Make UnitClass.cls or UnitControl.ctl
Private Sub cmdMakeUnitControl_Click()
    'Check for at least two selected units.
    If Not atLeastTwoUnits Then Exit Sub
    'Check for ambiguous Class/Control Property Names.
    If ambiguousNames Then Exit Sub
    
    'Build a Class Module:
    If optClass.Value = True Then
        BuildUnitClassFromSkeletonFile App.Path & "\Skeletons" & "\SkeletonClass.dat"
    End If
    'Build a Narrow User_Control:
    If optNarrow.Value = True Then
        BuildUnitControlFromSkeletonFile App.Path & "\Skeletons" & "\Skeleton_Narrow.dat"
    End If
    'Build a Wide_Beta User_Control:
    If optWideBeta.Value = True Then
        BuildUnitControlFromSkeletonFile App.Path & "\Skeletons" & "\Skeleton_Wide_Beta.dat"
    End If
End Sub

'Check for at least two selected units
Private Function atLeastTwoUnits() As Boolean
    Dim promptStr As String
    If ListView2.ListItems.Count >= 2 Then
        atLeastTwoUnits = True
    Else
        promptStr = "At least TWO units must be selected."
        MsgBox (promptStr)
        atLeastTwoUnits = False
    End If
End Function

'Check for ambiguous Control Property Names
Private Function ambiguousNames() As Boolean
    Dim ndx As Integer, names() As String
    Dim curName As String, ndx2 As Integer
    Dim ambgName As String, promptStr As String
    Dim ClsCtlStr As String
    
    'assume false return
    ambiguousNames = False
    'load property names into names() array
    For ndx = 1 To ListView2.ListItems.Count
        ReDim Preserve names(ndx)
        names(ndx) = ListView2.ListItems(ndx).SubItems(3)
    Next
    'iterate through each name in the array
    For ndx = 1 To UBound(names())
        curName = names(ndx)
        'iterate through the names again
        For ndx2 = ndx + 1 To UBound(names())
            If curName = names(ndx2) Then
                ambiguousNames = True
                ambgName = curName
            End If
        Next
    Next
    If ambiguousNames Then
        If optClass.Value = True Then
            ClsCtlStr = "Class"
        Else
            ClsCtlStr = "Control"
        End If
        promptStr = "Ambiguous name: '" & ambgName & "' present in " & ClsCtlStr & " Property Names." & vbCrLf & vbCrLf
        promptStr = promptStr & "Rename at least one '" & ambgName & "' " & ClsCtlStr & " Property Name."
        MsgBox (promptStr)
    End If
End Function

'Opens an input text file containing the skeleton of a unit
'conversion class module: SkeletonClass.dat, and an output file
'UnitClass.cls.
'It translates the skeleton file's data into the source code
'necessary to implement the conversions of a specific set of
'units, (len, mass, force ...) and writes that code to an output
'UnitClass.cls class file.
Private Sub BuildUnitClassFromSkeletonFile(skelFile As String)
    Dim LineInputString As String
    Dim classFile As String
    Dim printLineInputString As Boolean
    Dim tmpStr1 As String
    Dim tmpStr2 As String
    Dim symStr As String
    Dim ndx As Integer
    
    On Error GoTo FileErrorBuildClassFile
    
    classFile = App.Path & "\" & "UnitClass.cls"
    If Not SaveClassFile(classFile) Then Exit Sub
    tmpStr1 = Right$(m_ClassFileTitle, 4)
    tmpStr2 = Replace(m_ClassFileTitle, tmpStr1, "")
    m_ClassFileTitle = tmpStr2
    'Debug.Print classFile
    Open skelFile For Input As #1 'skelFileNum
    Open classFile For Output As #2 'classFileNum
    'Debug.Print "here"
    Do While Not EOF(1)
    'Debug.Print "here"
    
        printLineInputString = True
        Line Input #1, LineInputString
        
        If LineInputString = "FLAG_1" Then
            printLineInputString = False
            'Attribute VB_Name = "UnitClass"
            tmpStr2 = "Attribute VB_Name = " & Chr(34) & m_ClassFileTitle & Chr(34)
            Print #2, tmpStr2
        End If
        
        'Do Load Units section:
        If LineInputString = "FLAG_2" Then
            printLineInputString = False
            For ndx = 1 To ListView2.ListItems.Count
                'Class Symbol: (2) -> symStr
                symStr = ListView2.ListItems(ndx).SubItems(2)
                'Class Property Name: (3)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                'Default Value: (1)
                tmpStr2 = ListView2.ListItems(ndx).SubItems(1) & "#"
                'Build it:
                '    Call LoadUnit("right angle", "right_angle", 4)
                tmpStr2 = "    Call LoadUnit(" & Chr(34) & symStr & Chr(34) & ", " & Chr(34) & tmpStr1 & Chr(34) & ", " & tmpStr2 & ")"
                Print #2, tmpStr2
            Next
        End If
        
        'Do Property Accessors section:
        If LineInputString = "FLAG_3" Then
            printLineInputString = False
            '======================= [right angle]
            'Public Property Get right_angle() As Double
            '    right_angle = Units(IndexOf("right_angle")).currentValue
            'End Property
            'Public Property Let right_angle(ByVal new_val As Double)
            '    UpdateUnits IndexOf("right_angle"), new_val
            'End Property
            For ndx = 1 To ListView2.ListItems.Count
                'Class Symbol: (2) -> symStr
                symStr = ListView2.ListItems(ndx).SubItems(2)
                tmpStr1 = "'======================= [" & symStr & "]"
                Print #2, tmpStr1
                'Public Property Get right_angle() As Double
                'Class Property Name: (3)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "Public Property Get " & tmpStr1 & "() As Double"
                Print #2, tmpStr2
                '    right_angle = Units(IndexOf("right_angle")).currentValue
                tmpStr2 = "    " & tmpStr1 & " = Units(IndexOf(" & Chr(34) & tmpStr1 & Chr(34) & ")).currentValue"
                Print #2, tmpStr2
                'End Property
                Print #2, "End Property"
                'Public Property Let right_angle(ByVal new_val As Double)
                tmpStr2 = "Public Property Let " & tmpStr1 & "(ByVal new_val As Double)"
                Print #2, tmpStr2
                '    UpdateUnits IndexOf("right_angle"), new_val
                tmpStr2 = "    UpdateUnits IndexOf(" & Chr(34) & tmpStr1 & Chr(34) & "), new_val"
                Print #2, tmpStr2
                Print #2, "End Property"
                
                'Public Property Get right_angle_err() As Boolean
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "Public Property Get " & tmpStr1 & "_err" & "() As Boolean"
                Print #2, tmpStr2
                '    right_angle_err = Units(IndexOf("right_angle")).bError
                tmpStr2 = "    " & tmpStr1 & "_err" & " = Units(IndexOf(" & Chr(34) & tmpStr1 & Chr(34) & ")).bError"
                Print #2, tmpStr2
                Print #2, "End Property"
            Next
        End If
        
        If printLineInputString Then
            Print #2, LineInputString
        End If
    Loop
    Close #1
    Close #2
    MsgBox "UnitClass file was successfully created!"
Exit Sub
FileErrorBuildClassFile:
    Close #1
    Close #2
    MsgBox "File Error:" & vbCrLf & vbCrLf & "In subroutine: BuildUnitClassFromSkeletonFile()"
End Sub

'Opens an input text file containing the skeleton of a unit
'conversion user control: Skeleton.dat, and an output file
'UnitControl.ctl.
'It translates the skeleton file's data into the source code
'necessary to implement the conversions of a specific set of
'units, (len, mass, force ...) and writes that code to an output
'UnitControl.ctl user-control file.
Private Sub BuildUnitControlFromSkeletonFile(skelFile As String)
    Dim LineInputString As String
    Dim controlFile As String
    Dim printLineInputString As Boolean
    Dim tmpStr1 As String
    Dim tmpStr2 As String
    Dim symStr As String
    Dim ndx As Integer
    
    On Error GoTo FileErrorBuildControlFile
    
    controlFile = App.Path & "\" & "UnitControl.ctl"
    If Not SaveControlFile(controlFile) Then Exit Sub
    tmpStr1 = Right$(m_ControlFileTitle, 4)
    tmpStr2 = Replace(m_ControlFileTitle, tmpStr1, "")
    m_ControlFileTitle = tmpStr2
    'Debug.Print skelFile
    Open skelFile For Input As #1 'skelFileNum
    Open controlFile For Output As #2 'controlFileNum
    'Debug.Print "here"
    Do While Not EOF(1)
    'Debug.Print "here"
    
        printLineInputString = True
        Line Input #1, LineInputString
        
        If LineInputString = "FLAG_0" Then
            printLineInputString = False
            'Begin VB.UserControl UnitControl
            tmpStr2 = "Begin VB.UserControl " & m_ControlFileTitle
            Print #2, tmpStr2
        End If
        If LineInputString = "FLAG_1" Then
            printLineInputString = False
            'Attribute VB_Name = "UnitControl"
            tmpStr2 = "Attribute VB_Name = " & Chr(34) & m_ControlFileTitle & Chr(34)
            Print #2, tmpStr2
        End If
        If LineInputString = "FLAG_2" Then
            printLineInputString = False
            'dim UnitElementCls classes
            For ndx = 1 To ListView2.ListItems.Count
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "Private Unit_" & tmpStr1 & " As New UnitElementCls"
                Print #2, tmpStr2
            Next
        End If
        If LineInputString = "FLAG_3" Then
            printLineInputString = False
            'replace m with tmpStr1
            '    UpdateUnits "m", Unit_m.curVal
            symStr = ListView2.ListItems(1).SubItems(2)
            tmpStr1 = ListView2.ListItems(1).SubItems(3)
            tmpStr2 = "    UpdateUnits " & Chr(34) & symStr & Chr(34) & ", Unit_" & tmpStr1 & ".curVal"
            Print #2, tmpStr2
        End If
        If LineInputString = "FLAG_4" Then
            printLineInputString = False
            'build control property accessors
            For ndx = 1 To ListView2.ListItems.Count
                'one line at a time
                
                'comment line
                'tmpStr1 = ListView2.ListItems(ndx).Text
                symStr = ListView2.ListItems(ndx).SubItems(2)
                tmpStr2 = "'=================[ " & symStr & " ]"
                Print #2, tmpStr2
                
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                'Public Property Get cur_yd() As Double
                tmpStr2 = "Public Property Get " & tmpStr1 & "() As Double"
                Print #2, tmpStr2
                
                '    cur_yd = Unit_yd.curVal
                tmpStr2 = "    " & tmpStr1 & " = Unit_" & tmpStr1 & ".curVal"
                Print #2, tmpStr2
                
                'End Property
                Print #2, "End Property"
                
                'Public Property Let cur_yd(ByVal New_cur_yd As Double)
                tmpStr2 = "Public Property Let " & tmpStr1 & "(ByVal New_" & tmpStr1 & " As Double)"
                Print #2, tmpStr2
                
                '    Unit_yd.curVal = New_cur_yd
                tmpStr2 = "    Unit_" & tmpStr1 & ".curVal = New_" & tmpStr1
                Print #2, tmpStr2
                
                '    UpdateUnits "yd", Unit_yd.curVal
                tmpStr2 = "    UpdateUnits " & Chr(34) & symStr & Chr(34) & ", Unit_" & tmpStr1 & ".curVal"
                Print #2, tmpStr2
                
                '    PropertyChanged "cur_yd"
                tmpStr2 = "    PropertyChanged " & Chr(34) & tmpStr1 & Chr(34)
                Print #2, tmpStr2
                
                'End Property
                Print #2, "End Property"
                
                'Public Property Get cur_yd_err() As Boolean
                tmpStr2 = "Public Property Get " & tmpStr1 & "_err" & "() As Boolean"
                Print #2, tmpStr2
                
                '    cur_yd_err = Unit_yd.bError
                tmpStr2 = "    " & tmpStr1 & "_err" & " = Unit_" & tmpStr1 & ".bError"
                Print #2, tmpStr2
                
                'End Property
                Print #2, "End Property"
                
                'blank line
                Print #2,
            Next
        End If
        If LineInputString = "FLAG_5" Then
            printLineInputString = False
            'fill the initialize routine
            For ndx = 1 To ListView2.ListItems.Count
                '    Unit_yd.defVal = 1.0936132983377
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "    Unit_" & tmpStr1 & ".defVal = " & ListView2.ListItems(ndx).SubItems(1) & "#"
                Print #2, tmpStr2
            Next
        End If
        If LineInputString = "FLAG_6" Then
            printLineInputString = False
            'fill the initProps routine
            For ndx = 1 To ListView2.ListItems.Count
                '    Unit_yd.curVal = SigFigs(Unit_yd.defVal, m_ctl_sigDigits)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "    Unit_" & tmpStr1 & ".curVal = Unit_" & tmpStr1 & ".defVal"
                Print #2, tmpStr2
            Next
        End If
        If LineInputString = "FLAG_7" Then
            printLineInputString = False
            'fill the ReadProps routine
            For ndx = 1 To ListView2.ListItems.Count
                '    Unit_yd.curVal = PropBag.ReadProperty("cur_yd", Unit_yd.defVal)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "    Unit_" & tmpStr1 & ".curVal = PropBag.ReadProperty(" & Chr(34) & tmpStr1 & Chr(34) & ", Unit_" & tmpStr1 & ".defVal)"
                Print #2, tmpStr2
            Next
        End If
        If LineInputString = "FLAG_8" Then
            printLineInputString = False
            'fill the WriteProps routine
            For ndx = 1 To ListView2.ListItems.Count
                '    Call PropBag.WriteProperty("cur_yd", Unit_yd.curVal, Unit_yd.defVal)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "    Call PropBag.WriteProperty(" & Chr(34) & tmpStr1 & Chr(34) & ", Unit_" & tmpStr1 & ".curVal, Unit_" & tmpStr1 & ".defVal)"
                Print #2, tmpStr2
            Next
        End If
        If LineInputString = "FLAG_9" Then
            printLineInputString = False
            'fill the StartUp routine
            For ndx = 1 To ListView2.ListItems.Count
                '    Unit_yd.uName = "yd": UnitCol.Add Unit_yd, "yd"
                symStr = ListView2.ListItems(ndx).SubItems(2)
                tmpStr1 = ListView2.ListItems(ndx).SubItems(3)
                tmpStr2 = "    Unit_" & tmpStr1 & ".uName = " & Chr(34) & symStr & Chr(34) & ": UnitCol.Add Unit_" & tmpStr1 & ", " & Chr(34) & symStr & Chr(34)
                Print #2, tmpStr2
            Next
        End If
        
        If printLineInputString Then
            Print #2, LineInputString
        End If
    Loop
    Close #1
    Close #2
    MsgBox "UnitControl file was successfully created!"
Exit Sub
FileErrorBuildControlFile:
    Close #1
    Close #2
    MsgBox "File Error:" & vbCrLf & vbCrLf & "In subroutine: BuildUnitControlFromSkeletonFile()"
End Sub

Private Function SaveClassFile(fName As String) As Boolean
    CommonDialog1.CancelError = True
    On Error GoTo CancelErrHandler
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.Filter = "cls files (*.cls)|*.cls"
    CommonDialog1.FileName = fName
    CommonDialog1.ShowSave
    fName = CommonDialog1.FileName
    m_ClassFileTitle = CommonDialog1.FileTitle
    SaveClassFile = True
Exit Function
CancelErrHandler:
    'User pressed the Cancel button
    fName = ""
    SaveClassFile = False
End Function

Private Function SaveControlFile(fName As String) As Boolean
    CommonDialog1.CancelError = True
    On Error GoTo CancelErrHandler
    CommonDialog1.Flags = cdlOFNOverwritePrompt
    CommonDialog1.Filter = "ctl files (*.ctl)|*.ctl"
    CommonDialog1.FileName = fName
    CommonDialog1.ShowSave
    fName = CommonDialog1.FileName
    m_ControlFileTitle = CommonDialog1.FileTitle
    SaveControlFile = True
Exit Function
CancelErrHandler:
    'User pressed the Cancel button
    fName = ""
    SaveControlFile = False
End Function

Private Sub optClass_Click()
    ListView2.ColumnHeaders(3).Text = "Class Symbol"
    ListView2.ColumnHeaders(4).Text = "Class Property Name"
    cmdSetSymbol.Caption = "Edit Class Symbol..."
    cmdSetPropName.Caption = "Edit Class Property Name..."
    cmdMakeUnitControl.Caption = "Make UnitClass.cls..."
End Sub

Private Sub optNarrow_Click()
    ListView2.ColumnHeaders(3).Text = "Control Symbol"
    ListView2.ColumnHeaders(4).Text = "Control Property Name"
    cmdSetSymbol.Caption = "Edit Control Symbol..."
    cmdSetPropName.Caption = "Edit Control Property Name..."
    cmdMakeUnitControl.Caption = "Make UnitControl.ctl..."
End Sub

Private Sub optWideBeta_Click()
    ListView2.ColumnHeaders(3).Text = "Control Symbol"
    ListView2.ColumnHeaders(4).Text = "Control Property Name"
    cmdSetSymbol.Caption = "Edit Control Symbol..."
    cmdSetPropName.Caption = "Edit Control Property Name..."
    cmdMakeUnitControl.Caption = "Make UnitControl.ctl..."
End Sub
