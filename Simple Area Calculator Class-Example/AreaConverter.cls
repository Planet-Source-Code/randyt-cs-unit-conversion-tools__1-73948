VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AreaConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Unit Conversion Class
Option Explicit

Private Type UnitElement
    nameString      As String
    propertyName    As String
    defaultValue    As Double
    currentValue    As Double
    bError          As Boolean
End Type
'local array to hold property values
Private Units() As UnitElement
Private UnitsIndex As Integer 'for Class_Initialize()
'IndexOf: key - propertyName, value - Units()index.
Private IndexOf As Collection 

Private Sub LoadUnit(nameStr As String, propName As String, defValue As Double)
'for loading Units() array elements on Class_Initialize()
    UnitsIndex = UnitsIndex + 1 '<- module level variable
    ReDim Preserve Units(UnitsIndex)
    Units(UnitsIndex).nameString = nameStr
    Units(UnitsIndex).propertyName = propName
    Units(UnitsIndex).defaultValue = defValue
    Units(UnitsIndex).currentValue = defValue
    Units(UnitsIndex).bError = False
    'IndexOf: key - propertyName, value - Units()index.
    IndexOf.Add UnitsIndex, propName 
End Sub

'A unit value has been changed.
Private Sub UpdateUnits(changedIndex As Integer, newVal As Double)
'Update all other unit values proportionally
'using the constant ratio principal:
'
'     current_foot      default_foot
'    --------------- = --------------- = CONSTANT_RATIO
'     current_meter     default_meter
'
    Dim divisor As Double
    Dim ndx As Integer
    Dim tmpDbl As Double
    
    'No Units().defaultValue should ever = zero,
    'or division by zero error will occur!
    divisor = Units(changedIndex).defaultValue
    'Set the .currentValue of each unit:
    For ndx = 0 To UBound(Units)
        On Error GoTo numError
        If divisor <> 0 Then
            Units(ndx).bError = False
            tmpDbl = CDbl(Units(ndx).defaultValue) / divisor * newVal
            If tmpDbl > -1E-308 And tmpDbl < 0 Then
                tmpDbl = 0
            End If
            If tmpDbl < 1E-308 And tmpDbl > 0 Then
                tmpDbl = 0
            End If
            If tmpDbl > 1E+308 Then
                tmpDbl = 1E+308
                Units(ndx).bError = True
            End If
            If tmpDbl < -1E+308 Then
                tmpDbl = -1E+308
                Units(ndx).bError = True
            End If
            Units(ndx).currentValue = tmpDbl
        Else
            MsgBox "Division by zero in UpdateUnits()"
            Stop
        End If
    Next
Exit Sub
numError:
    'Debug.Print Err.Description
    Units(ndx).bError = True
    If Sgn(tmpDbl) = 1 Then
        tmpDbl = 1E+308
    Else
        tmpDbl = -1E+308
    End If
    Resume Next
End Sub

'///////////////////[ PUBLIC METHODS ]/////////////////////
'get ubound
Public Function getUnitIndexUbound() As Integer
    getUnitIndexUbound = UBound(Units())
End Function
'get NameString
Public Function getNameString(unitIndex As Integer) As String
    'If unitIndex < 0 Then unitIndex = 0
    'If unitIndex > UBound(Units()) Then unitIndex = UBound(Units())
    getNameString = Units(unitIndex).nameString
End Function
'get PropertyName
Public Function getPropertyName(unitIndex As Integer) As String
    'If unitIndex < 0 Then unitIndex = 0
    'If unitIndex > UBound(Units()) Then unitIndex = UBound(Units())
    getPropertyName = Units(unitIndex).propertyName
End Function
'Get/Let By Index:
Public Property Get unitValue(unitIndex As Integer) As Double
    'If unitIndex < 0 Then unitIndex = 0
    'If unitIndex > UBound(Units()) Then unitIndex = UBound(Units())
    unitValue = Units(unitIndex).currentValue
End Property
Public Property Let unitValue(unitIndex As Integer, new_val As Double)
    'If unitIndex < 0 Then unitIndex = 0
    'If unitIndex > UBound(Units()) Then unitIndex = UBound(Units())
    UpdateUnits unitIndex, new_val
End Property
'Get bError:
Public Property Get getError(unitIndex As Integer) As Boolean
    'If unitIndex < 0 Then unitIndex = 0
    'If unitIndex > UBound(Units()) Then unitIndex = UBound(Units())
    getError = Units(unitIndex).bError
End Property
'//////////////////////////////////////////////////////////

Private Sub Class_Initialize()
    Set IndexOf = New Collection '<- module level declaration
    'load Units() array
    UnitsIndex = -1 '<- module level variable
    '///////////[ Load Units ]//////////////////////
    Call LoadUnit("yd^2", "cur_yd_sqr", 1.1959900462997#)
    Call LoadUnit("ft^2", "cur_ft_sqr", 10.763910416697#)
    Call LoadUnit("in^2", "cur_in_sqr", 1550.0031000062#)
    Call LoadUnit("m^2", "cur_m_sqr", 1#)
    Call LoadUnit("cm^2", "cur_cm_sqr", 10000#)
    Call LoadUnit("mm^2", "cur_mm_sqr", 1000000#)
End Sub

'///////////////////[ PUBLIC PROPERTIES ]/////////////////////

'======================= [yd^2]
Public Property Get cur_yd_sqr() As Double
    cur_yd_sqr = Units(IndexOf("cur_yd_sqr")).currentValue
End Property
Public Property Let cur_yd_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_yd_sqr"), new_val
End Property
Public Property Get cur_yd_sqr_err() As Boolean
    cur_yd_sqr_err = Units(IndexOf("cur_yd_sqr")).bError
End Property
'======================= [ft^2]
Public Property Get cur_ft_sqr() As Double
    cur_ft_sqr = Units(IndexOf("cur_ft_sqr")).currentValue
End Property
Public Property Let cur_ft_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_ft_sqr"), new_val
End Property
Public Property Get cur_ft_sqr_err() As Boolean
    cur_ft_sqr_err = Units(IndexOf("cur_ft_sqr")).bError
End Property
'======================= [in^2]
Public Property Get cur_in_sqr() As Double
    cur_in_sqr = Units(IndexOf("cur_in_sqr")).currentValue
End Property
Public Property Let cur_in_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_in_sqr"), new_val
End Property
Public Property Get cur_in_sqr_err() As Boolean
    cur_in_sqr_err = Units(IndexOf("cur_in_sqr")).bError
End Property
'======================= [m^2]
Public Property Get cur_m_sqr() As Double
    cur_m_sqr = Units(IndexOf("cur_m_sqr")).currentValue
End Property
Public Property Let cur_m_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_m_sqr"), new_val
End Property
Public Property Get cur_m_sqr_err() As Boolean
    cur_m_sqr_err = Units(IndexOf("cur_m_sqr")).bError
End Property
'======================= [cm^2]
Public Property Get cur_cm_sqr() As Double
    cur_cm_sqr = Units(IndexOf("cur_cm_sqr")).currentValue
End Property
Public Property Let cur_cm_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_cm_sqr"), new_val
End Property
Public Property Get cur_cm_sqr_err() As Boolean
    cur_cm_sqr_err = Units(IndexOf("cur_cm_sqr")).bError
End Property
'======================= [mm^2]
Public Property Get cur_mm_sqr() As Double
    cur_mm_sqr = Units(IndexOf("cur_mm_sqr")).currentValue
End Property
Public Property Let cur_mm_sqr(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_mm_sqr"), new_val
End Property
Public Property Get cur_mm_sqr_err() As Boolean
    cur_mm_sqr_err = Units(IndexOf("cur_mm_sqr")).bError
End Property

'Handy External Program Routines:
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
