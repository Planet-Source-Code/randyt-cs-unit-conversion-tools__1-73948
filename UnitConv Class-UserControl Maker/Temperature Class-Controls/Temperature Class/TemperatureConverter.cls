VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TempConverter"
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
    Dim ndx As Integer
    Dim tmpDbl As Double
   
    On Error GoTo numError
    Select Case changedIndex
        Case Is = 0 '°K
            Units(0).bError = False
            ndx = 0: tmpDbl = newVal: setRange tmpDbl, ndx
            Units(IndexOf("cur_Kelvin")).currentValue = tmpDbl
            If Units(IndexOf("cur_Kelvin")).currentValue < 0 Then Units(IndexOf("cur_Kelvin")).currentValue = 0
            ndx = 1: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Kelvin")).currentValue - 273.15
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Celsius")).currentValue = tmpDbl
            ndx = 2: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Celsius")).currentValue * 9 / 5 + 32
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Fahrenheit")).currentValue = tmpDbl
            ndx = 3: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Kelvin")).currentValue * 9 / 5
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Rankine")).currentValue = tmpDbl
        Case Is = 1 '°C
            Units(1).bError = False
            ndx = 1: tmpDbl = newVal: setRange tmpDbl, ndx
            Units(IndexOf("cur_Celsius")).currentValue = tmpDbl
            If Units(IndexOf("cur_Celsius")).currentValue < -273.15 Then Units(IndexOf("cur_Celsius")).currentValue = -273.15
            ndx = 0: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Celsius")).currentValue + 273.15
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Kelvin")).currentValue = tmpDbl
            ndx = 3: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Kelvin")).currentValue * 9 / 5
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Rankine")).currentValue = tmpDbl
            ndx = 2: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Rankine")).currentValue - 459.67
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Fahrenheit")).currentValue = tmpDbl
        Case Is = 2 '°F
            Units(2).bError = False
            ndx = 2: tmpDbl = newVal: setRange tmpDbl, ndx
            Units(IndexOf("cur_Fahrenheit")).currentValue = tmpDbl
            If Units(IndexOf("cur_Fahrenheit")).currentValue < -459.67 Then Units(IndexOf("cur_Fahrenheit")).currentValue = -459.67
            ndx = 1: Units(ndx).bError = False
            tmpDbl = (Units(IndexOf("cur_Fahrenheit")).currentValue - 32) * 5 / 9
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Celsius")).currentValue = tmpDbl
            ndx = 3: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Fahrenheit")).currentValue + 459.67
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Rankine")).currentValue = tmpDbl
            ndx = 0: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Rankine")).currentValue * 5 / 9
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Kelvin")).currentValue = tmpDbl
        Case Is = 3 '°R
            Units(3).bError = False
            ndx = 3: tmpDbl = newVal: setRange tmpDbl, ndx
            Units(IndexOf("cur_Rankine")).currentValue = tmpDbl
            If Units(IndexOf("cur_Rankine")).currentValue < 0 Then Units(IndexOf("cur_Rankine")).currentValue = 0
            ndx = 2: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Rankine")).currentValue - 459.67
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Fahrenheit")).currentValue = tmpDbl
            ndx = 1: Units(ndx).bError = False
            tmpDbl = (Units(IndexOf("cur_Fahrenheit")).currentValue - 32) * 5 / 9
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Celsius")).currentValue = tmpDbl
            ndx = 0: Units(ndx).bError = False
            tmpDbl = Units(IndexOf("cur_Rankine")).currentValue * 5 / 9
            setRange tmpDbl, ndx
            Units(IndexOf("cur_Kelvin")).currentValue = tmpDbl
    End Select
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
Private Sub setRange(ByRef dblNum As Double, indx As Integer)
    If dblNum > -1E-308 And dblNum < 0 Then
        dblNum = 0
    End If
    If dblNum < 1E-308 And dblNum > 0 Then
        dblNum = 0
    End If
    If dblNum > 1E+308 Then
        dblNum = 1E+308
        Units(indx).bError = True
    End If
    If dblNum < -1E+308 Then
        dblNum = -1E+308
        Units(indx).bError = True
    End If
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
    Call LoadUnit("°Kelvin", "cur_Kelvin", 293.15)
    Call LoadUnit("°Celsius", "cur_Celsius", 20#)
    Call LoadUnit("°Fahrenheit", "cur_Fahrenheit", 68#)
    Call LoadUnit("°Rankine", "cur_Rankine", 527.67)
End Sub

'///////////////////[ PUBLIC PROPERTIES ]/////////////////////

'======================= [°Kelvin]
Public Property Get cur_Kelvin() As Double
    cur_Kelvin = Units(IndexOf("cur_Kelvin")).currentValue
End Property
Public Property Let cur_Kelvin(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_Kelvin"), new_val
End Property
Public Property Get cur_Kelvin_err() As Boolean
    cur_Kelvin_err = Units(IndexOf("cur_Kelvin")).bError
End Property
'======================= [°Celsius]
Public Property Get cur_Celsius() As Double
    cur_Celsius = Units(IndexOf("cur_Celsius")).currentValue
End Property
Public Property Let cur_Celsius(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_Celsius"), new_val
End Property
Public Property Get cur_Celsius_err() As Boolean
    cur_Celsius_err = Units(IndexOf("cur_Celsius")).bError
End Property
'======================= [°Fahrenheit]
Public Property Get cur_Fahrenheit() As Double
    cur_Fahrenheit = Units(IndexOf("cur_Fahrenheit")).currentValue
End Property
Public Property Let cur_Fahrenheit(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_Fahrenheit"), new_val
End Property
Public Property Get cur_Fahrenheit_err() As Boolean
    cur_Fahrenheit_err = Units(IndexOf("cur_Fahrenheit")).bError
End Property
'======================= [°Rankine]
Public Property Get cur_Rankine() As Double
    cur_Rankine = Units(IndexOf("cur_Rankine")).currentValue
End Property
Public Property Let cur_Rankine(ByVal new_val As Double)
    UpdateUnits IndexOf("cur_Rankine"), new_val
End Property
Public Property Get cur_Rankine_err() As Boolean
    cur_Rankine_err = Units(IndexOf("cur_Rankine")).bError
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
