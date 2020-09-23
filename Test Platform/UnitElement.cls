VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnitElementCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvardefVal As Double 'local copy
Private mvarcurVal As Double 'local copy
Private mvaruName As String 'local copy
Private mvarbError As Boolean 'local copy

Public Property Let bError(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bError = 5
    mvarbError = vData
End Property
Public Property Get bError() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bError
    bError = mvarbError
End Property

Public Property Let uName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.uName = 5
    mvaruName = vData
End Property
Public Property Get uName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.uName
    uName = mvaruName
End Property

Public Property Let curVal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.curVal = 5
    mvarcurVal = vData
End Property
Public Property Get curVal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.curVal
    curVal = mvarcurVal
End Property

Public Property Let defVal(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.defVal = 5
    mvardefVal = vData
End Property
Public Property Get defVal() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.defVal
    defVal = mvardefVal
End Property



