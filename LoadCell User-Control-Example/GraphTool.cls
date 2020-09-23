VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

' ///////////////////////[ GraphTool.cls ]/////////////////////////////
'
' Links some useful graphical utility methods to a PictureBox Control.
'
' CopyRight (c) 2000, by: Randy Manning
'
' Language: Visual Basic, V6.0
'
'//////////////////////////////////////////////////////////////////////
Option Explicit

' pBox holds a reference to the PictureBox associated with
' this GraphTool instance.
Private pBox As PictureBox

' usrYmax holds the largest (topMost) users
' y-coordinate value displayable on the PictureBox.
Private usrYmax As Single

' usrYmin holds the smallest (bottomMost) users
' y-coordinate value displayable on the PictureBox.
Private usrYmin As Single

' usrXmin holds the smallest (leftMost) users
' x-coordinate value displayable on the PictureBox.
Private usrXmin As Single

' usrXmax holds the largest (rightMost) users
' x-coordinate value displayable on the PictureBox.
Private usrXmax As Single

' userHeight holds the PictureBox height in user units.
' Calculated and set by the Constructor. indirectly user resettable.
'
'   ======== [*note, the sign of userHeight is negative] =========
Private userHeight As Single

' userWidth holds the PictureBox width in user units.
' Calculated and set by the Constructor. indirectly user resettable.
Private userWidth As Single

' txtNumZeroWidth holds the width of the letter "0" zero
' in the users x axis units.
' Calculated and set by the Constructor. indirectly user resettable.
Private txtNumZeroWidth As Single

' txtNumZeroHeight holds the height of the letter "0" zero
' in the users y axis units.
' Calculated and set by the Constructor. indirectly user resettable.
Private txtNumZeroHeight As Single

'logBaseXaxis holds the base to which the usrLogXaxis() and
'usrPowerXaxis() functions operate.
'set by the Constructor to a default of base 10.
'user resettable by a call to onXaxisSetLogBase().
Private logBaseXaxis As Single

'onXaxisShowLogLabels holds the state of of how labels are to be displayed:
'onXaxisShowLogLabels = FALSE (default)
Private onXaxisShowLogLabels As Boolean

'logBaseYaxis holds the base to which the usrLogYaxis() and
'usrPowerYaxis() functions operate.
'set by the Constructor to a default of base 10.
'user resettable by a call to onYaxisSetLogBase().
Private logBaseYaxis As Single

'onYaxisShowLogLabels holds the state of of how labels are to be displayed:
'onYaxisShowLogLabels = FALSE (default)
Private onYaxisShowLogLabels As Boolean

' GraphTool Constructor method: SetUp()
' The parameter sequence defines a box: (x1,y1) - (x2,y2)
Public Sub SetUp(pBx As PictureBox, xMin As Single, yMax As Single, xMax As Single, yMin As Single)
    
    'Use this method initially, to link the PictureBox to the
    'GraphTool instance.
    
    'GraphTools' internal instance variable assignments:

    'Set the "pBox" (PictureBox) instance variable:
    Set pBox = pBx

    'Assign this GraphTools instance variables:
    'parameter sequence defines the box: (x1,y1) - (x2,y2)
    SetUserScale xMin, yMax, xMax, yMin
    
    'Set the default logBaseXaxis variable to base 10:
    onXaxisSetLogBase 10
    
    'Set the default logBaseYaxis variable to base 10:
    onYaxisSetLogBase 10
    
    'Show normal labels on the X axis:
    onXaxisShowNormalLabels
    
    'Show normal labels on the Y axis:
    onYaxisShowNormalLabels
    
End Sub

' Sets the users' coordinate system scale.
' User-send whenever different user scaling is desired.
' The parameter sequence defines a box: (x1,y1) - (x2,y2)
Public Sub SetUserScale(xMin As Single, yMax As Single, xMax As Single, yMin As Single)
    
    'Parameter to instance variable assignments:
    usrXmin = xMin
    usrYmax = yMax
    usrXmax = xMax
    usrYmin = yMin
    
    'Indirect instance variable assignments:
    '   ======== [*note, the sign of userHeight is negative] =========
    userHeight = usrYmin - usrYmax
    userWidth = usrXmax - usrXmin
    
    'Set the mathematical coordinate system on the PictureBox :
    'i.e., for usage of VB's graphics methods.
    pBox.ScaleHeight = userHeight
    pBox.ScaleTop = usrYmax
    pBox.ScaleWidth = userWidth
    pBox.ScaleLeft = usrXmin
    
    'Set the txtNumZeroWidth & txtNumZeroHeight variables:
    'i.e., to their values in user units.
    txtNumZeroWidth = pBox.TextWidth("0")
    'note, the sign of pBox.TextHeight("0") is negative
    txtNumZeroHeight = -pBox.TextHeight("0")

End Sub

' Draws the x-axis:
Public Sub drawXaxis(color As Long)
    
    'only if the x-axis lies between the visible y-limits:
    If ((usrYmax >= 0#) And (0# >= usrYmin)) Then
    
        'Draw the the x-axis:
        pBox.Line (usrXmin, 0#)-(usrXmax, 0#), color
        
    End If

End Sub

' Draws the y-axis:
Public Sub drawYaxis(color As Long)

    'only if the y-axis lies between the visible x-limits:
    If ((usrXmax >= 0#) And (0# >= usrXmin)) Then
        
        'draw the the y-axis:
        pBox.Line (0#, usrYmin)-(0#, usrYmax), color
       
    End If

End Sub

' Draws vertical grid lines at a given horizontal spacing interval
' through and away from the origin (0,0)
Public Sub drawGridLinesVertical(horizontalSpacing As Single, color As Long)
    
    Dim xLeft As Single, xLoc As Single
    
    'Find the leftmost visible grid line located at an integer number
    'of horizontalSpacing intervals away from the vertical grid line
    'through the origin 0.0, even if the origin is not visible.
    xLeft = (Int(usrXmin / horizontalSpacing)) * horizontalSpacing
    
    'Draw the grid lines:
    For xLoc = xLeft To usrXmax Step horizontalSpacing
        
        pBox.Line (xLoc, usrYmin)-(xLoc, usrYmax), color
        
    Next xLoc
    
End Sub

' Draws horizontal grid lines at a given vertical spacing interval
' through and away from the origin (0,0)
Public Sub drawGridLinesHorizontal(verticalSpacing As Single, color As Long)
    
    Dim yBot As Single, yLoc As Single
    
    'Find the bottommost visible grid line located at an integer number
    'of verticalSpacing intervals away from the horizontal grid line
    'through the origin 0.0, even if the origin is not visible.
    yBot = (Int(usrYmin / verticalSpacing)) * verticalSpacing

    'Draw the grid lines:
    For yLoc = yBot To usrYmax Step verticalSpacing
      
      pBox.Line (usrXmin, yLoc)-(usrXmax, yLoc), color
      
    Next yLoc
    
End Sub

' Draws y-axis ticks along any horizontal coordinate, and of varying length.
' The length of the tick marks is set by the lengthMultiplier parameter;
' Using a value of 1.0 will cause the tick marks to be drawn approximately
' the same length as the width that the letter "0" [zero] would appear in
' the current font setting of the pBox.
' 0.5 => half the width of "0" zero, 2.0 => double... etc.
Public Sub drawVerticalAxisTicks(verticalSpacing As Single, horizontalCoordinate As Single, lengthMultiplier As Single, color As Long)

    Dim yBot As Single, yLoc As Single
    Dim xLeft As Single, xRight As Single
    
    'Find the bottommost visible tick mark located at an integer number
    'of verticalSpacing intervals away from the origin 0.0, even if the
    'origin is not visible.
    yBot = (Int(usrYmin / verticalSpacing)) * verticalSpacing
    
    'Set the left & right extents of the tick marks:
    xLeft = horizontalCoordinate - (txtNumZeroWidth * lengthMultiplier) / 2
    xRight = horizontalCoordinate + (txtNumZeroWidth * lengthMultiplier) / 2
    
    'Draw the the ticks:
    For yLoc = yBot To usrYmax Step verticalSpacing
    
      pBox.Line (xLeft, yLoc)-(xRight, yLoc), color
            
    Next yLoc

End Sub

' Draws x-axis ticks along any vertical coordinate, and of varying length.
' The length of the tick marks is set by the lengthMultiplier parameter;
' Using a value of 1.0 will cause the tick marks to be drawn approximately
' the same length as 60% of the height that the letter "0" [zero] would
' appear in the current font setting of the pBox.
' 0.5 => half the width of 0.6 * "0" zero height , 2.0 => double... etc.
Public Sub drawHorizontalAxisTicks(horizontalSpacing As Single, verticalCoordinate As Single, lengthMultiplier As Single, color As Long)

    Dim xLeft As Single, xLoc As Single
    Dim yTop As Single, yBot As Single

    'Find the leftmost visible tick mark located at an integer number
    'of horizontalSpacing intervals away from the origin 0.0, even if
    'the origin is not visible.
    xLeft = (Int(usrXmin / horizontalSpacing)) * horizontalSpacing
    
    'Set the top & bottom extents of the tick marks:
    yTop = verticalCoordinate + (txtNumZeroHeight * lengthMultiplier) * 0.3
    yBot = verticalCoordinate - (txtNumZeroHeight * lengthMultiplier) * 0.3
    
    'Draw the the ticks:
    For xLoc = xLeft To usrXmax Step horizontalSpacing
        
        pBox.Line (xLoc, yBot)-(xLoc, yTop), color
      
    Next xLoc
    
End Sub
   
' Draws y-axis number labels along any horizontal coordinate
'
' if parameter: sideLeftNegOrRightPos < 0 then
'        draw number labels to the left side of horizontalCoordinate.
'        the numbers will appear right-justified.
' else
'        draw number labels to the right side of horizontalCoordinate.
'        the numbers will appear left-justified.
Public Sub drawVerticalAxisLabels(verticalSpacing As Single, horizontalCoordinate As Single, sideLeftNegOrRightPos As Integer, formatStr As String, color As Long)
    
    Dim yBot As Single, yLoc As Single, powerYloc As Single
    Dim horizOffset As Single, txtWidth As Single
    Dim savedCurX, savedCurY
    Dim oldColor As Long
   
    'save the pBox's CurrentX & CurrentY variables:
    savedCurX = pBox.CurrentX
    savedCurY = pBox.CurrentY
    
    'Find the bottommost visible number label located at an integer number
    'of verticalSpacing intervals away from the origin 0.0, even if the
    'origin is not visible.
    yBot = (Int(usrYmin / verticalSpacing)) * verticalSpacing
    
    'Save the old forecolor:
    oldColor = pBox.ForeColor
    'Set text color
    pBox.ForeColor = color
        
    'Draw the labels:
    For yLoc = yBot To usrYmax Step verticalSpacing
      
        'get the width of the number label in the
        'pBox's scale units:
        If onYaxisShowLogLabels Then
            powerYloc = usrPowerYaxis(yLoc)
            txtWidth = pBox.TextWidth(Format(powerYloc, formatStr))
        Else 'Normal labels:
            txtWidth = pBox.TextWidth(Format(yLoc, formatStr))
        End If
        
        'which side of horizontalCoordinate do we want
        'the number labels to appear ?
        If sideLeftNegOrRightPos < 0 Then
            'we want left side of horizontalCoordinate
            horizOffset = -txtNumZeroWidth - txtWidth
        Else
            'we want right side of horizontalCoordinate
            horizOffset = txtNumZeroWidth
        End If

        'set the pBox's CurrentX & CurrentY variables:
        pBox.CurrentX = horizontalCoordinate + horizOffset
        pBox.CurrentY = yLoc + txtNumZeroHeight / 2
                        
        'Print the number label on the pBox:
        If yLoc <> 0# Then
            If onYaxisShowLogLabels Then
                pBox.Print Format(powerYloc, formatStr)
            Else 'Normal labels:
                pBox.Print Format(yLoc, formatStr)
            End If
        End If
        
    Next yLoc
    
    'Restore the forecolor:
    pBox.ForeColor = oldColor
    
    'Restore the pBox's CurrentX & CurrentY variables:
    pBox.CurrentX = savedCurX
    pBox.CurrentY = savedCurY

End Sub

' Draws x-axis number labels along any vertical coordinate (elevation).
'
' if parameter: sideBotNegOrTopPos < 0 then
'        draw number labels to the bottom side of verticalCoordinate.
' else
'        draw number labels to the top side of verticalCoordinate.
'
' the numbers will appear centered about verticalCoordinate.
Public Sub drawHorizontalAxisLabels(horizontalSpacing As Single, verticalCoordinate As Single, sideBotNegOrTopPos As Integer, formatStr As String, color As Long)
    
    Dim xLeft As Single, xLoc As Single, powerXloc As Single
    Dim vertOffset As Single, txtHeight As Single, txtWidth As Single
    Dim savedCurX, savedCurY
    Dim oldColor As Long
   
    'Save the pBox's CurrentX & CurrentY variables:
    savedCurX = pBox.CurrentX
    savedCurY = pBox.CurrentY
    
    'Find the leftmost visible number label located at an integer number
    'of horizontalSpacing intervals away from the origin 0.0, even if
    'the origin is not visible.
    xLeft = (Int(usrXmin / horizontalSpacing)) * horizontalSpacing
    
    'Save the old forecolor:
    oldColor = pBox.ForeColor
    'Set text color
    pBox.ForeColor = color
    
    'Draw the labels:
    For xLoc = xLeft To usrXmax Step horizontalSpacing
        
        'get the height & width of the number label in the
        'pBox's scale units:
        If onXaxisShowLogLabels Then
            powerXloc = usrPowerXaxis(xLoc)
            txtWidth = pBox.TextWidth(Format(powerXloc, formatStr))
            txtHeight = -pBox.TextHeight(Format(powerXloc, formatStr))
        Else 'Normal labels:
            txtWidth = pBox.TextWidth(Format(xLoc, formatStr))
            txtHeight = -pBox.TextHeight(Format(xLoc, formatStr))
        End If
        
        'which side of verticalCoordinate do we want
        'the number labels to appear ?
        If sideBotNegOrTopPos < 0 Then
            'we want bottom side of verticalCoordinate
            vertOffset = -txtHeight / 2
        Else
            'we want top side of verticalCoordinate
            vertOffset = txtHeight + txtHeight / 4
        End If
        
        'set the pBox's CurrentX & CurrentY variables:
        pBox.CurrentX = xLoc - txtWidth / 2
        pBox.CurrentY = verticalCoordinate + vertOffset
        
        'Print the number label on the pBox:
        If xLoc <> 0# Then
            If onXaxisShowLogLabels Then
                pBox.Print Format(powerXloc, formatStr)
            Else 'Normal labels:
                pBox.Print Format(xLoc, formatStr)
            End If
        End If
        
    Next xLoc
    
    'Restore the forecolor:
    pBox.ForeColor = oldColor
        
    'Restore the pBox's CurrentX & CurrentY variables to
    'the values they had before function entry:
    pBox.CurrentX = savedCurX
    pBox.CurrentY = savedCurY
    
End Sub

'Returns LOG[logBaseXaxis](X) = ln(X)/ln(logBaseXaxis)
'logBaseXaxis is private but may be changed via
'a call to onXaxisSetLogBase().
Public Function usrLogXaxis(X As Single) As Single
    
    'This is the sister function to usrPowerXaxis()
    'use these two functions as an inverse pair for all
    'log / power functions related to graph plots; this way,
    'all you have to do is alter the value of the base [logBaseXaxis]
    'to alter the logrithimic scale of graphs and any associated
    'data value windows.
    ' ... And you won't have to rewrite ANY code!

    'logBaseXaxis = 10#      'log to the base 10
    'logBaseXaxis = Exp(1#)  'log to the base e
    'logBaseXaxis = 2^(1/12) 'log to the base of a musical scale
    
    'VBasic spells the natural log function as: Log()
    usrLogXaxis = Log(X) / Log(logBaseXaxis)

End Function

'Returns y = logBaseXaxis^X
'logBaseXaxis is private but may be changed via
'a call to onXaxisSetLogBase().
Public Function usrPowerXaxis(X As Single) As Double
    
    'This is the sister function to usrLogXaxis()
    'use these two functions as an inverse pair for all
    'log / power functions related to graph plots; this way,
    'all you have to do is alter the value of the base [logBaseXaxis]
    'to alter the logrithimic scale of graphs and any associated
    'data value windows.
    ' ... And you won't have to rewrite ANY code!
    
    'logBaseXaxis = 10#      'power to the base 10
    'logBaseXaxis = Exp(1#)  'power to the base e
    'logBaseXaxis = 2^(1/12) 'power to the base of a musical scale
    
    'usrPowerXaxis = Exp(X * Log(logBaseXaxis))
    usrPowerXaxis = logBaseXaxis ^ X

End Function

'Sets the base to which the usrLogXaxis() and usrPowerXaxis()functions operate.
'
'When using logarithmic axes => the pBox coordinate on the axis,
'(i.e., pBox_MouseMove event X's & Y's) represents the [power number]
'for example: coordinate 4 on the X axis represents logBaseXaxis^4th power.
Public Sub onXaxisSetLogBase(base As Single)

    'onXaxisSetLogBase(10)        'to the base 10
    'onXaxisSetLogBase(Exp(1))    'to the base e
    'onXaxisSetLogBase(2^(1/12))  'to the base of a musical scale

    logBaseXaxis = base
    
End Sub

'onXaxisShowNormalLabels (default)
'A (normal) axis is also a log axis (if) the pBox coordinates and
'labels are to be displayed as powers; however, i/o to the pBox
'coordinates must be done via log/power conversion functions
'supplied by the user. i.e., all values to/from the users real
'world function need to be transformed into the pBox's coordinates
'by the user via their x-form functions.
'On all log axes, tick & label spacings should always be specified
'in terms of powers, the SetUp() parameters should be specified in
'terms of powers as well, reguardless of whether the user wishes to
'label the coordinates as either the number or the power of that
'coordinate.
'For user feedback (mouse cursor position labels)
'convert the coordinates' mouse position (power) into its' number or
'function value via a call to the usrPowerX|Yaxis() function:
' numberValue = usrPowerX|Yaxis(pBox_Coordinate).
'Use the opposite procedure to plot a pBox coordinate or function
'value onto a log axis:
' pBox_Coordinate = usrLogX|Yaxis(numberValue or functionValue)
Public Sub onXaxisShowNormalLabels()

    onXaxisShowLogLabels = False '(default)

End Sub

'onXaxisShowLogarithimicLabels
'A (normal) axis is also a log axis (if) the pBox coordinates and
'labels are to be displayed as powers; however, i/o to the pBox
'coordinates must be done via log/power conversion functions
'supplied by the user. i.e., all values to/from the users real
'world function need to be transformed into the pBox's coordinates
'by the user via their x-form functions.
'On all log axes, tick & label spacings should always be specified
'in terms of powers, the SetUp() parameters should be specified in
'terms of powers as well, reguardless of whether the user wishes to
'label the coordinates as either the number or the power of that
'coordinate.
'For user feedback (mouse cursor position labels)
'convert the coordinates' mouse position (power) into its' number or
'function value via a call to the usrPowerX|Yaxis() function:
' numberValue = usrPowerX|Yaxis(pBox_Coordinate).
'Use the opposite procedure to plot a pBox coordinate or function
'value onto a log axis:
' pBox_Coordinate = usrLogX|Yaxis(numberValue or functionValue)
Public Sub onXaxisShowLogarithimicLabels()

    onXaxisShowLogLabels = True
    
End Sub

'Returns LOG[logBaseYaxis](Y) = ln(Y)/ln(logBaseYaxis)
'logBaseYaxis is private but may be changed via
'a call to onYaxisSetLogBase().
Public Function usrLogYaxis(Y As Single) As Single
    
    'This is the sister function to usrPowerYaxis()
    'use these two functions as an inverse pair for all
    'log / power functions related to graph plots; this way,
    'all you have to do is alter the value of the base [logBaseYaxis]
    'to alter the logrithimic scale of graphs and any associated
    'data value windows.
    ' ... And you won't have to rewrite ANY code!

    'logBaseYaxis = 10#      'log to the base 10
    'logBaseYaxis = Exp(1#)  'log to the base e
    'logBaseYaxis = 2^(1/12) 'log to the base of a musical scale
    
    'VBasic spells the natural log function as: Log()
    usrLogYaxis = Log(Y) / Log(logBaseYaxis)

End Function

'Returns x = logBaseYaxis^Y
'logBaseYaxis is private but may be changed via
'a call to onYaxisSetLogBase().
Public Function usrPowerYaxis(Y As Single) As Double
    
    'This is the sister function to usrLogYaxis()
    'use these two functions as an inverse pair for all
    'log / power functions related to graph plots; this way,
    'all you have to do is alter the value of the base [logBaseYaxis]
    'to alter the logrithimic scale of graphs and any associated
    'data value windows.
    ' ... And you won't have to rewrite ANY code!
    
    'logBaseYaxis = 10#      'power to the base 10
    'logBaseYaxis = Exp(1#)  'power to the base e
    'logBaseYaxis = 2^(1/12) 'power to the base of a musical scale
    
    'usrPowerYaxis = Exp(Y * Log(logBaseYaxis))
    usrPowerYaxis = logBaseYaxis ^ Y

End Function

'Sets the base to which the usrLogYaxis() and usrPowerYaxis()functions operate.
'
'When using logarithmic axes => the pBox coordinate on the axis,
'(i.e., pBox_MouseMove event X's & Y's) represents the [power number]
'for example: coordinate 4 on the Y axis represents logBaseYaxis^4th power.
Public Sub onYaxisSetLogBase(base As Single)

    'onYaxisSetLogBase(10)        'to the base 10
    'onYaxisSetLogBase(Exp(1))    'to the base e
    'onYaxisSetLogBase(2^(1/12))  'to the base of a musical scale

    logBaseYaxis = base
    
End Sub

'onYaxisShowNormalLabels (default)
'A (normal) axis is also a log axis (if) the pBox coordinates and
'labels are to be displayed as powers; however, i/o to the pBox
'coordinates must be done via log/power conversion functions
'supplied by the user. i.e., all values to/from the users real
'world function need to be transformed into the pBox's coordinates
'by the user via their x-form functions.
'On all log axes, tick & label spacings should always be specified
'in terms of powers, the SetUp() parameters should be specified in
'terms of powers as well, reguardless of whether the user wishes to
'label the coordinates as either the number or the power of that
'coordinate.
'For user feedback (mouse cursor position labels)
'convert the coordinates' mouse position (power) into its' number or
'function value via a call to the usrPowerX|Yaxis() function:
' numberValue = usrPowerX|Yaxis(pBox_Coordinate).
'Use the opposite procedure to plot a pBox coordinate or function
'value onto a log axis:
' pBox_Coordinate = usrLogX|Yaxis(numberValue or functionValue)
Public Sub onYaxisShowNormalLabels()

    onYaxisShowLogLabels = False '(default)

End Sub

'onYaxisShowLogarithimicLabels
'A (normal) axis is also a log axis (if) the pBox coordinates and
'labels are to be displayed as powers; however, i/o to the pBox
'coordinates must be done via log/power conversion functions
'supplied by the user. i.e., all values to/from the users real
'world function need to be transformed into the pBox's coordinates
'by the user via their x-form functions.
'On all log axes, tick & label spacings should always be specified
'in terms of powers, the SetUp() parameters should be specified in
'terms of powers as well, reguardless of whether the user wishes to
'label the coordinates as either the number or the power of that
'coordinate.
'For user feedback (mouse cursor position labels)
'convert the coordinates' mouse position (power) into its' number or
'function value via a call to the usrPowerX|Yaxis() function:
' numberValue = usrPowerX|Yaxis(pBox_Coordinate).
'Use the opposite procedure to plot a pBox coordinate or function
'value onto a log axis:
' pBox_Coordinate = usrLogX|Yaxis(numberValue or functionValue)
Public Sub onYaxisShowLogarithimicLabels()

    onYaxisShowLogLabels = True
    
End Sub

'Notes:
'May want to add a scientific notated display mode to the labels.
'If so, then create a public method that converts a given number
'into a scientific notated string containing a specified number
'of significant digits.
'It should display the unconverted number if of small length, and
'the scientific notated number if of large length.
'I would expect that this mode sould be the default.

