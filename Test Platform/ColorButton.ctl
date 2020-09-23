VERSION 5.00
Begin VB.UserControl ColorButton 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1290
   ScaleHeight     =   540
   ScaleWidth      =   1290
   ToolboxBitmap   =   "ColorButton.ctx":0000
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "ColorButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'ColorButton UserControl
'VB6.0 (SP-4)
'By Randy Manning, Jun-2011

'Most of the basic structure of the ColorButton User Control
'was created with the help of the ActiveX Control Interface
'Wizard provided with VB6 Professional Edition. This wizard
'is a very handy tool for creating User Controls. It can save
'you tons of time, keep you from making mistakes and by
'examining the code it generates, teach you a lot about how
'ActiveX Controls are designed to work internally.
'If you have the ActiveX Control Interface Wizard available
'open this control module with it and take a look. The wizard
'is located under the Add_Ins menu. You may have to add it to
'your Add-Ins menu by selecting the Add-In Manager...

'I would also like to acknowledge Rod Stephens, author of the
'book: Visual Basic Graphics Programming, ISBN 0-471-35599-2
'This is the best legacy VB graphics programming book that I
'have ever seen. The book is filled with all kinds of really
'neat 'here's how you do it professionally' graphics stuff.
'It takes you from coordinate-systems, lines and circles all
'the way up to ray-tracing. I used graphic organization and
'drawing techniques from the examples in his book to structure
'and implement the drawing of the ColorButton control graphics.

'I found the slick little code snippets for text outlining and
'engraving methods used to draw the button caption text in an
'example program from Planet-Source-Code. However, the outlining
'code was incomplete and this caused parts of the outline to
'appear blurred on some letters. I discovered the source of the
'problem and fixed the code. The result produces nice-looking
'sharply contrasted outlines in most cases.

'Note, I cosmetically cleaned up ColorButton's code-formatting
'and properly declared its module level variables as suggested by
'Lorin's comment to my original ColorButton code posted on
'Planet-Source-Code. Lorin's advice is a very good example of
'constructive criticism.

Option Explicit

'API Types: [cooridnates: in vbPixels]
Private Type POINTAPI 'For MoveToEx()
    x As Long
    y As Long
End Type

'API Declarations: [cooridnates: in vbPixels]
Private Declare Function Rectangle Lib "gdi32" ( _
    ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long, ByVal X2 As Long, _
    ByVal Y2 As Long) As Long
    
Private Declare Function MoveToEx Lib "gdi32" ( _
    ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, lpPoint As POINTAPI) As Long
    
Private Declare Function LineTo Lib "gdi32" ( _
    ByVal hdc As Long, ByVal X1 As Long, _
    ByVal Y1 As Long) As Long

'For drawing Caption: [cooridnates: in vbPixels]
Private Declare Function TextOut Lib "gdi32" _
    Alias "TextOutA" (ByVal hdc As Long, _
    ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal lpString As String, _
    ByVal nCount As Long) As Long

'Note, the following 'm_' prefixed constants
'and variables are all declared Private and
'they all have 'module-level' scope.

'Default Property Values:
Private Const m_def_ShowFocusRect = True
Private Const m_def_OutlineCaption = False
Private Const m_def_Caption = "ColorButton"
Private Const m_def_Appear3D = False

'Property Variables:
Private m_ShowFocusRect As Boolean
Private m_OutlineCaption As Boolean
Private m_Caption As String
Private m_Appear3D As Boolean

'Event Declarations:
Public Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Public Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Public Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Public Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Public Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick

'Local Variables:
Private m_hasFocus As Boolean
Private m_mouse_Down As Boolean
Private m_mouseDownAndMovingInsideControl As Boolean
Private m_show_Button_Down As Boolean

'//////////[ User Control Internal Events ]////////////////////

Private Sub UserControl_GotFocus()
    m_hasFocus = True
    UserControl_Resize
End Sub

Private Sub UserControl_LostFocus()
    m_hasFocus = False
    UserControl_Resize
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    
    '7-Jun-2011.
    'Added UserControl.SetFocus because you can right-click or
    'scrollWheel-click on a ColorButton and you'll also get a click()
    'event. Right-button and scrollWheel MouseDown() events do not
    'automatically shift the focus to a ColorButton.
    '
    'Example: comment the UserControl.SetFocus statement below and
    'left-click a ColorButton to set focus to it. Then right-click
    'another ColorButton. The first button will still have the focus
    'and you'll also get a click() event from the second button.
    
    UserControl.SetFocus '<- Ensure ColorButton has focus
    m_mouse_Down = True
    m_show_Button_Down = True
    UserControl_Resize
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

'Click-down-and-hold on a ColorButton and drag the cursor back and
'forth between the inside and the outside of the ColorButton's
'boundary... The button visually alternates between up/down states.
'This is the code that causes this visual effect to occur.
'Note, the ColorButton generates a click() event only when both a
'mouse_down and a mouse_up event occur sequentially within the
'control boundary.
'This means that you are not committed to a click() event unless
'you release the mouse button within the boundary of the control.
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    
    'x and y are in vbTwips units relative to
    'top-left coordinate of UserControl
    If m_mouse_Down Then
        If x > 0 And x < ScaleWidth And y > 0 And y < ScaleHeight Then
            'Mouse is down and moving inside the control boundary
            If m_mouseDownAndMovingInsideControl = False Then
                m_mouseDownAndMovingInsideControl = True
                Call mouseDownAndMovingEnterControlBoundary
            End If
        Else
            'Mouse is down and moving outside the control boundary
            If m_mouseDownAndMovingInsideControl = True Then
                m_mouseDownAndMovingInsideControl = False
                Call mouseDownAndMovingExitControlBoundary
            End If
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub
Private Sub mouseDownAndMovingExitControlBoundary()
    m_show_Button_Down = False
    UserControl_Resize
End Sub
Private Sub mouseDownAndMovingEnterControlBoundary()
    m_show_Button_Down = True
    UserControl_Resize
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, _
    x As Single, y As Single)
    
    m_mouse_Down = False
    m_show_Button_Down = False
    UserControl_Resize
    RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace And m_hasFocus Then
        m_show_Button_Down = True
        UserControl_Resize
    End If
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace And m_hasFocus Then
        m_show_Button_Down = False
        UserControl_Resize
        UserControl_Click '<- generate a [Space Bar] click() event
    End If
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'////////////////[ Graphics Section ]/////////////////////////

'This is where the ColorButton, actually the control, gets drawn:
Private Sub UserControl_Resize()
    Dim X1          As Long
    Dim Y1          As Long
    Dim X2          As Long
    Dim Y2          As Long
    Dim cFocusRect  As Long
    Dim cHighlight  As Long
    Dim cShadow     As Long
    Dim oldFC       As Long
    Dim oldPoint    As POINTAPI
    
    Cls 'Erase the control
    
    'User Control ScaleMode is: vbTwips
    'API ScaleMode is: vbPixels
       
    'Draw Caption in up/down state.
    'Set the Caption print location: (CurrentX, CurrentY)
    'such that the printed Caption will appear 'centered'
    'within the control:
    CurrentY = (ScaleHeight / 2) - (TextHeight(m_Caption) / 2)
    CurrentX = (ScaleWidth / 2) - (TextWidth(m_Caption) / 2)
    If m_show_Button_Down Then
        'Offset Caption print location (CurrentX,
        'CurrentY) by one pixel down and one pixel
        'to the right of button up-state location:
        Y1 = ScaleY(1, vbPixels, vbTwips)
        X1 = ScaleX(1, vbPixels, vbTwips)
        CurrentY = CurrentY + Y1
        CurrentX = CurrentX + X1
    End If
    'Draw the Caption: (using current Font).
    If UserControl.Enabled() Then
        If m_OutlineCaption Then  'Draw outlined
            X1 = ScaleX(CurrentX, vbTwips, vbPixels)
            Y1 = ScaleY(CurrentY, vbTwips, vbPixels)
            outline X1, Y1, m_Caption, Picture1.ForeColor()
        Else  'Print normally
            Print m_Caption
        End If
    Else 'Draw with disabled look: engraved
        X1 = ScaleX(CurrentX, vbTwips, vbPixels)
        Y1 = ScaleY(CurrentY, vbTwips, vbPixels)
        engrave X1, Y1, m_Caption
    End If
    
    'Set button outer rectangle colors: 'grayscale'
    cFocusRect = vbBlack
    cHighlight = vbWhite
    cShadow = RGB(128, 128, 128)
    
    'Draw up/down/focus state 3D Button-Border:
    If m_hasFocus Then 'Draw focus rectangles:
        'Convert ScaleHeight and ScaleWidth
        'to vbPixels for API call:
        Y2 = ScaleY(ScaleHeight, vbTwips, vbPixels)
        X2 = ScaleX(ScaleWidth, vbTwips, vbPixels)
        'Draw outer focus rectangle
        oldFC = UserControl.ForeColor()
        UserControl.ForeColor() = cFocusRect
        Rectangle hdc, 0, 0, X2, Y2
        UserControl.ForeColor() = Picture1.ForeColor()
        'Draw inner focus rectangle
        If m_ShowFocusRect Then
            Rectangle hdc, 5, 5, X2 - 5, Y2 - 5
        End If
        UserControl.ForeColor() = oldFC
        If m_show_Button_Down Then 'Show: down - with focus
            'Convert ScaleHeight and ScaleWidth
            'to vbPixels for API call:
            Y2 = ScaleY(ScaleHeight, vbTwips, vbPixels)
            X2 = ScaleX(ScaleWidth, vbTwips, vbPixels)
            'Draw first inner-border rectangle
            oldFC = UserControl.ForeColor()
            UserControl.ForeColor() = cShadow
            Rectangle hdc, 1, 1, X2 - 1, Y2 - 1
            UserControl.ForeColor() = oldFC
        Else 'Show: up - with focus
            'Convert ScaleHeight and ScaleWidth
            'to vbPixels for API call:
            Y2 = ScaleY(ScaleHeight, vbTwips, vbPixels)
            X2 = ScaleX(ScaleWidth, vbTwips, vbPixels)
            'Draw first inner-border rectangle
            oldFC = UserControl.ForeColor()
            UserControl.ForeColor() = cHighlight
            Rectangle hdc, 1, 1, X2 - 1, Y2 - 1
            MoveToEx hdc, X2 - 2, 1, oldPoint
            UserControl.ForeColor() = cShadow
            LineTo hdc, X2 - 2, Y2 - 2
            LineTo hdc, 0, Y2 - 2
            MoveToEx hdc, X2 - 3, 2, oldPoint
            LineTo hdc, X2 - 3, Y2 - 3
            LineTo hdc, 1, Y2 - 3
            UserControl.ForeColor() = oldFC
        End If
    Else 'Show: up - without focus
        'Do not draw: outer focus rectangle.
        'Convert ScaleHeight and ScaleWidth
        'to vbPixels for API call:
        Y2 = ScaleY(ScaleHeight, vbTwips, vbPixels)
        X2 = ScaleX(ScaleWidth, vbTwips, vbPixels)
        oldFC = UserControl.ForeColor()
        UserControl.ForeColor() = cHighlight
        Rectangle hdc, 0, 0, X2, Y2
        MoveToEx hdc, X2 - 1, 0, oldPoint
        UserControl.ForeColor() = cShadow
        LineTo hdc, X2 - 1, Y2 - 1
        LineTo hdc, 0, Y2 - 1
        MoveToEx hdc, X2 - 2, 1, oldPoint
        LineTo hdc, X2 - 2, Y2 - 2
        LineTo hdc, 0, Y2 - 2
        UserControl.ForeColor() = oldFC
    End If
End Sub

'Draw outlined text, (x,y) in vbPixels
Private Sub outline(x As Long, y As Long, capStr As String, _
    outercol As Long)
    'This is a slick little trick... Print the text in its
    'outline color offset by one pixel in each direction
    'about the text's normal origin (x,y). Then print the
    'text in its fore color at its normal origin (x,y).
    'That's it.
    'This simple technique works because almost all letters
    'in any font of any size will have at least one pixel of
    'the background color separating them from the next letter.
    'Your text ends up with a one-pixel outline color in any
    'direction. Just make sure you cover all eight possible
    'directions N, E, S, W, NE, SE, SW, NW that are one
    'pixel away from the text origin (x,y) when printing the
    'outline.
    
    Dim oldFC As Long
    Dim lenCapStr As Integer
    
    lenCapStr = Len(capStr)
    oldFC = UserControl.ForeColor
    'Draw the text Outline in outline color:
    UserControl.ForeColor = outercol
    TextOut UserControl.hdc, x - 1, y - 1, capStr, lenCapStr
    TextOut UserControl.hdc, x - 1, y, capStr, lenCapStr
    TextOut UserControl.hdc, x - 1, y + 1, capStr, lenCapStr
    TextOut UserControl.hdc, x, y - 1, capStr, lenCapStr
    'TextOut UserControl.hDC, x, y, capStr, lenCapStr
    TextOut UserControl.hdc, x, y + 1, capStr, lenCapStr
    TextOut UserControl.hdc, x + 1, y - 1, capStr, lenCapStr
    TextOut UserControl.hdc, x + 1, y, capStr, lenCapStr
    TextOut UserControl.hdc, x + 1, y + 1, capStr, lenCapStr
    'Draw the central text in original ForeColor:
    UserControl.ForeColor = oldFC
    TextOut UserControl.hdc, x, y, capStr, lenCapStr
End Sub

'Draw disabled text, (x,y) in vbPixels
Private Sub engrave(x, y, capStr As String)
    Dim oldFC As Long
    
    'DisabledHighlight:
    oldFC = UserControl.ForeColor
    UserControl.ForeColor = Picture1.FillColor()
    TextOut UserControl.hdc, x + 1, y + 1, capStr, Len(capStr)
    'DisabledColor:
    UserControl.ForeColor = Picture1.BackColor
    TextOut UserControl.hdc, x, y, capStr, Len(capStr)
    UserControl.ForeColor = oldFC
End Sub

'/////////////[ User Control Properties Section ]///////////////

'Initialize Properties for User Control'
'Only happens once; When User Control is first placed on a form.
Private Sub UserControl_InitProperties()
    m_Appear3D = m_def_Appear3D
    Set UserControl.Font = Ambient.Font
    m_Caption = m_def_Caption
    m_OutlineCaption = m_def_OutlineCaption
    m_ShowFocusRect = m_def_ShowFocusRect
    Picture1.ForeColor = &H0&       'OutlineColor holder
    Picture1.BackColor = &HA0A0A0   'DisabledColor holder
    Picture1.FillColor = &HFFFFFF   'DisabledHighlight holder
    StartUp
End Sub

'Load property values from storage.
'Happens every time the control is re-created, but not the first
'time, when the control is first placed on a form, in which case
'UserControl_InitProperties() is called instead.
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Appear3D = PropBag.ReadProperty("Appear3D", m_def_Appear3D)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_OutlineCaption = PropBag.ReadProperty("OutlineCaption", m_def_OutlineCaption)
    m_ShowFocusRect = PropBag.ReadProperty("ShowFocusRect", m_def_ShowFocusRect)
    Picture1.ForeColor = PropBag.ReadProperty("OutlineColor", &H0&)
    Picture1.BackColor = PropBag.ReadProperty("DisabledColor", &HA0A0A0)
    Picture1.FillColor = PropBag.ReadProperty("DisabledHighlight", &HFFFFFF)
    StartUp
End Sub

'Write property values to storage.
'Happens when the control is about to be destroyed during design-time
'only, if one or more of its property values has been changed by the
'programmer. These changes have not yet been saved to the property
'bag. We save them here so that they can be restored by ReadProperties
'the next time the control is re-created, at either design-time or
'run-time.
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appear3D", m_Appear3D, m_def_Appear3D)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("OutlineCaption", m_OutlineCaption, m_def_OutlineCaption)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowFocusRect, m_def_ShowFocusRect)
    Call PropBag.WriteProperty("OutlineColor", Picture1.ForeColor, &H0&)
    Call PropBag.WriteProperty("DisabledColor", Picture1.BackColor, &HA0A0A0)
    Call PropBag.WriteProperty("DisabledHighlight", Picture1.FillColor, &HFFFFFF)
End Sub

'Called from either InitProperties() or ReadProperties()
'after all of the control properties have either been
'initialized or restored.
Private Sub StartUp()
    If m_Appear3D Then
        UserControl.BorderStyle() = vbFixedSingle
    Else
        UserControl.BorderStyle() = vbBSNone
    End If
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get Appear3D() As Boolean
    Appear3D = m_Appear3D
End Property

Public Property Let Appear3D(ByVal New_Appear3D As Boolean)
    m_Appear3D = New_Appear3D
    If m_Appear3D Then
        UserControl.BorderStyle() = vbFixedSingle
    Else
        UserControl.BorderStyle() = vbBSNone
    End If
    PropertyChanged "Appear3D"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl_Resize
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    
    'If UserControl.MouseIcon.Width = 0 Then
    '    'no icon:
    '    UserControl.MousePointer = vbDefault '0
    'Else
    '    'icon:
    '    UserControl.MousePointer = vbCustom '99
    'End If
    'PropertyChanged "MousePointer"
    
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    'Range: 0-15,99
    If New_MousePointer < 0 Then New_MousePointer = 0
    If New_MousePointer > 99 Then New_MousePointer = 99
    If New_MousePointer > 15 And New_MousePointer < 99 Then
        New_MousePointer = 15
    End If
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    UserControl_Resize
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    UserControl_Resize
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,ColorButton
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    UserControl_Resize
    PropertyChanged "Caption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get OutlineCaption() As Boolean
    OutlineCaption = m_OutlineCaption
End Property

Public Property Let OutlineCaption(ByVal New_OutlineCaption As Boolean)
    m_OutlineCaption = New_OutlineCaption
    UserControl_Resize
    PropertyChanged "OutlineCaption"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,false
Public Property Get ShowFocusRect() As Boolean
    ShowFocusRect = m_ShowFocusRect
End Property

Public Property Let ShowFocusRect(ByVal New_ShowFocusRect As Boolean)
    m_ShowFocusRect = New_ShowFocusRect
    UserControl_Resize
    PropertyChanged "ShowFocusRect"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'///////////////////////////////////////////////////////////////
'The following Property procedures use a Picture Control to
'hold their color values. The Picture Control color properties
'are used here only because they automatically provide a handy
'color selection interface in the ColorButton's Properties
'Window at design time. The picturebox is otherwise
'non-functional, disabled and not visible.
'///////////////////////////////////////////////////////////////

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,BackColor
Public Property Get DisabledColor() As OLE_COLOR
Attribute DisabledColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    DisabledColor = Picture1.BackColor
End Property

Public Property Let DisabledColor(ByVal New_DisabledColor As OLE_COLOR)
    Picture1.BackColor() = New_DisabledColor
    UserControl_Resize
    PropertyChanged "DisabledColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,FillColor
Public Property Get DisabledHighlight() As OLE_COLOR
Attribute DisabledHighlight.VB_Description = "Returns/sets the color used to fill in shapes, circles, and boxes."
    DisabledHighlight = Picture1.FillColor
End Property

'For some unknown reason the ActiveX Control Interface Wizard
'always deletes this Property Let procedure whenever I run it.
'So here's a copy that you can re-paste if you need to run
'the ActiveX Control Interface Wizard.
'
'Public Property Let DisabledHighlight(ByVal New_DisabledHighlight As OLE_COLOR)
'    Picture1.FillColor() = New_DisabledHighlight
'    UserControl_Resize
'    PropertyChanged "DisabledHighlight"
'End Property

Public Property Let DisabledHighlight(ByVal New_DisabledHighlight As OLE_COLOR)
    Picture1.FillColor() = New_DisabledHighlight
    UserControl_Resize
    PropertyChanged "DisabledHighlight"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Picture1,Picture1,-1,ForeColor
Public Property Get OutlineColor() As OLE_COLOR
Attribute OutlineColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    OutlineColor = Picture1.ForeColor
End Property

Public Property Let OutlineColor(ByVal New_OutlineColor As OLE_COLOR)
    Picture1.ForeColor() = New_OutlineColor
    UserControl_Resize
    PropertyChanged "OutlineColor"
End Property

