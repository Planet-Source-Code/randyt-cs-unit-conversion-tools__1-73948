Attribute VB_Name = "WinSubClass"
Option Explicit

Public OldListView1WindowProc As Long

Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const GWL_WNDPROC = (-4)

Public Function NewListView1WindowProc(ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const WM_NCDESTROY = &H82
Const WM_VSCROLL = &H115
Const WM_HSCROLL = &H114
    
    Dim ret As Long

    'If we're being destroyed, restore the original WindowProc.
    If msg = WM_NCDESTROY Then
        ret = SetWindowLong(hwnd, GWL_WNDPROC, OldListView1WindowProc)
        'if ret is not zero then there was no error.
        'Debug.Print "destroyed", ret
    End If

    If msg = WM_VSCROLL Then
        ListView1_ScrollEvent
    End If
    
    If msg = WM_HSCROLL Then
        ListView1_ScrollEvent
    End If
    
    NewListView1WindowProc = CallWindowProc( _
        OldListView1WindowProc, hwnd, msg, wParam, _
        lParam)
End Function

'hide the textbox
Sub ListView1_ScrollEvent()
    'Debug.Print "Scroll Event"
    FormTest.HideTextBox
End Sub
