Attribute VB_Name = "ModINI"
'Module for accessing Units.INI file

Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As String, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Public Function GetINIString(headingStr As String, keyStr As String) As String
    Dim INIfile As String
    Dim tmpStr As String * 255
    
    INIfile = App.Path & "\Units.INI"
    Call GetPrivateProfileString(headingStr, keyStr, "None", tmpStr, Len(tmpStr), INIfile)
    GetINIString = CString(tmpStr)
End Function

Private Function CString(aStr As String) As String
'Returns aStr up to but not including any null (0) terminator
    Dim nullLoc As Long
    
    CString = aStr
    nullLoc = InStr(aStr, Chr(0))
    If nullLoc Then
        'Debug.Print "nullLoc = " & nullLoc
        CString = Left(aStr, nullLoc - 1)
    End If
End Function
