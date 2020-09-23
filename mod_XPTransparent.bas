Attribute VB_Name = "mod_XPTransparent"
Option Explicit
'based on Luke H's comment on 'Dr. Fire Transparent Control' txtCodeId=50160
'I converted it to use object so that you can send any control with a Hwnd to it
'I also added OS detection and unacceptable control detection
'
Private Type OSVERSIONINFO
  dwOSVersionInfoSize   As Long
  dwMajorVersion        As Long
  dwMinorVersion        As Long
  dwBuildNumber         As Long
  dwPlatformId          As Long
  szCSDVersion          As String * 128
End Type
'-----------------------
'Transparency stuff
'-----------------------
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32.dll" (ByVal hWnd As Long, _
                                                                      ByVal crKey As Long, _
                                                                      ByVal bAlpha As Byte, _
                                                                      ByVal dwFlags As Long) As Long
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Public Sub SetTransparency(obj As Object, _
                           Optional ByVal SupressMessages As Boolean = False)

  'This Sub only works in WIn XP on Objects that have the BackColor, Visible and hWnd Properties
  '
  'Call from Form_Load
  'To set form use
  '     SetTransparency Me
  'To set a control use
  '     SetTransparency ControlName
  '
  'SupressMessages:
  'by default this routine displays error messages
  'However if you plan to use it while looping through a form's Controls array
  'then either set the optional parameter to True
  '     Set Transparency ctrl, True
  'Or
  'simple edit the paramer above to
  ' Optional SupressMessages As Boolean = True
  'and the sub will fail Silently.
  'NOTE only the safety stuff is mine.
  'The transparency stuff is pure cut'n'paste
  '
  '-----------------------
  'SAFETY VARIABLES
  '-----------------------
  
  Const LW_KEY    As Long = &H1
  Const W_E       As Long = &H80000
  Const G_E       As Long = (-20)
  Dim CrashReason As Long      ' there are 4 things that can cause this routine to fail
  Dim OldBCol     As Long      'preserve old BackColour in case the control doesn't have hWnd
  Dim objVisible  As Boolean   'preserve visiblity in just in case
  Dim strErrMsg   As String    'create various possible error messages

  '-----------------------
  'Transparency Variables
  '-----------------------
  ' 65516 Luke H used this value for G_E but it didn't always work
  ' so I switched to the value in 'Dr. Fire Transparent Control'
  On Error GoTo NoHwnd
  If WinVersion = "Windows XP" Then
    CrashReason = 1 ' if no BackColor Property
    objVisible = obj.Visible
    obj.Visible = False
    CrashReason = 2 ' if no BackColor Property
    OldBCol = obj.BackColor
    obj.BackColor = W_E
    CrashReason = 3 ' if no hWnd Property
    '-----------------------
    'Transparency code
    With obj
      SetWindowLong .hWnd, G_E, GetWindowLong(.hWnd, G_E) Or W_E
      SetLayeredWindowAttributes .hWnd, W_E, 0, LW_KEY
      .Visible = True
      '-----------------------
    End With 'obj
   Else
    obj.BackColor = OldBCol ' restore BackColor if not in XP
    obj.Visible = objVisible
    If Not SupressMessages Then
      MsgBox "Sorry! You are using " & WinVersion & "." & _
 "Sub SetTransparency uses 'SetLayeredWindowAttributes' API which is only available to XP."
    End If
  End If

Exit Sub

NoHwnd:
  If Not SupressMessages Then
    strErrMsg = "Error(" & Err.Number & ") " & Err.Description
    If Err.Number = 438 Then
      Select Case CrashReason
       Case 1
        strErrMsg = strErrMsg & vbNewLine & _
         "'SetTransparency' only works on controls which have the 'Visible' property."
       Case 2
        strErrMsg = strErrMsg & vbNewLine & _
         "'SetTransparency' only works on controls which have the 'BackColor' property."
       Case 3
        strErrMsg = strErrMsg & vbNewLine & _
         "'SetTransparency' only works on controls which have the 'hWnd' property."
      End Select
    End If
    MsgBox strErrMsg
  End If
  On Error Resume Next
  'turnoff NoHwnd error checking to allow any Properties that do exist to be restored
  obj.BackColor = OldBCol
  obj.Visible = objVisible
  Err.Clear
  On Error GoTo 0

End Sub

Private Function WinVersion() As String

  'MODIFIED FROM MS article
  'Determine If Screen Saver Is Running by Using Visual Basic 6.0' Article ID:315725
  
  Dim osinfo As OSVERSIONINFO

  osinfo.dwOSVersionInfoSize = 148
  osinfo.szCSDVersion = Space$(128)
  GetVersionExA osinfo
  With osinfo
    Select Case .dwPlatformId
     Case 1
      If .dwMinorVersion = 0 Then
        WinVersion = "Windows 95"
       ElseIf .dwMinorVersion = 10 Then
        WinVersion = "Windows 98"
      End If
     Case 2
      If .dwMajorVersion = 3 Then
        WinVersion = "Windows NT 3.51"
       ElseIf .dwMajorVersion = 4 Then
        WinVersion = "Windows NT 4.0"
       ElseIf .dwMajorVersion >= 5 Then
        WinVersion = "Windows XP"
      End If
     Case Else
      WinVersion = "Unknown Windows Version"
    End Select
  End With

End Function

':)Code Fixer V2.7.7 (15/12/2004 2:09:51 PM) 20 + 126 = 146 Lines Thanks Ulli for inspiration and lots of code.

