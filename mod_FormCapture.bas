Attribute VB_Name = "mod_FormCapture"
Option Explicit
'based on 'Screen Capture The Inventive Way'
'at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=33946&lngWId=1
'  Created By: Behrooz Sangani
'  Published Date: 19/04/2002
'  Email:   bs20014@yahoo.com
'  WebSite: http://www.geocities.com/bs20014
'  Legal Copyright: Behrooz Sangani © 19/04/2002
'  Use and modify for free but keep the copyright!
Private picTmp                 As VB.PictureBox
Private Const VK_SNAPSHOT      As Long = &H2C    'Snapshot button
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, _
                                              ByVal bScan As Byte, _
                                              ByVal dwFlags As Long, _
                                              ByVal dwExtraInfo As Long)

Public Sub CaptureActive(Frm As Form, _
                         strFileName As String)

  'call this if you need to capture another form
  'Force the form to show

  Frm.Show
  'allow time for the form to draw and get focus
  DoEvents
  'It is now the form with focus, so
  CaptureSelf Frm, strFileName
  Frm.Hide

End Sub

Public Sub CaptureSelf(Frm As Form, _
                       ByVal strFileName As String)

  'pic As PictureBox,
  'call this if you need to capture the current Form
  '
  'Create a temporary PictureBox to hold and transfer the image
  
  Dim Ctrl As Control

  'check the temporary PictureBox doesn't already exist
  For Each Ctrl In Frm.Controls
    If Ctrl.Name = "picTmp" Then
      GoTo AlreadyExists ' skip creating it
    End If
  Next Ctrl
  Set picTmp = Frm.Controls.Add("VB.PictureBox", "picTmp")
AlreadyExists:
  Clipboard.Clear
  keybd_event VK_SNAPSHOT, 1, 0, 0
  DoEvents
  picTmp.Picture = Clipboard.GetData()
  SavePicture picTmp.Picture, App.Path & "\" & strFileName

End Sub

':)Code Fixer V2.7.7 (15/12/2004 2:09:52 PM) 12 + 42 = 54 Lines Thanks Ulli for inspiration and lots of code.

