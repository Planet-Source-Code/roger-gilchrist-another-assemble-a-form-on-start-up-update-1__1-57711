VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frm_Blocks 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1260
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrBlocks 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   360
   End
   Begin PicClip.PictureClip picTClipBlocks 
      Left            =   1080
      Top             =   240
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.Image piece 
      Height          =   255
      Index           =   0
      Left            =   1800
      Top             =   360
      Width           =   375
   End
End
Attribute VB_Name = "frm_Blocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'See mod_blockMover for details
Private EscPressed     As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)

  If KeyAscii = vbKeyEscape Then
    EscPressed = True
    KeyAscii = 0
    'eat the KeyPress
    '(not needed here but always a safe thing to do)
  End If

End Sub

Private Sub Form_Load()

  On Error Resume Next
  EscPressed = False
  Randomize Timer
  If RowFromList = 0 Then
    'then frm_demo has not been initialized so do it so that
    'RowFromList, ColFromList, StartfromList, BuildfromList
    ' are set by loading frm_demo
    Load frm_Demo
    ' in your own code you could set them explicitly or with random values
    ' Row and Col minimim value 2 max value any size but depending on your system 10 - 20.
    ' Start and Build check the number of members in StartStyle and BuildStyle enums
  End If
  CreateThePieces frm_Demo, Me, picTClipBlocks, piece, "myself.bmp", RowFromList, ColFromList, StartfromList, BuildfromList
  tmrBlocks.Enabled = True
  On Error GoTo 0

End Sub

Private Sub piece_Click(Index As Integer)

  EscPressed = True

End Sub

Private Sub tmrBlocks_Timer()

  If FullyAssembled(piece) Or EscPressed Then
    tmrBlocks.Enabled = False
    frm_Demo.Show
    Unload Me
   Else
    DrawForm piece
  End If

End Sub

':)Code Fixer V2.7.7 (15/12/2004 2:09:42 PM) 3 + 52 = 55 Lines Thanks Ulli for inspiration and lots of code.
