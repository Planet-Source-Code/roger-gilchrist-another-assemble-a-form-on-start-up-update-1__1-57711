VERSION 5.00
Begin VB.Form frm_Demo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assemble a Form on Start Up Mk.2a"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8310
   Icon            =   "frm_Demo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frm_Demo.frx":030A
   ScaleHeight     =   388
   ScaleMode       =   0  'User
   ScaleWidth      =   554
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Exit"
      Height          =   375
      Index           =   2
      Left            =   7080
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ListBox lstCols 
      Height          =   2400
      Left            =   7200
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.ListBox lstRows 
      Height          =   2400
      Left            =   6240
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Random"
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdDemo 
      Caption         =   "Apply"
      Height          =   375
      Index           =   0
      Left            =   3720
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.ListBox lstBuild 
      Height          =   840
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.ListBox lstStart 
      Height          =   2400
      Left            =   720
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Timer tmrFlasher 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   4800
   End
   Begin VB.Label lblFooters 
      BackStyle       =   0  'Transparent
      Caption         =   "Start          Build                                              Rows    Cols"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   840
      TabIndex        =   8
      Top             =   3960
      Width           =   6975
   End
   Begin VB.Label lblFlasher 
      BackStyle       =   0  'Transparent
      Caption         =   "The Form may be used after assembly!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   5160
      Width           =   6975
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "frm_Demo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'See mod_blockMover for details
Private bFlash     As Boolean

Private Sub cmdDemo_Click(Index As Integer)

  Select Case Index
   Case 0 'Apply
    'get the values selected by user
    StartfromList = lstStart.ListIndex
    BuildfromList = lstBuild.ListIndex
    RowFromList = lstRows.ListIndex + 3
    ColFromList = lstCols.ListIndex + 3
   Case 1 'Random
    'get values at random set from ListCounts for each value
    StartfromList = Int(Rnd * lstStart.ListCount)
    BuildfromList = Int(Rnd * lstBuild.ListCount)
    RowFromList = Int(Rnd * lstRows.ListCount) + 3
    ColFromList = Int(Rnd * lstCols.ListCount) + 3
   Case 2 'Exit
    'close down
    tmrFlasher.Enabled = False
    Set frm_Blocks = Nothing
    Set frm_Demo = Nothing
    End
  End Select
  'take a new picture so that the form rebuild
  'with the correct highlights in the lists
  DoCapture
  frm_Blocks.Show
  frm_Demo.Hide

End Sub

Public Sub DoCapture()

  CaptureSelf Me, "myself.bmp"

End Sub

Private Sub Form_Activate()

  ' Not really necessary here but if a module modified values
  ' this would provide feedback to the lists

  lstStart.ListIndex = StartMode
  lstBuild.ListIndex = BuildMode
  lstRows.ListIndex = pRows - 3
  lstCols.ListIndex = pCols - 3

End Sub

Private Sub Form_Load()

  Dim I As Long

  tmrFlasher.Enabled = True
  With lstStart
    .AddItem "Shuffled"
    .AddItem "OffEdge"
    .AddItem "OnEdge"
    .AddItem "Centred"
    .AddItem "TopLeft"
    .AddItem "TopMid"
    .AddItem "TopRight"
    .AddItem "BotLeft"
    .AddItem "BotMid"
    .AddItem "BotRight"
    .AddItem "MidLeft"
    .AddItem "MidRight"
    .AddItem "Cross"
    .AddItem "Cross2"
    .AddItem "Swap"
    .AddItem "Quad"
    .AddItem "Wedge"
    .AddItem "wedge2"
  End With
  With lstBuild
    .AddItem "One"
    .AddItem "Wall"
    .AddItem "WallDown"
    .AddItem "All"
    .AddItem "Jump"
    .AddItem "UnSwap"
    .ListIndex = 0
  End With
  For I = 3 To 30
    lstRows.AddItem I
    lstCols.AddItem I
  Next I
  'CHANGE THESE TO CHANGE THE STARTUP BEHAVIOUR
  'set to initial value of 10 X 10 (ListIndex(7) + 3)
  lstRows.ListIndex = 17
  lstCols.ListIndex = 25
  'set the inital counstruction modes
  lstStart.ListIndex = 12
  lstBuild.ListIndex = 3
  '
  'set the public variables that the mod uses
  'NOTE  + 3 converts ListIndex to displayed value
  '
  RowFromList = lstRows.ListIndex + 3
  ColFromList = lstCols.ListIndex + 3
  StartfromList = lstStart.ListIndex
  BuildfromList = lstBuild.ListIndex

End Sub

Private Sub Form_Unload(Cancel As Integer)

  'quit

  tmrFlasher.Enabled = False
  Set frm_Demo = Nothing
  Set frm_Blocks = Nothing
  End

End Sub

Private Sub mnuHelp_Click()

  MsgBox "Thanks to Chris Seelbach for the original idea, hope you like this." & vbNewLine & _
       vbNewLine & _
       "1. Press [Esc] or Click any piece to stop the effect." & vbNewLine & _
       "2. High Row X Col values are slower." & vbNewLine & _
       "2. 'All' starts slow with any single point StartModes." & vbNewLine & _
       "3. 'One' & 'UnSwap' slow down, it takes time for Rnd to find last few pieces." & vbNewLine & _
       "4. 'Jump' is fast; pieces jump into place, but all off-form start styles look the same." & vbNewLine & _
       "5. 'UnSwap' is only for Start'Swap' is 'Jump' for others." & vbNewLine & _
       "6. If you use XP then frm_blocks background is transparent." & vbNewLine & _
       "7. 'myself.bmp' is snapshot used to create effects" & vbNewLine & _
       "    Demo take a snapshot on first run (spoils the effect a bit) once created this won't happen again, unless you delete the file." & vbNewLine & _
       "    Clicking 'Apply' or 'Random' creates a new snapshot to keep the image in sync with the real form." & vbNewLine & _
       "8. Demo displays settings used if 'Random' clicked." & vbNewLine & _
       "9. Right/Bottom edge may be incomplete if col/row number is not an integer factor of width/height." & vbNewLine & _
       "    Factors of 8400/6600 best for Demo.", , "HELP"


End Sub

Private Sub tmrFlasher_Timer()

  'flashes the label
  'bFlash could also be set as a Static Variabl in this procedure
  'but it is less memory intensive to make it a Private variable

  lblFlasher.ForeColor = IIf(bFlash, &HFFC0FF, &HFFFFFF)
  bFlash = Not bFlash

End Sub


':)Code Fixer V2.7.7 (15/12/2004 2:09:41 PM) 3 + 137 = 140 Lines Thanks Ulli for inspiration and lots of code.

