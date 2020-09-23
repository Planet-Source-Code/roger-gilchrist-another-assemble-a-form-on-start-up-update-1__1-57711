Attribute VB_Name = "mod_BlockMover"
Option Explicit
'''
'''suggested by Chris Seelbach's 'A$$emble a Form on start-up"
'''at http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57585&lngWId=1
'''This modifiction contains almost nothing of his code except for a few variable and
'''control names and the picture on frm_Demo.
'''I also changed the control 'piece' from a picturebox to an image control for lower memory overheads
''' I didn't reproduce  Chris's wedge building (required far too much data gathering)
'''
' ver 2
'Thanks to Raul Fragoso's 'Verify if a point is inside a polygon (convex and non-convex) <txtCodeId=32682>
'I have added the wedge build to the program. I have modified his formulea to only deal with triangles
' and tweaked it a bit
'REQUIREMENTS (the following modules are called from this one)
'mod_FormCapture
'mod_XPTransparent
Public StartfromList            As Long
Public BuildfromList            As Long
Public RowFromList              As Long
Public ColFromList              As Long
'these Enums control where the pieces are first placed
'if you think of a new one add an Enum here
'and insert the code to construct it in sub SetStartPositions
Public Enum StartStyle
  Shuffled
  OffEdge
  OnEdge
  Centred
  TopLeft
  TopMid
  TopRight
  BotLeft
  BotMid
  BotRight
  MidLeft
  MidRight
  Cross
  Cross2
  Swap
  Quad
  Wedge
  Wedge2
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Shuffled, OffEdge, OnEdge, Centred, TopLeft, TopMid, TopRight, BotLeft, BotMid, BotRight, MidLeft, MidRight, Cross, Cross2, Swap
Private Quad, Wedge, Wedge2
#End If
'These Enums control how the pieces move to correct positions
'if you think of a new one add an Enum here
'and insert the code to DrawForm to select it and a seperate Sub to execute it
Public Enum BuildStyle
  One
  Wall
  WallDown
  All
  Jump
  UnSwap
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private One, Wall, WallDown, All, Jump
#End If
'Type and Variable to store correct positions
Type Pos
  oX                            As Long
  oY                            As Long
  Done                          As Boolean
End Type
Private OriginalPos()           As Pos
'
'variables used to break up the image
Private arrPieces()             As Variant
Private m_Vertices(1 To 3)      As Pos
Private Pwidth                  As Long
Private pHeight                 As Long
Public pRows                    As Long
Public pCols                    As Long
Private bricks                  As Long
Private brickDown               As Long
Private PI                      As Single
Private m_BuildMode             As Long
Private m_StartMode             As Long

Private Function ATan2(ByVal opp As Single, _
                       ByVal adj As Single) As Single

  'Part of Raul Fragoso's 'point in polygon code'
  'only modifiation is that it calculates Pi inline rather than using a constant
  
  Dim angle As Single

  If PI = 0 Then
    ' only necesary if not already set
    PI = 4 * Atn(1)
  End If
  ' Get the basic angle.
  If Abs(adj) < 0.0001 Then
    angle = PI / 2
   Else
    angle = Abs(Atn(opp / adj))
  End If
  ' See if we are in quadrant 2 or 3.
  If adj < 0 Then
    ' angle > PI/2 or angle < -PI/2.
    angle = PI - angle
  End If
  ' See if we are in quadrant 3 or 4.
  If opp < 0 Then
    angle = -angle
  End If
  ' Return the result.
  ATan2 = angle

End Function

Public Property Get BuildMode() As BuildStyle

  BuildMode = m_BuildMode

End Property

Public Property Let BuildMode(mode As BuildStyle)

  m_BuildMode = mode

End Property

Public Sub CreateThePicture(frmSource As Form, _
                            ByVal strFileName As String, _
                            Optional ByVal bSilent As Boolean = False)

  If Not FileExists(App.Path & "\" & strFileName) Then
    frmSource.Show
    DoEvents
    CaptureActive frmSource, strFileName
    frmSource.Hide
    If Not bSilent Then
      MsgBox "Apologies," & vbNewLine & _
       "The first time you run the program has to take a snapshot of the form you wish to build." & vbNewLine & _
       "As long as you don't delete the file 'myself.bmp' you won't see this message again", , "Snapshot"
    End If
  End If

End Sub

Public Sub CreateThePieces(frmSource As Form, _
                           frmEffect As Form, _
                           picKlip As PictureClip, _
                           arrPieces As Variant, _
                           ByVal strFileName As String, _
                           ByVal Rows As Long, _
                           ByVal Cols As Long, _
                           ByVal Start As StartStyle, _
                           ByVal Build As BuildStyle)

  'frmSource  = the form that will be built
  'frmEffect  = the form that will show the effect
  'picKlip    = a PictureClip control on frmEffect
  'arrPieces  = an Image control on frmEffect (indexed 0)
  'strFileName= a filename to save the image that is used for the effect
  'Next line only hits the first time you run
  
  Dim I As Long

  SetTransparency frmEffect
  FitFormToForm frmSource, frmEffect
  pRows = Rows 'FromList
  pCols = Cols 'FromList
  bricks = pRows * pCols - 1
  brickDown = 0
  On Error Resume Next
  'This only hits the first time you run
  CreateThePicture frm_Demo, "myself.bmp"
  'for reuse you need to unload the pieces
  For I = arrPieces.Count - 1 To 1 Step -1
    Unload arrPieces(I)
  Next I
  With picKlip
    .Cols = pCols
    .Rows = pRows
    .Picture = LoadPicture(App.Path & "\" & strFileName)
    'initialize root control of pieces and get the basic dimensions
    arrPieces(0).Picture = .GraphicCell(0)
    Pwidth = arrPieces(0).Width
    pHeight = arrPieces(0).Height
    'create the pieces
    For I = 1 To bricks
      Load arrPieces(I)
      arrPieces(I).Picture = .GraphicCell(I)
    Next I
  End With 'PictureClip1
  frmEffect.Width = pCols * Pwidth
  frmEffect.Height = pRows * pHeight
  SetStartPositions arrPieces, Start, Build 'StartfromList, BuildfromList
  On Error GoTo 0

End Sub

Private Sub DoAll(arrPieces As Variant)

  'each piece in turn is moved to correct place
  
  Dim I As Long

  For I = 0 To pRows * pCols - 1
    MovePiece arrPieces(I)
  Next I

End Sub

Private Sub DoJump(arrPieces As Variant)

  'this just jumps peces to the correct place, faster but less graphic
  
  Dim S1 As Long

  Do
    S1 = Int(Rnd * (bricks + 1))
  Loop Until OriginalPos(S1).Done = False
  With OriginalPos(S1)
    .Done = True
    arrPieces(S1).Left = .oY
    arrPieces(S1).Top = .oX
  End With 'OriginalPos(S1)

End Sub

Private Sub DoOne(arrPieces As Variant)

  Dim lbrick As Long

  lbrick = Int(Rnd * (bricks + 1))
  Do
    If MovePiece(arrPieces(lbrick)) Then
      Exit Do
    End If
  Loop

End Sub

Private Sub DoUnSwap(arrPieces As Variant)

  '<:-) :WARNING: Untyped Parameters use Variants which use excessive memory.
  'Dim I As Long
  '<:-) :WARNING: Unused Dim commented out
  
  Dim S1 As Long
  Dim S2 As Long
  Dim Tx As Long
  Dim Ty As Long

  S1 = Int(Rnd * pRows * pCols)
  S2 = Int(Rnd * pRows * pCols)
  If Not OriginalPos(S1).Done Then
    If Not OriginalPos(S2).Done Then
      If StartMode <> Swap Then
        Tx = OriginalPos(S2).oX
        Ty = OriginalPos(S2).oY
       Else
        Tx = arrPieces(S1).Top
        Ty = arrPieces(S1).Left
      End If
      arrPieces(S1).Top = arrPieces(S2).Top
      arrPieces(S1).Left = arrPieces(S2).Left
      arrPieces(S2).Top = Tx
      arrPieces(S2).Left = Ty
      TestCorrectPosition arrPieces(S1)
      TestCorrectPosition arrPieces(S2)
    End If
  End If

End Sub

Private Sub DoWall(arrPieces As Variant)

  Do
    If MovePiece(arrPieces(bricks)) Then
      Exit Do
    End If
  Loop
  bricks = bricks - 1
  If bricks < 0 Then
    bricks = 0
  End If

End Sub

Private Sub DoWallDown(arrPieces As Variant)

  Do
    If MovePiece(arrPieces(brickDown)) Then
      Exit Do
    End If
  Loop
  brickDown = brickDown + 1
  If brickDown > bricks Then
    brickDown = bricks
  End If

End Sub

Public Sub DrawForm(arrPieces As Variant)

  Select Case m_BuildMode
   Case Wall
    DoWall arrPieces
   Case WallDown
    DoWallDown arrPieces
   Case One
    DoOne arrPieces
   Case All
    DoAll arrPieces
   Case Jump
    DoJump arrPieces
   Case UnSwap
    'If m_StartMode = Swap Then
    DoUnSwap arrPieces
    'End If
  End Select

End Sub

Public Function FileExists(strFileName As String) As Boolean

  FileExists = LenB(Dir(strFileName))

End Function

Public Sub FitFormToForm(frmSource As Form, _
                         frmEffect As Form)

  With frmSource
    frmEffect.Move .Left, .Top, .Width - 150, .Height - 100
  End With

End Sub

Public Function FullyAssembled(arrPieces As Variant) As Boolean

  Dim I     As Single
  Dim tdone As Long

  For I = 0 To pCols * pRows - 1
    If arrPieces(I).Top = OriginalPos(I).oX Then
      If arrPieces(I).Left = OriginalPos(I).oY Then
        tdone = tdone + 1
      End If
    End If
  Next I
  FullyAssembled = tdone = pCols * pRows

End Function

Private Function GetAngle(ByVal Ax As Single, _
                          ByVal Ay As Single, _
                          ByVal Bx As Single, _
                          ByVal By As Single, _
                          ByVal Cx As Single, _
                          ByVal Cy As Single) As Single

  'Part of Raul Fragoso's 'point in polygon code'
  'I condensed 3 of Raul Fragoso's proceudres DotProduct, GetAngle into one
  'Original notes from Raul Fragoso
  ''' Return the cross product AB x BC.
  ''' The cross product is a vector perpendicular to AB
  ''' and BC having length |AB| * |BC| * Sin(theta) and
  ''' with direction given by the right-hand rule.
  ''' For two vectors in the X-Y plane, the result is a
  ''' vector with X and Y components 0 so the Z component
  ''' gives the vector's length and direction.
  ''' dot product AB · BC.
  ''' Note that AB · BC = |AB| * |BC| * Cos(theta).
  'GetAngle = ATan2(cross_product, dot_product)

  GetAngle = ATan2((Ax - Bx) * (Cy - By) - (Ay - By) * (Cx - Bx), (Ax - Bx) * (Cx - Bx) + (Ay - By) * (Cy - By))

End Function

Public Function InTriangle(ByVal x As Long, _
                           ByVal Y As Long, _
                           ByVal Ax As Single, _
                           ByVal Ay As Single, _
                           ByVal Bx As Single, _
                           ByVal By As Single, _
                           ByVal Cx As Single, _
                           ByVal Cy As Single) As Boolean

  'Part of Raul Fragoso's 'point in polygon code'
  'modified from his PointInPolygon routine
  'the coordinates of the triangle points are sent as parameters
  'where his set  m_Vertices array else where (a better approach if the shape is not known)
  'NOTE m_Vertices is 1-based
  
  Dim total_angle As Single
  Dim pt          As Long

  m_Vertices(1).oX = Ax
  m_Vertices(1).oY = Ay
  m_Vertices(2).oX = Bx
  m_Vertices(2).oY = By
  m_Vertices(3).oX = Cx
  m_Vertices(3).oY = Cy
  ' Get the angle between the point and the first and last vertices.
  '(RG Added Explanation) because the For loop can't make this link in the edge joins
  total_angle = GetAngle(m_Vertices(3).oX, m_Vertices(3).oY, x, Y, m_Vertices(1).oX, m_Vertices(1).oY)
  ' Add the angles from the point to each other pair of vertices.
  For pt = 1 To 2
    total_angle = total_angle + GetAngle(m_Vertices(pt).oX, m_Vertices(pt).oY, x, Y, m_Vertices(pt + 1).oX, m_Vertices(pt + 1).oY)
  Next pt
  ' The total angle should be 2 * PI or -2 * PI if the point is in the polygon
  'and close to zero if the point is outside the polygon.
  'PointInPolygon = (Abs(total_angle) > 0.000001)
  '(RG Added)For some one with more math than me, above is the original code
  '          but for some reason while it works perfectly in Raul's code
  '          it didn't work in my code until I changed the test to this
  InTriangle = (Abs(total_angle) > PI / 2)

End Function


Public Function MovePiece(arrPiece As Variant) As Boolean

  'moves a piece from where ever it is to its home position
  'returns true when the piece is in the correct place
  
  Dim XOK        As Boolean
  Dim YOK        As Boolean
  Dim pieceIndex As Long

  pieceIndex = arrPiece.Index
  If OriginalPos(pieceIndex).Done Then
    'already in correct position so exit
    MovePiece = True
   Else
    With arrPiece
      .ZOrder 0
      If Not XOK Then
        If Abs(.Top - OriginalPos(pieceIndex).oX) < pHeight Then
          'the piece is close to correct position so just jump to it
          .Top = OriginalPos(pieceIndex).oX
          XOK = True
         Else
          'piece still needs to be moved
          If .Top < OriginalPos(pieceIndex).oX Then
            .Top = .Top + pHeight
           Else
            .Top = .Top - pHeight
          End If
        End If
      End If
      If Not YOK Then
        If Abs(.Left - OriginalPos(pieceIndex).oY) < Pwidth Then
          'the piece is close to correct position so just jump to it
          .Left = OriginalPos(pieceIndex).oY
          YOK = True
         Else
          'piece still needs to be moved
          If .Left < OriginalPos(pieceIndex).oY Then
            .Left = .Left + Pwidth
           Else
            .Left = .Left - Pwidth
          End If
        End If
      End If
    End With
    TestCorrectPosition arrPiece
    DoEvents
  End If

End Function

Private Function Quadrant(ByVal PX As Long, _
                          ByVal PY As Long) As Long

  'A (0,0)        B  (0,C*W)
  '        |
  '   1    |   2
  '------MX,MY---------
  '   4    |  3
  '        |
  'D(R*H,0)      C (R*H,C*W)
  
  Dim MX As Long
  Dim MY As Long

  MX = pRows * pHeight / 2
  MY = pCols * Pwidth / 2
  If PX >= 0 And PX <= MX Then 'in top half
    If PY <= MY Then 'in left half
      Quadrant = 1 'in left half
     Else
      Quadrant = 2 'in right half
    End If
   Else 'in bottom half
    If PY <= MY Then
      Quadrant = 3 'in left half
     Else
      Quadrant = 4 'in right half
    End If
  End If

End Function

Public Sub SetStartPositions(arrPieces As Variant, _
                             Start As StartStyle, _
                             Build As BuildStyle)

  
  Dim Tx         As Long
  Dim Ty         As Long

  Dim lWidth     As Long
  Dim lHeight    As Long
  Dim lMidWidth  As Long
  Dim lMidHeight As Long
  Dim I          As Long
  Dim SwapTarget As Long
  'Dim WedgeMode  As Long
  StoreCorrectPositions arrPieces
  StartMode = Start
  BuildMode = Build
  lWidth = pCols * Pwidth
  lHeight = pRows * pHeight
  lMidWidth = lWidth / 2
  lMidHeight = lHeight / 2
  If StartMode = Swap Then
    For I = 0 To bricks
      With arrPieces(I)
        .Top = OriginalPos(I).oX
        .Left = OriginalPos(I).oY
      End With
    Next I
  End If
  For I = 0 To bricks
    With arrPieces(I)
      Select Case StartMode
       Case Shuffled 'on screen shuffled
        .Top = Int(Rnd * lHeight)
        .Left = Int(Rnd * lWidth)
       Case OffEdge
        If Rnd > 0.5 Then
          .Top = IIf(Rnd > 0.5, -pHeight, lHeight + pHeight)
          .Left = Int(Rnd * lWidth)
         Else
          .Top = Int(Rnd * pCols * pHeight)
          .Left = IIf(Rnd > 0.5, -Pwidth, lWidth + Pwidth)
        End If
       Case OnEdge
        If Rnd > 0.5 Then
          .Top = IIf(Rnd > 0.5, 0, (pRows - 1) * pHeight)
          .Left = Int(Rnd * lWidth)
         Else
          .Top = Int(Rnd * lHeight - pHeight)
          .Left = IIf(Rnd > 0.5, 0, (pCols - 1) * Pwidth)
        End If
       Case Centred
        .Top = lMidHeight
        .Left = lMidWidth
       Case TopLeft
        .Top = -pHeight
        .Left = -Pwidth
       Case TopMid
        .Top = -pHeight
        .Left = Pwidth * pRows / 2
       Case TopRight
        .Top = -pHeight
        .Left = (Pwidth + 1) * pRows
       Case BotLeft
        .Top = (pRows + 1) * pHeight
        .Left = -Pwidth
       Case BotMid
        .Top = (pRows + 1) * pHeight
        .Left = Pwidth * pRows / 2
       Case BotRight
        .Top = (pRows + 1) * pHeight
        .Left = (Pwidth + 1) * pRows
       Case MidLeft
        .Left = -Pwidth
        .Top = lMidHeight
       Case MidRight
        .Left = lWidth + Pwidth
        .Top = lMidHeight
       Case Cross
        If Rnd > 0.5 Then
          .Left = lMidWidth
          .Top = OriginalPos(I).oX '
         Else
          .Top = lMidHeight
          .Left = OriginalPos(I).oY
        End If
       Case Cross2
        Select Case WedgeID(OriginalPos(I).oX, OriginalPos(I).oY)
         Case 1, 3
          .Left = lMidWidth
          .Top = OriginalPos(I).oX '
         Case Else
          .Top = lMidHeight
          .Left = OriginalPos(I).oY
        End Select
       Case Swap
        Do
          SwapTarget = Int(Rnd * pRows * pCols)
        Loop While SwapTarget = I
        Tx = .Top
        Ty = .Left
        .Top = arrPieces(SwapTarget).Top
        .Left = arrPieces(SwapTarget).Left
        arrPieces(SwapTarget).Top = Tx
        arrPieces(SwapTarget).Left = Ty
        '       End If
       Case Quad
        Select Case Quadrant(OriginalPos(I).oX, OriginalPos(I).oY)
         Case 1
          .Top = -pHeight
          .Left = OriginalPos(I).oY
         Case 2
          .Top = OriginalPos(I).oX
          .Left = (Pwidth + 1) * pCols
         Case 3
          .Top = OriginalPos(I).oX
          .Left = -Pwidth
         Case 4
          .Top = (pHeight + 1) * pRows
          .Left = OriginalPos(I).oY
        End Select
       Case Wedge
        Select Case WedgeID(OriginalPos(I).oX, OriginalPos(I).oY)
         Case 1
          .Top = -pHeight
          .Left = OriginalPos(I).oY
         Case 2
          .Top = OriginalPos(I).oX
          .Left = (Pwidth) * pCols
         Case 3
          .Top = (pHeight + 1) * pRows
          .Left = OriginalPos(I).oY
         Case 4
          .Top = OriginalPos(I).oX
          .Left = -Pwidth
        End Select
       Case Wedge2
        Select Case WedgeID(OriginalPos(I).oX, OriginalPos(I).oY)
         Case 1
          .Left = -Pwidth
          .Top = OriginalPos(I).oX '
         Case 2
          .Top = -pHeight
          .Left = OriginalPos(I).oY
         Case 3
          .Top = OriginalPos(I).oX
          .Left = (Pwidth) * pCols - 1
         Case 4
          .Top = lHeight
          .Left = OriginalPos(I).oY
        End Select
      End Select
      '      .Visible = True
    End With 'piece(I)
  Next I
  'seperate process to stop Swap showig in incorrect position
  arrPieces(0).Parent.Visible = False
  For I = 0 To bricks
    With arrPieces(I)
      .Visible = True
    End With 'piece(I)
  Next I
  arrPieces(0).Parent.Visible = True

End Sub

Public Property Get StartMode() As StartStyle

  StartMode = m_StartMode

End Property

Public Property Let StartMode(mode As StartStyle)

  m_StartMode = mode

End Property

Private Sub StoreCorrectPositions(arrPieces As Variant)

  Dim I As Long
  Dim J As Long
  Dim K As Long

  ReDim OriginalPos(pRows * pCols) As Pos
  'store correct positions
  For I = 0 To pRows - 1
    For J = 0 To pCols - 1
      With OriginalPos(K)
        .oX = I * pHeight
        .oY = J * Pwidth
        arrPieces(K).Top = .oX
        arrPieces(K).Left = .oY
        .Done = False
      End With 'OriginalPos(K)
      K = K + 1
    Next J
  Next I

End Sub

Private Sub TestCorrectPosition(varPiece As Variant)


  With varPiece
    If .Top = OriginalPos(.Index).oX Then
      If .Left = OriginalPos(.Index).oY Then
        OriginalPos(.Index).Done = True
      End If
    End If
  End With 'varPiece

End Sub

Private Function WedgeID(ByVal PX As Long, _
                         ByVal PY As Long) As Long

  'Thanks to Raul Fragoso for the code that supports this
  
  Dim Ax As Single
  Dim Ay As Single
  Dim Bx As Single
  Dim By As Single
  Dim Cx As Single
  Dim Cy As Single
  Dim Dx As Single
  Dim Dy As Single
  Dim MX As Single
  Dim MY As Single

  'A(0,0)  B(0,C*W)
  '   \ 1  /
  '    \  /
  '   4 \/ 2       <- MX,MY
  '     /\
  '    /  \
  '   / 3  \
  'D(R*H,0) C(R*H,C*W)
  MX = pRows * pHeight / 2
  MY = pCols * Pwidth / 2
  Ax = 0
  Ay = 0
  Bx = 0
  By = pCols * Pwidth
  Cx = pRows * pHeight
  Cy = pCols * Pwidth
  Dx = pRows * pHeight
  Dy = 0
  ' the '- 1's and ' + 1's allow it to cope with the problem of being on the edge
  ' which the formulea in InTriangle has problems with
  If InTriangle(PX, PY, Ax - 1, Ay, Bx - 1, By, MX, MY) Then
    WedgeID = 1
   ElseIf InTriangle(PX, PY, MX, MY, Bx - 1, By - 1, Cx + 1, Cy + 1) Then
    WedgeID = 2
   ElseIf InTriangle(PX, PY, Dx, Dy - 1, Cx, Cy, MX - 1, MY) Then
    WedgeID = 3
   Else
    WedgeID = 4
  End If

End Function


':)Code Fixer V2.7.7 (15/12/2004 2:09:48 PM) 85 + 702 = 787 Lines Thanks Ulli for inspiration and lots of code.

