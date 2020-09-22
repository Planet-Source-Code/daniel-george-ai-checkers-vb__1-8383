Attribute VB_Name = "GUI"
Sub DrawBoard(Tempboard As BoardState)
Dim Lng1 As Long, Lng2 As Long, index1 As Long, Index2 As Long
Static LastSquare As Long, X As Byte, Y As Byte

Static DrawnBoard(0 To 63) As Long
Static PrevDrawn As Byte
Dim NewBoard(0 To 63) As Long

If PrevDrawn = 0 Then
  For Lng1 = 0 To 63
    DrawnBoard(Lng1) = -3
  Next Lng1
  PrevDrawn = 1
End If

For Lng1 = 0 To 63
  NewBoard(Lng1) = -1
Next Lng1

For Lng1 = 0 To 23
  Index2 = General.IConvert(Tempboard.Pieces(Lng1).X, Tempboard.Pieces(Lng1).Y)
  If Tempboard.Pieces(Lng1).Live = 1 Then NewBoard(Index2) = Lng1 ' Else NewBoard(Index2) = -1
Next

If Not SquareSelect = 65 Then NewBoard(SquareSelect) = -2

For Lng1 = 0 To 63
  
  'If FrmMain.opReversBoard = True Then Lng2 = 63 - Lng1 Else Lng2 = Lng1 'reverse
  
  If ReverseBoard = 1 Then Lng2 = 63 - Lng1 Else Lng2 = Lng1
  
  If DrawnBoard(Lng1) <> NewBoard(Lng1) Then
    Select Case NewBoard(Lng1)
      Case -1
        
        General.XYConvert CByte(Lng1), X, Y
        If X Mod 2 <> Y Mod 2 Then
          If Not FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbBlack Then
            FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbBlack
            FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(0).Picture
          End If
        Else
          'If Not FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbBlack Then
            FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbWhite
            FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(0).Picture
          'End If
        End If
      
      Case -2
        
        FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbYellow: FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(1).Picture
        
      
      Case Else
        Select Case Tempboard.Pieces(NewBoard(Lng1)).player
          Case Is = 0
            If Tempboard.Pieces(NewBoard(Lng1)).Double = 0 Then FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbBlue: FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(3).Picture
            If Tempboard.Pieces(NewBoard(Lng1)).Double = 1 Then FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbCyan: FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(2).Picture
          Case Is = 1
            If Tempboard.Pieces(NewBoard(Lng1)).Double = 0 Then FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbRed: FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(4).Picture
            If Tempboard.Pieces(NewBoard(Lng1)).Double = 1 Then FrmMain!Grid1(Lng2).BackColor = ColorConstants.vbMagenta: FrmMain!Grid1(Lng2).Picture = FrmMain!Picture1(5).Picture
        End Select
    End Select
  End If
  
Next

 
For Lng1 = 0 To 63
  DrawnBoard(Lng1) = NewBoard(Lng1)
Next Lng1

SetPlayerLights

End Sub

Sub Redraw(Tempboard As BoardState)
Dim int1 As Long, int2 As Long, int3 As Long



  For int2 = 0 To 63
    
    If ReverseBoard = 1 Then int1 = 63 - int2 Else int1 = int2
    
      If ((int2 - (int2 Mod 8)) / 8) Mod 2 <> (int2 Mod 8) Mod 2 Then
        'If Not FrmMain!Grid1(int1).BackColor = ColorConstants.vbBlack Then
          FrmMain!Grid1(int1).BackColor = ColorConstants.vbBlack
          FrmMain!Grid1(int1).Picture = FrmMain!Picture1(0).Picture
        'End If
      Else
        'If Not FrmMain!Grid1(int1).BackColor = ColorConstants.vbWhite Then
          FrmMain!Grid1(int1).BackColor = ColorConstants.vbWhite
          
          FrmMain!Grid1(int1).Picture = FrmMain!Picture1(0).Picture
        'End If
      End If
    
  Next
  
  For int2 = 0 To 23
    If Tempboard.Pieces(int2).Live = 0 Then GoTo 1
    int1 = General.IConvert(Tempboard.Pieces(int2).X, Tempboard.Pieces(int2).Y)
    If ReverseBoard = 1 Then int1 = 63 - int1
    'If int1 < 0 Then GoTo 1
    
      Select Case Tempboard.Pieces(int2).player
        Case Is = 0
          If Tempboard.Pieces(int2).Double = 0 Then FrmMain!Grid1(int1).BackColor = ColorConstants.vbBlue: FrmMain!Grid1(int1).Picture = FrmMain!Picture1(3).Picture
          If Tempboard.Pieces(int2).Double = 1 Then FrmMain!Grid1(int1).BackColor = ColorConstants.vbCyan: FrmMain!Grid1(int1).Picture = FrmMain!Picture1(2).Picture
        Case Is = 1
          If Tempboard.Pieces(int2).Double = 0 Then FrmMain!Grid1(int1).BackColor = ColorConstants.vbRed: FrmMain!Grid1(int1).Picture = FrmMain!Picture1(4).Picture
          If Tempboard.Pieces(int2).Double = 1 Then FrmMain!Grid1(int1).BackColor = ColorConstants.vbMagenta: FrmMain!Grid1(int1).Picture = FrmMain!Picture1(5).Picture
      End Select
      
1
  Next


If Not SquareSelect = 65 Then FrmMain!Grid1(SquareSelect).BackColor = ColorConstants.vbYellow: FrmMain!Grid1(SquareSelect).Picture = FrmMain!Picture1(1).Picture

SetPlayerLights
End Sub

Sub MakeMove(Moves() As move, ByRef state As BoardState)
Dim int1 As Byte, int2 As Byte, X As Byte, Y As Byte
Dim Square1 As CPiece, Square2 As CPiece, Square3 As CPiece

int1 = UBound(Moves)
If int1 < 1 Then Exit Sub

ReDim Preserve PastBoardStates(UBound(PastBoardStates) + 1)
CopyState PastBoardStates(UBound(PastBoardStates)), CurrentGUIBoard

For int2 = 1 To int1

  With Moves(int2)
  
  If Not .MoveFromIndex = 65 And Not .MoveToIndex = 65 Then
  'valid move
    General.XYConvert .MoveFromIndex, X, Y
    Square1 = CheckSquare(X, Y, state)
    
    If Square1.Live = 1 Then
      SquareSelect = .MoveFromIndex
      'GUI.DrawBoard state
      GUI.DrawBoard state
      Sleep .Delay
      
      SquareSelect = 65
      
      If Not .KillIndex = 65 Then
      
        General.XYConvert .KillIndex, X, Y
        Square3 = CheckSquare(X, Y, state)
        'kill this piece
        If Square3.Live = 1 Then state.Pieces(Square3.pNumber).Live = 0
        
      End If
    
      General.XYConvert .MoveToIndex, X, Y
      state.Pieces(Square1.pNumber).X = X
      state.Pieces(Square1.pNumber).Y = Y
      
      
      SquareSelect = .MoveToIndex
      GUI.DrawBoard state
      Sleep .Delay * 2
      SquareSelect = 65
      GUI.DrawBoard state
      
    End If
  
  Else
  'no move or take piece
  
    If Not .KillIndex = 65 Then
      
      General.XYConvert .KillIndex, X, Y
      Square3 = CheckSquare(X, Y, state)
      'kill this piece
      If Square3.Live = 1 Then state.Pieces(Square3.pNumber).Live = 0
      
    End If
    
  
  End If

End With

Next int2

SetPlayerLights
End Sub

Function CheckUserMove(move As move)
Static DoubleTakePiece

End Function

Sub SetPlayerLights()
  FrmMain.Light(CurrentPlayerTurn).Visible = True
  FrmMain.Light(NotB(CurrentPlayerTurn)).Visible = False
  
  FrmMain.Light(CurrentPlayerTurn + 2).Visible = False
  FrmMain.Light(NotB(CurrentPlayerTurn) + 2).Visible = True
  
  SetPlayernames
  
End Sub

Private Sub SetPlayernames()
  FrmMain.lblNames(0) = PlayerNames(0)
  FrmMain.lblNames(1) = PlayerNames(1)
End Sub
