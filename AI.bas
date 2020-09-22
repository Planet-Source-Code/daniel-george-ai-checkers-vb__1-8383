Attribute VB_Name = "AI"
Public MaxAiTime As Long
Public MaxDepth As Byte

Const ARRAY_SIZE As Long = 100000
Const INCREMENT As Long = 1000
Dim AIEndTime As Long
Dim FinishAIMove As Boolean

Public FinishedMove As Boolean

Dim MovesX(1 To 8) As Integer
Dim MovesY(1 To 8) As Integer

Dim states(ARRAY_SIZE) As BoardState
Dim UpperState As Long  'highest state in array
Dim HighestState As Long 'highest used state

Private Sub GenerateState(state As BoardState, ID As Long)
Dim pieceN As Byte
Dim playerN As Byte
Dim spiece As Piece
Dim moveN As Byte
Dim Square1 As CPiece
Dim Square2 As CPiece
Dim ChildID As Long
Dim WinStatus As EndGameType

If GetTickCount >= AIEndTime Then FinishAIMove = True
If FinishAIMove Then Exit Sub

playerN = state.PlayerTurn

For pieceN = 0 To 23
  
  spiece = state.Pieces(pieceN)
  
  If spiece.Live = 0 Then GoTo nextPiece
  If Not spiece.player = playerN Then GoTo nextPiece
  
  
  For moveN = 1 To 4
    
    If spiece.Double = 0 Then 'check piece can make a move in this direction
      If playerN = COMP Then If moveN = 3 Or moveN = 4 Then GoTo nextMove
      If playerN = HUMAN Then If moveN = 1 Or moveN = 2 Then GoTo nextMove
    End If
    
    'can make move with this piecen in this direction
    
    'square is des square
    Square1 = CheckSquare(spiece.X + MovesX(moveN), spiece.Y + MovesY(moveN), state)
    
    If Square1.isSquare = 0 Then GoTo nextMove
    
    
    If Square1.Live = 0 Then
      
      NewState ChildID
      CopyState states(ChildID), state
      states(ChildID).Depth = state.Depth + 1
      states(ChildID).ParentID = ID
      
      'move piece
      states(ChildID).Pieces(pieceN).X = Square1.X
      states(ChildID).Pieces(pieceN).Y = Square1.Y
      
      If Square1.Y = 1 And playerN = HUMAN Then states(ChildID).Pieces(pieceN).Double = 1
      If Square1.Y = 8 And playerN = COMP Then states(ChildID).Pieces(pieceN).Double = 1
      
      states(ChildID).PlayerTurn = NotB(state.PlayerTurn)
      
      If GetTickCount > AIEndTime Then
        'Exit Sub
      End If
      
      WinStatus = CheckWin(states(ChildID))
      If WinStatus = NOEND Then
        
        If states(ChildID).Depth >= MaxDepth Then
          states(ChildID).Evaluation = EvaluateState(states(ChildID))
        Else
          states(ChildID).Evaluation = 0
          GenerateState states(ChildID), ChildID
        End If
      ElseIf WinStatus = HUMANWIN Then
          states(ChildID).Evaluation = -32000 / (states(ChildID).Depth + 1)
      ElseIf WinStatus = COMPWIN Then
          states(ChildID).Evaluation = 32000 / (states(ChildID).Depth + 1)
      ElseIf WinStatus = DRAWEND Then
          states(ChildID).Evaluation = 1
      End If
    
    Else ' contains enemy piece
      
      If Square1.player = playerN Then GoTo nextMove
      
      Square2 = CheckSquare(spiece.X + MovesX(moveN + 4), spiece.Y + MovesY(moveN + 4), state)
      If Square2.isSquare = 0 Then GoTo nextMove 'must be valid + empty
      If Square2.Live = 1 Then GoTo nextMove    'behind the piece to be taken
      
      'now it is valid to take the piece in square1
      NewState ChildID
      
      CopyState states(ChildID), state
      states(ChildID).Depth = state.Depth + 1
      states(ChildID).ParentID = ID
    
      'move piece to des
      states(ChildID).Pieces(pieceN).X = Square2.X
      states(ChildID).Pieces(pieceN).Y = Square2.Y
    
      'kill enemy
      states(ChildID).Pieces(Square1.pNumber).Live = 0
      
      If Square2.Y = 1 And playerN = HUMAN Then states(ChildID).Pieces(pieceN).Double = 1
      If Square2.Y = 8 And playerN = COMP Then states(ChildID).Pieces(pieceN).Double = 1
      
      states(ChildID).PlayerTurn = NotB(state.PlayerTurn)
      
      If GetTickCount > AIEndTime Then
        Exit Sub
      End If
      
      WinStatus = CheckWin(states(ChildID))
      If WinStatus = NOEND Then
        
        If states(ChildID).Depth >= MaxDepth Then
          states(ChildID).Evaluation = EvaluateState(states(ChildID))
        Else
          states(ChildID).Evaluation = 0
          GenerateState states(ChildID), ChildID
        End If
      ElseIf WinStatus = HUMANWIN Then
          states(ChildID).Evaluation = -32000 / (states(ChildID).Depth + 1)
      ElseIf WinStatus = COMPWIN Then
          states(ChildID).Evaluation = 32000 / (states(ChildID).Depth + 1)
      ElseIf WinStatus = DRAWEND Then
          states(ChildID).Evaluation = 1
      End If
    
    
    End If
    
    
nextMove:
  Next moveN
    
nextPiece:
Next pieceN

End Sub

Function EvaluateState(state As BoardState) As Integer
Dim Score(1) As Integer

  For int1 = 0 To 11
    If state.Pieces(int1).Live Then Score(COMP) = Score(COMP) + state.Pieces(int1).Double + 1
    If state.Pieces(int1 + 12).Live Then Score(HUMAN) = Score(HUMAN) + state.Pieces(int1 + 12).Double + 1
  Next
  
  EvaluateState = 500 * Score(COMP) / Score(HUMAN)
  
End Function

Private Sub NewState(Optional ByRef ChildID As Long) ' As BoardState
  
  HighestState = HighestState + 1
  If HighestState >= ARRAY_SIZE Then FinishAIMove = True: Exit Sub 'MsgBox "Out of memory": HighestState = 0 'TODO fix me
  ChildID = HighestState
  
End Sub



Private Sub EnlargeStates()
  UpperState = UpperState + INCREMENT
  'ReDim Preserve states(UpperState)
End Sub


'Public Sub StartMove(state As Boardstate)
'FinishedMove = False
'
'End Sub


Sub LoadMoveValues()
MovesY(1) = 1
MovesX(1) = -1

MovesY(2) = 1
MovesX(2) = 1

MovesY(3) = -1
MovesX(3) = 1

MovesY(4) = -1
MovesX(4) = -1

MovesY(5) = 2
MovesX(5) = -2

MovesY(6) = 2
MovesX(6) = 2

MovesY(7) = -2
MovesX(7) = 2

MovesY(8) = -2
MovesX(8) = -2
End Sub


Public Function FindPossTakes(state As BoardState, playerN As Byte) As Posstakemoves
Dim pieceN As Byte, moveN As Byte, Square1 As CPiece, Square2 As CPiece
Dim newX As Byte, newY As Byte, newX2 As Byte, newY2 As Byte, IndexFrom As Byte, IndexTo As Byte
Dim pieceDone As Byte

ReDim FindPossTakes.poss(0)

pieceDone = 99

  For pieceN = 0 To 23
    If pieceN >= 12 And playerN = COMP Then GoTo nextPiece
    If pieceN <= 11 And playerN = HUMAN Then GoTo nextPiece
    If state.Pieces(pieceN).Live = 0 Then GoTo nextPiece
    
    For moveN = 1 To 4
    
      If Not state.Pieces(pieceN).Double = 1 Then
        If playerN = COMP Then
          If moveN = 1 Or moveN = 2 Then GoTo nextMove
        Else
          If moveN = 3 Or moveN = 4 Then GoTo nextMove
        End If
      End If
      
      newX = state.Pieces(pieceN).X + MovesX(moveN)
      newY = state.Pieces(pieceN).Y + MovesY(moveN)
      
      Square1 = CheckSquare(newX, newY, state)
      
      If Square1.isSquare = 0 Then GoTo nextMove
      If Square1.Live = 0 Then GoTo nextMove
      If Square1.player = playerN Then GoTo nextMove
      'must be real square, with an enemy piece
      
      newX2 = state.Pieces(pieceN).X + MovesX(moveN + 4)
      newY2 = state.Pieces(pieceN).Y + MovesY(moveN + 4)
      
      Square2 = CheckSquare(newX2, newY2, state)
      
      If Square2.isSquare = 0 Then GoTo nextMove
      If Square2.Live = 1 Then GoTo nextMove
      'must be a valid empty square to jump into
      
      If Not pieceDone = pieceN Then
        FindPossTakes.nPoss = FindPossTakes.nPoss + 1
        ReDim Preserve FindPossTakes.poss(FindPossTakes.nPoss)
        FindPossTakes.poss(FindPossTakes.nPoss).MoveFrom = IConvert(state.Pieces(pieceN).X, state.Pieces(pieceN).Y)
        ReDim FindPossTakes.poss(FindPossTakes.nPoss).MoveTo(0)
        FindPossTakes.poss(FindPossTakes.nPoss).nMoveTo = 0
        pieceDone = pieceN
      End If
      
      FindPossTakes.poss(FindPossTakes.nPoss).nMoveTo = FindPossTakes.poss(FindPossTakes.nPoss).nMoveTo + 1
      ReDim Preserve FindPossTakes.poss(FindPossTakes.nPoss).MoveTo(FindPossTakes.poss(FindPossTakes.nPoss).nMoveTo)
      FindPossTakes.poss(FindPossTakes.nPoss).MoveTo(FindPossTakes.poss(FindPossTakes.nPoss).nMoveTo) = IConvert(newX, newY)
      
      
      
nextMove:
    Next moveN
    
nextPiece:
  Next pieceN
  
End Function


Public Sub MakeAIMove(state As BoardState) ', Moves() As move)
Dim startID As Long, chosenN As Long, StartTime As Long
Dim ChosenStates(255) As BoardState, HighestDepth As Byte

Dim TempDepth As Byte 'unnessecary

FrmMain.PicWaiting.Visible = True
FrmMain.MousePointer = 11
StartTime = GetTickCount
AIEndTime = StartTime + MaxAiTime

LoadMoveValues
HighestState = 0

If MaxDepth = 0 Then MaxDepth = 4

HighestDepth = 0
TempDepth = MaxDepth
MaxDepth = 0

FinishAIMove = False
Do

  NewState startID
  CopyState states(startID), state

  MaxDepth = MaxDepth + 1
  'If MaxDepth > TempDepth Then FinishAIMove = True
  
  If Not FinishAIMove Then GenerateState states(startID), startID

  If Not FinishAIMove Then chosenN = PassUpValue
  
  'TODO fixme
  If Not FinishAIMove Then
    If chosenN = 0 Then
      MsgBox "Stalemate": Exit Sub
    End If
  End If
  
  If Not FinishAIMove Then HighestDepth = MaxDepth
  If Not FinishAIMove Then CopyState ChosenStates(HighestDepth), states(chosenN)

  Debug.Print "Depth " & MaxDepth & "  choses " & chosenN
  
Loop While FinishAIMove = False

CopyState CurrentGUIBoard, ChosenStates(HighestDepth)

If HighestDepth = 0 Then MsgBox "Stalemate": Exit Sub
FrmMain.lblMaxDepth = HighestDepth

ReDim Preserve PastBoardStates(UBound(PastBoardStates) + 1)
CopyState PastBoardStates(UBound(PastBoardStates)), CurrentGUIBoard

CurrentPlayerTurn = HUMAN

GUI.DrawBoard CurrentGUIBoard

FrmMain.lblMoveNumber = HighestState
FrmMain.lblMoveRatio = Round((GetTickCount - StartTime) / 1000, 2)

FrmMain.PicWaiting.Visible = False
FrmMain.MousePointer = 0
MaxDepth = TempDepth
End Sub


'Write a sub to pass up values from the static evaluation generated for maxdepth nodes to those above, using minmax
'when depth is odd then minize, when depth is even the maximize

Function PassUpValue() As Long
Dim Depth As Integer, stateId As Long, MaxChoiceValue As Integer, MaxChoice As Long

For Depth = MaxDepth To 2 Step -1
  For stateId = 1 To HighestState
  
  If Not states(stateId).Depth = Depth Then GoTo 1 'next piece
      
  
  If Depth Mod 2 = 0 Then 'minimize
    If states(states(stateId).ParentID).Evaluation > states(stateId).Evaluation Or states(states(stateId).ParentID).Evaluation = 0 Then
      states(states(stateId).ParentID).Evaluation = states(stateId).Evaluation

    End If
  Else                'maximize
    If states(states(stateId).ParentID).Evaluation < states(stateId).Evaluation Or states(states(stateId).ParentID).Evaluation = 0 Then
      states(states(stateId).ParentID).Evaluation = states(stateId).Evaluation

    End If
  End If
  
1
  Next stateId
Next Depth

Depth = 1

MaxChoiceValue = -20000 'this should force a move
For stateId = 1 To HighestState
  If states(stateId).Depth = 1 Then
    If states(stateId).Evaluation > MaxChoiceValue Then MaxChoice = stateId: MaxChoiceValue = states(stateId).Evaluation
  End If
Next stateId

PassUpValue = MaxChoice
'maxchoice now contains the chosen computer move - at last!!

End Function
