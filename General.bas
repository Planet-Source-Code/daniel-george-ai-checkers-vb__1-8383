Attribute VB_Name = "General"

Function SlashDirectory(strPath As String) As String
If strPath = "" Then Exit Function

  If Not Right(strPath, 1) = "\" And Not Right(strPath, 1) = "/" Then
    SlashDirectory = strPath & "\"
  End If
End Function

Public Function CheckSquare(X As Byte, Y As Byte, state As BoardState) As CPiece
Dim pieceN As Byte

If X > 8 Or Y > 8 Or X < 1 Or Y < 1 Then CheckSquare.isSquare = 0: Exit Function
If X Mod 2 = Y Mod 2 Then CheckSquare.isSquare = 0: Exit Function

For pieceN = 0 To 23
  If state.Pieces(pieceN).Live Then
    If state.Pieces(pieceN).X = X Then
      If state.Pieces(pieceN).Y = Y Then
        'found piece
        CheckSquare.Double = state.Pieces(pieceN).Double
        CheckSquare.Live = 1
        CheckSquare.player = state.Pieces(pieceN).player
        CheckSquare.pNumber = pieceN
        Exit For
      End If
    End If
  End If
Next pieceN

'defualts even for empty squares
CheckSquare.X = X
CheckSquare.Y = Y
CheckSquare.Index = ((Y - 1) * 8) + X
CheckSquare.isSquare = 1
        
        
End Function



Sub InitBoard()
Dim X As Integer, Y As Integer, int2 As Integer, int1 As Integer

'For Int1 = 0 To 63
'  CurrentguiInt1).IsPiece = False  'blank
'Next Int1

int2 = 0

For int1 = 0 To 11

  int2 = ((int1 + 1) * 2) - 1
  If int1 > 3 And int1 < 8 Then int2 = int2 - 1
  
  If FullPieces = 1 Or int1 < 8 Then
    CurrentGUIBoard.Pieces(int1).Live = 1
  Else
    CurrentGUIBoard.Pieces(int1).Live = 0
  End If
  CurrentGUIBoard.Pieces(int1).Double = 0
  CurrentGUIBoard.Pieces(int1).player = 0
  CurrentGUIBoard.Pieces(int1).X = (int2 Mod 8) + 1
  CurrentGUIBoard.Pieces(int1).Y = ((int2 - (int2 Mod 8)) / 8) + 1

Next int1

int2 = 0
For int1 = 12 To 23



  int2 = ((int1 - 11) * 2) - 1
  If int1 >= 16 And int1 < 20 Then int2 = int2 - 1
  
  int2 = 63 - int2
  
  If FullPieces = 1 Or int1 < 20 Then
    CurrentGUIBoard.Pieces(int1).Live = 1
  Else
    CurrentGUIBoard.Pieces(int1).Live = 0
  End If
  
  CurrentGUIBoard.Pieces(int1).Double = 0
  CurrentGUIBoard.Pieces(int1).player = 1
  CurrentGUIBoard.Pieces(int1).X = (int2 Mod 8) + 1
  CurrentGUIBoard.Pieces(int1).Y = ((int2 - (int2 Mod 8)) / 8) + 1


Next int1

ReDim PastBoardStates(0 To 0)
CopyState PastBoardStates(0), CurrentGUIBoard

SquareSelect = 65
End Sub

Public Function IConvert(X As Byte, Y As Byte) As Byte
  If X > 8 Or X < 1 Or Y > 8 Or Y < 1 Then
    IConvert = 65
  Else
    IConvert = ((Y - 1) * 8) + X - 1
  End If
End Function

Public Sub XYConvert(Index As Byte, ByRef X As Byte, ByRef Y As Byte)
  If Index > 63 Or Index < 0 Then X = 0: Y = 0: Exit Sub
  Y = ((Index - (Index Mod 8)) / 8) + 1
  X = (Index Mod 8) + 1
End Sub


Public Function CheckWin(state As BoardState) As EndGameType
Dim pieceN As Byte, NPieces(0 To 1) As Byte

  For pieceN = 0 To 23
      If state.Pieces(pieceN).Live Then NPieces(state.Pieces(pieceN).player) = NPieces(state.Pieces(pieceN).player) + 1
  Next pieceN
  
  If NPieces(0) = 0 Then
    If NPieces(1) = 0 Then
      CheckWin = DRAWEND
    Else
      CheckWin = HUMANWIN
    End If
  ElseIf NPieces(1) = 0 Then
    CheckWin = COMPWIN
  Else
    CheckWin = NOEND
  End If
  
End Function

Public Function NotB(ByRef byte1 As Byte) As Byte
  If byte1 = 0 Then NotB = 1
End Function

Public Function ConvertPlayerN(player As Byte) As Byte
  If player = 0 Then ConvertPlayerN = 2 Else ConvertPlayerN = 1
End Function

Public Sub CopyState(ByRef statedes As BoardState, ByRef statefrom As BoardState)
Dim int1 As Byte
  
  For int1 = 0 To 23
    statedes.Pieces(int1) = statefrom.Pieces(int1)
  Next int1
  
'statedes.PlayerTurn = statefrom.PlayerTurn
End Sub
