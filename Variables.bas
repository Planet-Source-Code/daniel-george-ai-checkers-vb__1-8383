Attribute VB_Name = "Variables"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount Lib "kernel32" () As Long

'--------------- Types -----------------

Public Type Piece
  X As Byte
  Y As Byte
  player As Byte
  Double As Byte
  Live As Byte
End Type

Public Type CPiece
  X As Byte
  Y As Byte
  Index As Byte
  player As Byte
  Double As Byte
  Live As Byte
  pNumber As Byte
  isSquare As Byte
End Type

Public Type BoardState
  Pieces(23) As Piece
  Depth As Byte
  ParentID As Long
  Evaluation As Integer
  PlayerTurn As Byte
End Type

Public Type move
  KillIndex As Byte
  MoveFromIndex As Byte
  MoveToIndex As Byte
  Delay As Long
End Type

Public Type PossTakeMove
  MoveFrom As Byte 'origin square
  MoveTo() As Byte 'possible take pieces
  nMoveTo As Byte 'size of array
End Type

Public Type Posstakemoves
  poss() As PossTakeMove  'possible moves to take a piece
  nPoss As Byte 'size of array
End Type

Public Type FileSaveStructure
  description As String * 100
  state As BoardState
  PlayerTurn As Byte
End Type

'--------- Game wide variables ------

Global Const COMP As Byte = 0
Global Const HUMAN As Byte = 1

Global CurrentPlayerTurn As Byte
Global CurrentTurnNumber As Integer

Global PastBoardStates() As BoardState

Global CurrentGUIBoard As BoardState  'current displayed board

Global SquareSelect As Byte

Global Nplayers As Byte

Global ForceTake As Byte
Global MultiTake As Byte
Global ReverseBoard As Byte
Global FullPieces As Byte

Global PlayerNames(0 To 1) As String

Enum EndGameType
  COMPWIN
  HUMANWIN
  DRAWEND
  NOEND
End Enum
