VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create new game"
   ClientHeight    =   5580
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Game Rules"
      Height          =   4215
      Left            =   3960
      TabIndex        =   16
      Top             =   240
      Width           =   3375
      Begin VB.Frame Frame3 
         Caption         =   "Full piece set (24 instead of 16)"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   26
         Top             =   3360
         Width           =   2895
         Begin VB.OptionButton oppFullset 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   28
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton oppFullset 
            Caption         =   "Yes"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Force Taking"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   2895
         Begin VB.OptionButton oppForceTake 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   25
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton oppForceTake 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   24
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Multi-piece taking"
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   20
         Top             =   1440
         Width           =   2895
         Begin VB.OptionButton oppMultiTake 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton oppMultiTake 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Reverse Board"
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   2895
         Begin VB.OptionButton opReversBoard 
            Caption         =   "Yes"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton opReversBoard 
            Caption         =   "No"
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   18
            Top             =   240
            Value           =   -1  'True
            Width           =   615
         End
      End
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Start Game"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   15
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5880
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "Player 2"
      Height          =   3135
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   3615
      Begin VB.TextBox txtmaxTime 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1905
         TabIndex        =   11
         Text            =   "10.00"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox txtMaxDepth 
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Text            =   "4"
         Top             =   2040
         Width           =   975
      End
      Begin VB.OptionButton oppPlayerType2 
         Caption         =   "Human"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   9
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton oppPlayerType2 
         Caption         =   "Computer"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   8
         Top             =   1320
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtUserName 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   4
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum AI Time"
         Height          =   195
         Index           =   3
         Left            =   555
         TabIndex        =   13
         Top             =   2400
         Width           =   1245
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Maximum AI Depth"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Name"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Player 1"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.OptionButton oppPlayerType1 
         Caption         =   "Human"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton oppPlayerType1 
         Caption         =   "Computer"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtUserName 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   2445
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "&Name"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
  Unload Me
End Sub

Private Sub OKButton_Click()
If oppPlayerType2(0).Value = True Then 'comp

  Nplayers = 1
  CurrentPlayerTurn = 1
  txtMaxDepth_Change
  txtmaxTime_Change
  General.InitBoard
  GUI.DrawBoard CurrentGUIBoard

Else
  
  Nplayers = 2
  CurrentPlayerTurn = 1
  txtMaxDepth_Change
  txtmaxTime_Change
  General.InitBoard
  GUI.DrawBoard CurrentGUIBoard

End If
Unload Me
End Sub

Private Sub oppForceTake_Click(Index As Integer)
  Variables.ForceTake = Index
End Sub

Private Sub oppFullset_Click(Index As Integer)
FullPieces = Index
End Sub

Private Sub oppMultiTake_Click(Index As Integer)
  Variables.MultiTake = Index
End Sub

Private Sub oppPlayerType2_Click(Index As Integer)
If oppPlayerType2(1).Value = True Then
  txtMaxDepth.Enabled = False
  txtmaxTime.Enabled = False
Else
  txtMaxDepth.Enabled = True
  txtmaxTime.Enabled = True
End If
End Sub

Private Sub opReversBoard_Click(Index As Integer)
  Variables.ReverseBoard = Index
  SquareSelect = 65
  GUI.Redraw CurrentGUIBoard
End Sub



Private Sub cmdOK_Click()
  PlayerNames(0) = txtUserName(0)
  PlayerNames(1) = txtUserName(1)
  Unload Me
End Sub

Private Sub Form_Load()
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
  oppFullset(1) = True
  oppFullset_Click 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
End Sub


Private Sub txtMaxDepth_Change()
If IsNumeric(txtMaxDepth) Then AI.MaxDepth = txtMaxDepth
End Sub

Private Sub txtmaxTime_Change()
If IsNumeric(txtmaxTime) Then AI.MaxAiTime = txtmaxTime * 1000
End Sub
