VERSION 5.00
Begin VB.Form frmSetName 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Names"
   ClientHeight    =   1935
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1143.262
   ScaleMode       =   0  'User
   ScaleWidth      =   4873.129
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   3405
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   135
      Width           =   3405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   960
      TabIndex        =   2
      Top             =   1320
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2700
      TabIndex        =   3
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label lblLabels 
      Caption         =   "Player &2 Name"
      Height          =   270
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   735
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "Player &1 Name"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmSetName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdCancel_Click()
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
  Unload Me
End Sub

Private Sub cmdOK_Click()
  PlayerNames(0) = txtUserName(0)
  PlayerNames(1) = txtUserName(1)
  SetPlayerLights
  Unload Me
End Sub

Private Sub Form_Load()
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  txtUserName(0) = PlayerNames(0)
  txtUserName(1) = PlayerNames(1)
End Sub
