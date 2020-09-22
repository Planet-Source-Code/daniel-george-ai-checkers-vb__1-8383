VERSION 5.00
Begin VB.Form frmGameSave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Load or Save Game"
   ClientHeight    =   5160
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10380
   Icon            =   "frmGameSave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   10380
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Index           =   0
      Left            =   8280
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   80
      Top             =   4080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   8280
      Picture         =   "frmGameSave.frx":000C
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   79
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   8280
      Picture         =   "frmGameSave.frx":3212
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   78
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   8280
      Picture         =   "frmGameSave.frx":6418
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   77
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   8280
      Picture         =   "frmGameSave.frx":961E
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   76
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   8280
      Picture         =   "frmGameSave.frx":C824
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   75
      Top             =   4080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Index           =   6
      Left            =   8400
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   74
      Top             =   4080
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   0
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   71
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   1
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   70
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   2
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   69
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   3
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   68
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   4
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   67
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   5
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   66
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   6
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   65
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   7
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   64
      Top             =   720
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   8
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   63
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   9
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   62
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   10
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   61
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   11
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   60
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   12
      Left            =   8520
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   59
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   13
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   58
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   14
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   57
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   15
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   56
      Top             =   1080
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   16
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   55
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   17
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   54
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   18
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   53
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   19
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   52
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   20
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   51
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   21
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   50
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   22
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   49
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   23
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   48
      Top             =   1440
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   24
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   47
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   25
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   46
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   26
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   45
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   27
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   44
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   28
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   43
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   29
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   42
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   30
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   41
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   31
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   40
      Top             =   1800
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   32
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   39
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   33
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   38
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   34
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   37
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   35
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   36
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   36
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   35
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   37
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   34
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   38
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   33
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   39
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   32
      Top             =   2160
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   40
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   31
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   41
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   30
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   42
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   29
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   43
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   28
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   44
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   27
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   45
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   26
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   46
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   25
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   47
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   24
      Top             =   2520
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   48
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   23
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   49
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   22
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   50
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   21
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   51
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   20
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   52
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   19
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   53
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   18
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   54
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   17
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   55
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   16
      Top             =   2880
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   56
      Left            =   7080
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   15
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   57
      Left            =   7440
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   14
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   58
      Left            =   7800
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   13
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   59
      Left            =   8160
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   12
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   60
      Left            =   8520
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   11
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   61
      Left            =   8880
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   10
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   62
      Left            =   9240
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   9
      Top             =   3240
      Width           =   400
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   400
      Index           =   63
      Left            =   9600
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   8
      Top             =   3240
      Width           =   400
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete Game"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txtSaveName 
      Height          =   375
      Left            =   240
      MaxLength       =   100
      TabIndex        =   3
      Top             =   4560
      Width           =   4575
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Game"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Game"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox lstDescriptions 
      Height          =   2595
      ItemData        =   "frmGameSave.frx":FA2A
      Left            =   240
      List            =   "frmGameSave.frx":FA31
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label lblPlayerN 
      AutoSize        =   -1  'True
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   9120
      TabIndex        =   73
      Top             =   3840
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Turn:     Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7800
      TabIndex        =   72
      Top             =   3840
      Width           =   1260
   End
   Begin VB.Line Line2 
      X1              =   6720
      X2              =   6720
      Y1              =   600
      Y2              =   5040
   End
   Begin VB.Label lblSavePath 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6480
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Load description"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Save description"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   4200
      Width           =   1185
   End
End
Attribute VB_Name = "frmGameSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim SavePaths() As String 'all saved file paths
Dim SaveDescriptions() As String 'all saved file descriptions



Private Sub cmdLoad_Click()
Dim SaveInfo As FileSaveStructure

If Not FileLen(SavePaths(lstDescriptions.ListIndex + 1)) = Len(SaveInfo) Then
  MsgBox "Save corrupted"
  Exit Sub
End If

Open SavePaths(lstDescriptions.ListIndex + 1) For Binary As #1
  Get #1, , SaveInfo
Close #1

CurrentGUIBoard = SaveInfo.state
CurrentPlayerTurn = SaveInfo.playerturn

GUI.DrawBoard CurrentGUIBoard

frmGameSave.Hide
End Sub
'
Private Sub cmdSave_Click()
Dim savePath As String, int1 As Integer, description As String * 100
Dim SaveInfo As FileSaveStructure

savePath = General.SlashDirectory(App.Path) & "Saves\Save"

Do While Not (Dir(savePath & int1 & ".aic") = "")
  int1 = int1 + 1
Loop

SaveInfo.description = txtSaveName
SaveInfo.state = CurrentGUIBoard
SaveInfo.playerturn = CurrentPlayerTurn

  Open savePath & int1 & ".aic" For Binary As #1
    Put #1, , SaveInfo
  Close #1
  
  txtSaveName = ""
  RefreshSavedGames
  
  frmGameSave.Hide
  
End Sub

Private Sub cmdDelete_Click()

If SavePaths(lstDescriptions.ListIndex + 1) = "" Then Exit Sub
If Dir(SavePaths(lstDescriptions.ListIndex + 1)) = "" Then MsgBox "Error file not found": RefreshSavedGames: Exit Sub

Dim YesNo As VbMsgBoxResult
YesNo = MsgBox(SavePaths(lstDescriptions.ListIndex + 1), vbYesNo + vbExclamation, "Are you sure you want to delete this file?")
If YesNo = vbYes Then
  Kill SavePaths(lstDescriptions.ListIndex + 1)
  'MsgBox SavePaths(lstDescriptions.ListIndex + 1), , "File deleted"
End If

RefreshSavedGames
End Sub

Private Sub Form_Load()
  RefreshSavedGames
End Sub

Private Sub lstDescriptions_Click()
lblSavePath = Right(SavePaths(lstDescriptions.ListIndex + 1), Len(SavePaths(lstDescriptions.ListIndex + 1)) - Len(SlashDirectory(App.Path)))
DisplayPreview
End Sub

Private Sub DisplayPreview()

Dim SaveInfo As FileSaveStructure

If Not FileLen(SavePaths(lstDescriptions.ListIndex + 1)) = Len(SaveInfo) Then
  Exit Sub
End If

Open SavePaths(lstDescriptions.ListIndex + 1) For Binary As #1
  Get #1, , SaveInfo
Close #1

'CurrentGUIBoard = SaveInfo.state
'CurrentPlayerTurn = SaveInfo.playerturn

'GUI.DrawBoard CurrentGUIBoard
If SaveInfo.playerturn = COMP Then
  lblPlayerN = "2"
Else
  lblPlayerN = "1"
End If


Dim Tempboard As BoardState
Tempboard = SaveInfo.state

Dim int1 As Long, int2 As Long

  For int1 = 0 To 63
    
      If ((int1 - (int1 Mod 8)) / 8) Mod 2 <> (int1 Mod 8) Mod 2 Then
        'If Not Grid1(int1).BackColor = ColorConstants.vbBlack Then
          Grid1(int1).BackColor = ColorConstants.vbBlack
          'Grid1(int1).Picture = Picture1(0).Picture
        'End If
      Else
        'If Not Grid1(int1).BackColor = ColorConstants.vbWhite Then
          Grid1(int1).BackColor = ColorConstants.vbWhite
          
          'Grid1(int1).Picture = Picture1(0).Picture
        'End If
      End If
    
  Next
  
  For int2 = 0 To 23
    If Tempboard.Pieces(int2).Live = 0 Then GoTo 1
    int1 = General.IConvert(Tempboard.Pieces(int2).x, Tempboard.Pieces(int2).y)

    If int1 < 0 Then GoTo 1
    
      Select Case Tempboard.Pieces(int2).Player
        Case Is = 0
          If Tempboard.Pieces(int2).Double = 0 Then Grid1(int1).BackColor = ColorConstants.vbBlue ': Grid1(int1).Picture = Picture1(3).Picture
          If Tempboard.Pieces(int2).Double = 1 Then Grid1(int1).BackColor = ColorConstants.vbCyan ': Grid1(int1).Picture = Picture1(2).Picture
        Case Is = 1
          If Tempboard.Pieces(int2).Double = 0 Then Grid1(int1).BackColor = ColorConstants.vbRed ': Grid1(int1).Picture = Picture1(4).Picture
          If Tempboard.Pieces(int2).Double = 1 Then Grid1(int1).BackColor = ColorConstants.vbMagenta ': Grid1(int1).Picture = Picture1(5).Picture
      End Select
      
1
  Next


'If Not SquareSelect = 65 Then Grid1(SquareSelect).BackColor = ColorConstants.vbYellow: Grid1(SquareSelect).Picture = Picture1(1).Picture

End Sub

Private Sub txtSaveName_Change()
  cmdSave.Enabled = Not (txtSaveName = "")
End Sub

Private Sub RefreshSavedGames()

Dim int1 As Integer
Dim File1 As String 'temporary
Dim SaveInfo As FileSaveStructure

ReDim SavePaths(0)
ReDim SaveDescriptions(0)

File1 = Dir(SlashDirectory(App.Path) & "saves\*.aic")

Do While Not File1 = ""
  ReDim Preserve SavePaths(UBound(SavePaths) + 1)
  ReDim Preserve SaveDescriptions(UBound(SaveDescriptions) + 1)
  SavePaths(UBound(SavePaths)) = SlashDirectory(App.Path) & "saves\" & File1
  
  If FileLen(SavePaths(UBound(SavePaths))) = Len(SaveInfo) Then
  
    Open SavePaths(UBound(SavePaths)) For Binary Access Read As #1 'read the header(100chars)
      Get #1, , SaveInfo
    Close #1
    
    SaveDescriptions(UBound(SaveDescriptions)) = SaveInfo.description
  Else
  
    SaveDescriptions(UBound(SaveDescriptions)) = "Error - " & SavePaths(UBound(SavePaths))
    
  End If

  
  File1 = Dir
Loop

lstDescriptions.Clear
For int1 = 1 To UBound(SaveDescriptions) 'add all to list box
  lstDescriptions.AddItem SaveDescriptions(int1)
Next

End Sub
