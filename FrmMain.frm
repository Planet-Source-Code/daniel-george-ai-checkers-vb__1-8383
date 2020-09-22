VERSION 5.00
Begin VB.Form FrmMain 
   Caption         =   "Checkers - The A.I. Game"
   ClientHeight    =   8535
   ClientLeft      =   -2040
   ClientTop       =   450
   ClientWidth     =   12210
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   12210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Go Back"
      Height          =   615
      Left            =   9000
      TabIndex        =   87
      Top             =   4200
      Width           =   2415
   End
   Begin VB.CommandButton cmdNewGame 
      Caption         =   "Start New Game"
      Height          =   615
      Left            =   8520
      TabIndex        =   86
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Load or Save"
      Height          =   615
      Left            =   10320
      TabIndex        =   78
      Top             =   2280
      Width           =   1575
   End
   Begin VB.PictureBox PicWaiting 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   11280
      Picture         =   "FrmMain.frx":0442
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   77
      Top             =   7320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Index           =   6
      Left            =   8760
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   76
      Top             =   8760
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   12480
      Picture         =   "FrmMain.frx":0B84
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   75
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   11640
      Picture         =   "FrmMain.frx":3D8A
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   74
      Top             =   9000
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   9600
      Picture         =   "FrmMain.frx":6F90
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   73
      Top             =   9120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   10800
      Picture         =   "FrmMain.frx":A196
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   72
      Top             =   9120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   11040
      Picture         =   "FrmMain.frx":D39C
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   71
      Top             =   8640
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Index           =   0
      Left            =   10080
      ScaleHeight     =   1005
      ScaleWidth      =   1005
      TabIndex        =   70
      Top             =   8400
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   63
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   63
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   62
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   62
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   61
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   61
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   60
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   60
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   59
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   59
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   58
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   58
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   57
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   57
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   56
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   56
      Top             =   7200
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   55
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   55
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   54
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   54
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   53
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   53
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   52
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   52
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   51
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   51
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   50
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   50
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   49
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   49
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   48
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   48
      Top             =   6240
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   47
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   47
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   46
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   46
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   45
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   45
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   44
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   44
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   43
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   43
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   42
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   42
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   41
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   41
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   40
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   40
      Top             =   5280
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   39
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   39
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   38
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   38
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   37
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   37
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   36
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   36
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   35
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   35
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   34
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   34
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   33
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   33
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   32
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   32
      Top             =   4320
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   31
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   31
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   30
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   30
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   29
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   29
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   28
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   28
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   27
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   27
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   26
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   26
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   25
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   25
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   24
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   24
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   23
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   23
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   22
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   22
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   21
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   21
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   20
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   20
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   19
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   19
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   18
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   17
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   16
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   16
      Top             =   2400
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   15
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   15
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   14
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   14
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   13
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   12
      Left            =   4200
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   12
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   11
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   10
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   10
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   9
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   9
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   8
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   8
      Top             =   1440
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   7
      Left            =   7080
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   6
      Left            =   6120
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   6
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   5160
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   4
      Left            =   4200
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   3
      Left            =   3240
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   2
      Left            =   2280
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   1
      Left            =   1320
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Grid1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   0
      Left            =   360
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin VB.PictureBox Light 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   1
      Left            =   8280
      Picture         =   "FrmMain.frx":105A2
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   64
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Light 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   0
      Left            =   8280
      Picture         =   "FrmMain.frx":109E4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   66
      Top             =   1440
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox Light 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   2
      Left            =   8280
      Picture         =   "FrmMain.frx":10E26
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   67
      Top             =   1440
      Width           =   480
   End
   Begin VB.PictureBox Light 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Index           =   3
      Left            =   8280
      Picture         =   "FrmMain.frx":11268
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   65
      Top             =   480
      Width           =   480
   End
   Begin VB.Label lblXY 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   11400
      TabIndex        =   85
      Top             =   5500
      Width           =   450
   End
   Begin VB.Label lblPlayerScore 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   10920
      TabIndex        =   84
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblPlayerScore 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   10920
      TabIndex        =   83
      Top             =   480
      Width           =   855
   End
   Begin VB.Label lblpieceN 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8400
      TabIndex        =   82
      Top             =   5500
      Width           =   90
   End
   Begin VB.Label lblMaxDepth 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   81
      Top             =   8040
      Width           =   180
   End
   Begin VB.Label lblMoveRatio 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   80
      Top             =   7560
      Width           =   180
   End
   Begin VB.Label lblMoveNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8520
      TabIndex        =   79
      Top             =   7080
      Width           =   180
   End
   Begin VB.Line Line4 
      X1              =   8400
      X2              =   12000
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line2 
      Index           =   2
      X1              =   8280
      X2              =   15840
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   8280
      X2              =   12120
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label lblNames 
      AutoSize        =   -1  'True
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   8880
      TabIndex        =   69
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label lblNames 
      AutoSize        =   -1  'True
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   8880
      TabIndex        =   68
      Top             =   480
      Width           =   1590
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   8160
      X2              =   12120
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   240
      Y1              =   8400
      Y2              =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      Index           =   0
      X1              =   8160
      X2              =   8160
      Y1              =   8400
      Y2              =   240
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub cmdLoad_Click()
'Open "c:\windows\temp\checkers" & Text7 & ".txt" For Binary As #1
'  Get #1, , CurrentGUIBoard
'Close #1
'
'DrawBoard CurrentGUIBoard
'End Sub
'
'Private Sub cmdSave_Click()
'Open "c:\windows\temp\checkers" & Text7 & ".txt" For Binary As #1
'Put #1, , CurrentGUIBoard
'Close #1
'End Sub
'
'Private Sub Command1_Click()
'Depth.StartGame
'End Sub
'
'
'Private Sub Command2_Click()
'  SquareSelect = 65
'  Depth.DrawBoard CurrentGUIBoard
'  PieceMustMove = -1
'  ModConst.PlayerTurn = 0
'
'  ModConst.TurnTime(1) = GetTickCount()
'  ModConst.TurnTimeE(0) = False
'  ModConst.TurnTimeE(1) = False
'
'  'Sleep 50
'  Depth.AIMove
'End Sub
'
'Private Sub Command3_Click()
'Depth.Redraw CurrentGUIBoard
'End Sub
'
'Private Sub Command4_Click()
'Dim squares(0 To 8, 0 To 8) As Integer
'' -1=emptycomp   0=comp1  1=comp2 2=emptyhuman 3=human1   4=human2
'
'For PieceN = 1 To 16
'  temp1 = -1 + 2
'
'  If CurrentGUIBoard.Pieces(PieceN).Player = 1 Then temp1 = temp1 + 3
'  If CurrentGUIBoard.Pieces(PieceN).IsPiece Then temp1 = temp1 + 1
'  If CurrentGUIBoard.Pieces(PieceN).IsDouble Then temp1 = temp1 + 1
'
'  squares(CurrentGUIBoard.Pieces(PieceN).X, CurrentGUIBoard.Pieces(PieceN).Y) = temp1
'
'Next PieceN
'
'Open "C:\Program Files\Microsoft Visual Studio\MyProjects\Checker (C++)\input.txt" For Binary As #1
'  Put #1, , squares
'Close #1
'
'End Sub
'
'
'
'Private Sub Command6_Click()
'Depth.EvalCount
'End Sub
'
'Private Sub Command7_Click()
'Depth.GenCount
'End Sub
'
'Private Sub Form_Click()
'Dim fish As BoardState
'MsgBox Len(fish)
'End Sub
'
'Private Sub Form_Load()
'FrmMain.Width = 16000
'Text1_Change
'Text2_Change
'Text3_Change
'Text4_Change
'Text5_Change
'txtRatio_Change
'
'ModConst.Player1Name = "Bob"
'ModConst.Player1Mode = 1
'ModConst.Player2Name = "Daniel"
'ModConst.Player2Mode = 0
'
'For int1 = 0 To 63
'
'    If ((int1 - (int1 Mod 8)) / 8) Mod 2 = (int1 Mod 8) Mod 2 Then
'      FrmMain!Grid1(int1).BackColor = ColorConstants.vbBlack
'      FrmMain!Grid1(int1).Picture = FrmMain!Picture1(0).Picture
'    Else
'      FrmMain!Grid1(int1).BackColor = ColorConstants.vbWhite
'      FrmMain!Grid1(int1).Picture = FrmMain!Picture1(6).Picture
'    End If
'
'Next int1
'
'
'
'FrmMain!Label1(0) = Player1Name
'If Player1Mode = 0 Then FrmMain!HumanFace(0).Visible = True: FrmMain!Compface(0).Visible = False
'If Player1Mode = 1 Then FrmMain!HumanFace(0).Visible = False: FrmMain!Compface(0).Visible = True
'
'FrmMain!Label1(1) = Player2Name
'If Player2Mode = 0 Then FrmMain!HumanFace(1).Visible = True: FrmMain!Compface(1).Visible = False
'If Player2Mode = 1 Then FrmMain!HumanFace(1).Visible = False: FrmMain!Compface(1).Visible = True
'
'SetTurnLights 1
'
'Subs.PieceMustMove = -1
'Depth.MultiTake = True
'
'End Sub
'
'
'Private Sub Form_Unload(Cancel As Integer)
'FrmMain.Timer1.Enabled = False
'DoEvents
'End
'End Sub
'
'Private Sub GameTimer_Timer()
'Beep
'End Sub
'
''

Private Sub cmdNewGame_Click()
  frmNew.Show
End Sub

Private Sub cmdRedraw_Click()
  GUI.Redraw CurrentGUIBoard
End Sub

'
'
'Private Sub OpAutoDepth_Click()
'  Depth.AutoDepth = True
'End Sub
'
'Private Sub OpAutoDepth2_Click()
'  Depth.AutoDepth = False
'End Sub
'
'Private Sub OpExchange_Click()
'  Depth.Exchange = True
'End Sub
'
'Private Sub OpExchange2_Click()
'Depth.Exchange = False
'End Sub
'
'Private Sub opReversBoard_Click()
'Depth.DrawBoard CurrentGUIBoard
'End Sub
'
'
'Private Sub OpReverseBoard2_Click()
'Depth.DrawBoard CurrentGUIBoard
'End Sub
'
'Private Sub Option1_Click()
'Depth.MultiTake = False
'End Sub
'
'Private Sub Option2_Click()
'Depth.MultiTake = True
'End Sub
'
'Private Sub Text1_Change()
'
'If IsNumeric(Text1) Then
'  If Int(Text1) < 50 Then
'    MaxDepth = Int(Text1)
'  End If
'End If
'
'End Sub
'
'Private Sub Text2_Change()
'If IsNumeric(Text2) Then
'  CSinglePieceV = Int(Text2)
'End If
'End Sub
'
'Private Sub Text3_Change()
'If IsNumeric(Text3) Then
'  CDoublePieceV = Int(Text3)
'End If
'End Sub
'
'Private Sub Text5_Change()
'If IsNumeric(Text5) Then
'  MinTime = Int(Text5) * 1000
'End If
'End Sub
'Private Sub Text4_Change()
'If IsNumeric(Text4) Then
'  MaxTime = Int(Text4) * 1000
'End If
'End Sub
'
'Private Sub Text6_Change()
'If Text6 = "1" Then
'  PlayerTurn = 1
'ElseIf Text6 = "0" Then
'  PlayerTurn = 0
'End If
'
'If Not Text6 = "" Then Text6 = ""
'End Sub
'
'Sub Timer1_Timer()
'Dim Lng1 As Long, Lng2 As Long, Mins As Long, Secs As Long, MinsS As String, SecsS As String, int1 As Byte
'
'
'Timer1.Enabled = False
'Do While Timer1.Enabled = True
'
'  Lng1 = GetTickCount
'
'  If StartTimeE Then
'
'    Lng2 = Lng1 - StartTime
'
'    Mins = Round(((Lng2 - (Lng2 Mod 60000)) / 60000), 0)
'    Secs = Round((Lng2 Mod 60000) / 1000, 0)
'    If Secs = 60 Then Secs = 0: Mins = Mins + 1
'
'    SecsS = CStr(Secs)
'    If Len(SecsS) = 1 Then SecsS = "0" & SecsS
'
'    MinsS = CStr(Mins)
'
'    LblGameTime = MinsS & ":" & SecsS
'
'  End If
'  DoEvents
'
'  For int1 = 0 To 1
'    If TurnTimeE(int1) Then
'
'      Lng2 = Lng1 - TurnTime(int1)
'      Mins = Round(((Lng2 - (Lng2 Mod 60000)) / 60000), 0)
'      Secs = Round((Lng2 Mod 60000) / 1000, 0)
'      If Secs = 60 Then Secs = 0: Mins = Mins + 1
'
'      SecsS = CStr(Secs)
'      If Len(SecsS) = 1 Then SecsS = "0" & SecsS
'
'      MinsS = CStr(Mins)
'
'      Label2(int1) = MinsS & ":" & SecsS
'
'    End If
'  Next
'
'    DoEvents
'Loop
'
'End Sub
'
'Private Sub txtRatio_Change()
'If IsNumeric(txtRatio) Then
'  CRatioV = Int(txtRatio)
'End If
'End Sub
Private Sub cmdSave_Click()
frmGameSave.Show
End Sub




Private Sub Command1_Click()

Dim int1 As Byte


If Nplayers = 1 Then

  If UBound(PastBoardStates) <= 1 Then Exit Sub
  int1 = CurrentGUIBoard.PlayerTurn
  'int1 = NotB(int1)
  
  CopyState CurrentGUIBoard, PastBoardStates(UBound(PastBoardStates) - 1)
  ReDim Preserve PastBoardStates(UBound(PastBoardStates) - 2)
  CurrentGUIBoard.PlayerTurn = int1
  
  
  
  GUI.DrawBoard CurrentGUIBoard
  
  If Nplayers = 1 And CurrentPlayerTurn = COMP Then AI.MakeAIMove CurrentGUIBoard
  
Else
  
  If UBound(PastBoardStates) <= 0 Then Exit Sub
  
  int1 = CurrentGUIBoard.PlayerTurn
  int1 = NotB(int1)
  
  CopyState CurrentGUIBoard, PastBoardStates(UBound(PastBoardStates))
  ReDim Preserve PastBoardStates(UBound(PastBoardStates) - 1)
  CurrentGUIBoard.PlayerTurn = int1
  
  
  
  GUI.DrawBoard CurrentGUIBoard
  
  If Nplayers = 1 And CurrentPlayerTurn = COMP Then AI.MakeAIMove CurrentGUIBoard
  
End If
End Sub

Private Sub cmdStart1P_Click()
Nplayers = 1
CurrentPlayerTurn = 1
General.InitBoard
GUI.DrawBoard CurrentGUIBoard

End Sub

Private Sub cmdStart2P_Click()
Nplayers = 2
CurrentPlayerTurn = 1
General.InitBoard
GUI.DrawBoard CurrentGUIBoard
'GUI.DrawBoard CurrentGUIBoard
End Sub

Private Sub Command3_Click()
GUI.DrawBoard CurrentGUIBoard
End Sub

Private Sub Command4_Click()
SquareSelect = 65
GUI.DrawBoard CurrentGUIBoard
End Sub



Private Sub Command2_Click()
  AI.MakeAIMove CurrentGUIBoard
  CurrentPlayerTurn = HUMAN
End Sub

Private Sub Form_Load()
LoadMoveValues
'oppForceTake(1).Value = True
'oppForceTake_Click 1

'oppMultiTake(1).Value = True
'oppMultiTake_Click 1

'oppFullset(1).Value = True
'oppFullset_Click 1

PlayerNames(0) = "Player 2"
PlayerNames(1) = "Player 1"
SetPlayerLights
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblXY = ""
lblpieceN = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmGameSave
End Sub

Private Sub Grid1_Click(Index As Integer)
Dim Takepiece As Boolean, NewPos As Integer
'If FrmMain!opReversBoard Then Index = 63 - Index

Dim Square1 As CPiece, Square2 As CPiece, Square3 As CPiece
Dim X As Byte, Y As Byte
Dim cX As Integer, cY As Integer
Dim Moves() As move
Dim index1 As Byte

index1 = CByte(Index)
If ReverseBoard = 1 Then index1 = 63 - index1
'SquareSelect = 65


    
  If SquareSelect = 65 Then
    General.XYConvert index1, X, Y
    Square1 = CheckSquare(X, Y, CurrentGUIBoard)

    If Square1.isSquare = 0 Then Exit Sub
    If Square1.Live = 0 Then Exit Sub

    'If Square1.Player = COMP Then Exit Sub 'TODO
    
    SquareSelect = index1
    
    'GUI.DrawBoard CurrentGUIBoard
    GUI.DrawBoard CurrentGUIBoard
    Exit Sub
  End If

  If SquareSelect = index1 Then GoTo exitMove

  General.XYConvert SquareSelect, X, Y
  Square1 = CheckSquare(X, Y, CurrentGUIBoard)
  
  If Not Square1.player = CurrentPlayerTurn Then GoTo exitMove
  
  General.XYConvert index1, X, Y
  Square2 = CheckSquare(X, Y, CurrentGUIBoard)
  
  If Square2.isSquare = 0 Then GoTo exitMove

Dim PossTakes As Posstakemoves

'MsgBox Square1.Player
PossTakes = AI.FindPossTakes(CurrentGUIBoard, Square1.player)

  
  If Square2.Live = 0 Then
    If ForceTake = 1 And PossTakes.nPoss > 0 Then GoTo exitMove
  
    Dim int1 As Integer
    
    
    
    cX = CInt(Square2.X) - CInt(Square1.X)
    
    
    cY = CInt(Square2.Y) - CInt(Square1.Y)
    
    If Not (Abs(cX) = 1 And Abs(cY) = 1) Then GoTo exitMove
    
      If Not Square1.Double = 1 Then
        If Square1.player = COMP Then
          If cY = -1 Then GoTo exitMove
        Else
          If cY = 1 Then GoTo exitMove
        End If
      End If
      
    
    ReDim Moves(1)
    Moves(1).KillIndex = 65
    Moves(1).MoveFromIndex = SquareSelect
    Moves(1).MoveToIndex = index1
    Moves(1).Delay = 10
        
        If Square1.player = COMP Then
          If Square2.Y = 8 Then CurrentGUIBoard.Pieces(Square1.pNumber).Double = 1
        Else
          If Square2.Y = 1 Then CurrentGUIBoard.Pieces(Square1.pNumber).Double = 1
        End If
    
    
    CurrentPlayerTurn = NotB(CurrentPlayerTurn)
    
    GUI.MakeMove Moves, CurrentGUIBoard
    
    'normal move
  Else
    If Square2.player = Square1.player Then SquareSelect = index1: GUI.DrawBoard CurrentGUIBoard: Exit Sub
    
    cX = (CInt(Square1.X) - CInt(Square2.X))
    cY = (CInt(Square1.Y) - CInt(Square2.Y))
    If Not (Abs(cX) = 1 And Abs(cY) = 1) Then SquareSelect = 65: GUI.DrawBoard CurrentGUIBoard: Exit Sub
    
      If Not Square1.Double = 1 Then
        If Square1.player = COMP Then
          If cY = 1 Then GoTo exitMove
        Else
          If cY = -1 Then GoTo exitMove
        End If
      End If
      
      X = (CInt(Square2.X) - CInt(Square1.X)) + Square2.X
      Y = (CInt(Square2.Y) - CInt(Square1.Y)) + Square2.Y
      
      'square behind enemy
      Square3 = CheckSquare(X, Y, CurrentGUIBoard)
      If Square3.isSquare = 0 Or Square3.Live = 1 Then Exit Sub
      
      ReDim Moves(1)
      Moves(1).KillIndex = index1
      Moves(1).MoveFromIndex = SquareSelect
      Moves(1).MoveToIndex = General.IConvert(X, Y)
      Moves(1).Delay = 10
      
      If Square1.player = COMP Then
          If Square3.Y = 8 Then CurrentGUIBoard.Pieces(Square1.pNumber).Double = 1
        Else
          If Square3.Y = 1 Then CurrentGUIBoard.Pieces(Square1.pNumber).Double = 1
        End If
      
      GUI.MakeMove Moves, CurrentGUIBoard
      
      PossTakes = AI.FindPossTakes(CurrentGUIBoard, Square1.player)
      If PossTakes.nPoss = 0 Then
        CurrentPlayerTurn = NotB(CurrentPlayerTurn)
      End If
      
      GUI.DrawBoard CurrentGUIBoard
  End If

  DoEvents

If Nplayers = 1 And CurrentPlayerTurn = COMP Then AI.MakeAIMove CurrentGUIBoard

Exit Sub

exitMove:
  
  SquareSelect = 65
  GUI.DrawBoard CurrentGUIBoard

End Sub

Private Sub Grid1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Byte, y1 As Byte
Dim Square1 As CPiece
Static last As Integer

If last = Index And Not lblXY = "" Then Exit Sub
last = Index

General.XYConvert CByte(Index), x1, y1
Square1 = CheckSquare(x1, y1, CurrentGUIBoard)

lblXY = "X = " & x1 & vbNewLine & "Y = " & y1

If Square1.Live = 0 Then
  lblpieceN = ""
  Exit Sub
End If

lblpieceN = "Player " & ConvertPlayerN(Square1.player) & vbNewLine
If Square1.Double = 1 Then lblpieceN = lblpieceN & "Double piece" Else lblpieceN = lblpieceN & "Single Piece"

End Sub




Private Sub lblNames_Click(Index As Integer)
  frmSetName.Show
End Sub

Private Sub Light_Click(Index As Integer)
  If Index > 1 Then Index = Index - 2
  CurrentPlayerTurn = Index
  GUI.DrawBoard CurrentGUIBoard
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

Private Sub opReversBoard_Click(Index As Integer)
  Variables.ReverseBoard = Index
  SquareSelect = 65
  GUI.Redraw CurrentGUIBoard
End Sub



