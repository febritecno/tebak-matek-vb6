VERSION 5.00
Begin VB.Form frmHMMain 
   BackColor       =   &H00004080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tebak Matek"
   ClientHeight    =   5580
   ClientLeft      =   1812
   ClientTop       =   1860
   ClientWidth     =   9156
   Icon            =   "HMMain.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   9156
   Begin VB.CommandButton cmdExit 
      Caption         =   "K&eluar"
      Height          =   495
      Left            =   4140
      TabIndex        =   57
      Top             =   4740
      Width           =   1515
   End
   Begin VB.Timer tmrWinLoseTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6000
      Top             =   4800
   End
   Begin VB.CommandButton cmdNewGame 
      BackColor       =   &H008080FF&
      Caption         =   "&Main Lagi"
      Height          =   495
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "Langsung &Jawab !!"
      Height          =   495
      Left            =   360
      TabIndex        =   50
      Top             =   4740
      Width           =   1515
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   26
      Left            =   5220
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   3900
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   25
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   24
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   3900
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   23
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   22
      Left            =   2580
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   3900
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   21
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   3900
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   3180
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4260
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   300
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2580
      Width           =   615
   End
   Begin VB.CommandButton cmdLetterPool 
      BackColor       =   &H0000C0C0&
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblYouWinLose 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SAMPEAN MENANG!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   495
      Left            =   180
      TabIndex        =   56
      Top             =   1560
      Width           =   6615
   End
   Begin VB.Label lblGamesWon 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   8100
      TabIndex        =   55
      Top             =   4980
      Width           =   855
   End
   Begin VB.Label lblGamesPlayed 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   8100
      TabIndex        =   54
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Menang :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   6480
      TabIndex        =   53
      Top             =   4980
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Jumlah Main:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   6480
      TabIndex        =   52
      Top             =   4620
      Width           =   1575
   End
   Begin VB.Label lblClue 
      BackStyle       =   0  'Transparent
      Caption         =   "Clue : KECAMATAN"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   120
      TabIndex        =   49
      Top             =   60
      Width           =   4635
   End
   Begin VB.Label lblPick 
      BackStyle       =   0  'Transparent
      Caption         =   "Ketik Kata"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   315
      Left            =   180
      TabIndex        =   48
      Top             =   2100
      Width           =   1815
   End
   Begin VB.Line linGallows3 
      BorderWidth     =   3
      X1              =   8580
      X2              =   8580
      Y1              =   600
      Y2              =   3960
   End
   Begin VB.Line linGallows2 
      BorderWidth     =   3
      X1              =   8580
      X2              =   7620
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line linGallowsBase 
      BorderWidth     =   3
      X1              =   8940
      X2              =   8220
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linGallows1 
      BorderWidth     =   3
      X1              =   7620
      X2              =   7620
      Y1              =   600
      Y2              =   960
   End
   Begin VB.Shape shpHead 
      BorderWidth     =   3
      Height          =   735
      Left            =   6900
      Shape           =   3  'Circle
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Line linTorso 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7620
      X2              =   7620
      Y1              =   1680
      Y2              =   2640
   End
   Begin VB.Line linArm1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7620
      X2              =   6900
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line linArm2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7620
      X2              =   8340
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line linLeg1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   7620
      X2              =   7140
      Y1              =   2640
      Y2              =   3480
   End
   Begin VB.Line linLeg2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   7620
      X2              =   8100
      Y1              =   2640
      Y2              =   3480
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   20
      Left            =   6060
      TabIndex        =   20
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   19
      Left            =   5400
      TabIndex        =   19
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   18
      Left            =   4740
      TabIndex        =   18
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   17
      Left            =   4080
      TabIndex        =   17
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   16
      Left            =   3420
      TabIndex        =   16
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   15
      Left            =   2760
      TabIndex        =   15
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   14
      Left            =   2100
      TabIndex        =   14
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   13
      Left            =   1440
      TabIndex        =   13
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   12
      Left            =   780
      TabIndex        =   12
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   900
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   6060
      TabIndex        =   10
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   9
      Left            =   5400
      TabIndex        =   9
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   8
      Left            =   4740
      TabIndex        =   8
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   7
      Left            =   4080
      TabIndex        =   7
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   6
      Left            =   3420
      TabIndex        =   6
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   5
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   2100
      TabIndex        =   4
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   1440
      TabIndex        =   3
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   780
      TabIndex        =   2
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   675
   End
   Begin VB.Label lblPuzzleLetter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   5640
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.Menu Cmd_File 
      Caption         =   "&File"
      Begin VB.Menu Cmd_Back 
         Caption         =   "&Warna Backround"
         Begin VB.Menu Cmd_Hijau 
            Caption         =   "&Hijau"
         End
         Begin VB.Menu Cmd_Biru 
            Caption         =   "&Biru"
         End
         Begin VB.Menu Cmd_Merah 
            Caption         =   "&Merah"
         End
         Begin VB.Menu Cmd_Kuning 
            Caption         =   "&Kuning"
         End
      End
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu Cmd_About 
         Caption         =   "&About Me"
      End
      Begin VB.Menu Cmd_Close 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu Cmd_Main 
      Caption         =   "&Cara Main"
   End
End
Attribute VB_Name = "frmHMMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'**********************************************************************
'*                     Level Variables                            *
'**********************************************************************

    Private mstrLabelLetter(1 To 20)    As String * 1
    Private mstrPuzzle                  As String
    Private mintErrorCount              As Integer
    Private mintGamesPlayed             As Integer
    Private mintGamesWon                As Integer

    Private mintNbrOfPuzzles           As Integer
    Private mintNbrOfClues              As Integer
    Private mintLastPuzzle             As Integer

    Private Type HangmanPuzzleRecord
        Puzzle                          As String
        Clue                            As Integer
    End Type
    
    Private Type ClueCodeRecord
        ClueCode                        As Integer
        ClueDescription                 As String
    End Type
    
    Private maudtPuzzleRec()            As HangmanPuzzleRecord
    Private maudtClueRec()              As ClueCodeRecord


Private Sub Cmd_Biru_Click()
frmHMMain.BackColor = vbBlue
End Sub

Private Sub Cmd_Close_Click()
End
End Sub

Private Sub Cmd_Hijau_Click()
frmHMMain.BackColor = vbGreen
End Sub


Private Sub Cmd_Kuning_Click()
frmHMMain.BackColor = vbYellow
End Sub

Private Sub Cmd_Main_Click()
frmCRMain.Show
End Sub

Private Sub Cmd_Merah_Click()
frmHMMain.BackColor = vbRed
End Sub


'**********************************************************************
'*                       Event PROSEDUR                          *
'**********************************************************************

'----------------------------------------------------------------------
Private Sub Form_Load()
'----------------------------------------------------------------------

    LoadGameData
    
    CenterForm Me
    
    cmdNewGame_Click

End Sub


'----------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'----------------------------------------------------------------------

    If MsgBox("Mau Keluar Ya...? Tekan Yes", _
              vbYesNo + vbQuestion, _
              "Keluar") = vbNo Then
        Cancel = 1
    Else
        SaveLastPuzzleNumber
    End If

End Sub

'**********************************************************************
'*                 Command Button Event PROSEDUR                 *
'**********************************************************************

'----------------------------------------------------------------------
Private Sub cmdNewGame_Click()
'----------------------------------------------------------------------
If MsgBox("Mulai Main.....??", vbYesNo + vbQuestion, "Konfirmasi") = vbYes Then

    Dim blnPuzzleOK     As Boolean
    Dim intX            As Integer

    mintErrorCount = 0
    
    SetUpFormForNewGame
    
    Do
        GetNextPuzzle
        blnPuzzleOK = LoadPuzzleIntoInterface()
        If Not blnPuzzleOK Then
            If MsgBox("Puzzle '" & Trim$(mstrPuzzle) & "' does not fit on the " _
                    & "gameboard and/or contains a non-alphabetic character. " _
                    & "Try the next available puzzle?", _
                      vbYesNo, _
                      "Bad Puzzle") = vbNo Then
                SaveLastPuzzleNumber
                End
            End If
        End If
    Loop Until blnPuzzleOK

    For intX = 1 To 20
        lblPuzzleLetter(intX).Visible _
            = IIf(mstrLabelLetter(intX) = " ", False, True)
    Next
    Else
    Unload Me
End If

End Sub


'----------------------------------------------------------------------
Private Sub cmdSolve_Click()
'----------------------------------------------------------------------

    Dim strUserGuess    As String
    Dim intX            As Integer

    strUserGuess _
        = InputBox$("Angger wis roh jawabane ketik nang ngisor iki !!", _
                    "Langsung Jawab !!")

    If strUserGuess = "" Then Exit Sub

    If UCase$(Trim$(strUserGuess)) = Trim$(mstrPuzzle) Then
        lblYouWinLose.Caption = "SAMPEAN MENANG!!"
        mintGamesWon = mintGamesWon + 1
    Else
        shpHead.Visible = True
        linTorso.Visible = True
        linArm1.Visible = True
        linArm2.Visible = True
        linLeg1.Visible = True
        linLeg2.Visible = True
        lblYouWinLose.Caption = "SAMPEAN MATEK"
    End If

    mintGamesPlayed = mintGamesPlayed + 1
    lblGamesPlayed.Caption = CStr(mintGamesPlayed)
    lblGamesWon.Caption = CStr(mintGamesWon)
    lblYouWinLose.Visible = True
    tmrWinLoseTimer.Enabled = True
    For intX = 1 To 26
        cmdLetterPool(intX).Enabled = False
    Next
    For intX = 1 To 20
        If mstrLabelLetter(intX) <> " " Then
            lblPuzzleLetter(intX).Caption = mstrLabelLetter(intX)
        End If
        lblPuzzleLetter(intX).Enabled = False
    Next
    lblClue.Enabled = False
    lblPick.Enabled = False
    cmdSolve.Enabled = False

End Sub


'----------------------------------------------------------------------
Private Sub cmdLetterPool_Click(Index As Integer)
'----------------------------------------------------------------------

    Dim intX            As Integer
    Dim blnCorrectGuess As Boolean
    Dim blnPlayerWins   As Boolean
    Dim blnGameOver     As Boolean

    cmdLetterPool(Index).Enabled = False

    For intX = 1 To 20
        If mstrLabelLetter(intX) = cmdLetterPool(Index).Caption Then
            lblPuzzleLetter(intX).Caption = cmdLetterPool(Index).Caption
            blnCorrectGuess = True
        End If
    Next

    If blnCorrectGuess Then
        blnPlayerWins = True
        For intX = 1 To 20
            If lblPuzzleLetter(intX).Visible Then
                If lblPuzzleLetter(intX).Caption = "" Then
                    blnPlayerWins = False
                    Exit For
                End If
            End If
        Next
        If blnPlayerWins Then
            lblYouWinLose.Caption = "SAMPEAN MENANG!!"
            mintGamesWon = mintGamesWon + 1
            blnGameOver = True
        End If
    Else
        mintErrorCount = mintErrorCount + 1
        Select Case mintErrorCount
            Case 1
                shpHead.Visible = True
            Case 2
                linTorso.Visible = True
            Case 3
                linArm1.Visible = True
            Case 4
                linArm2.Visible = True
            Case 5
                linLeg1.Visible = True
            Case 6
                linLeg2.Visible = True
                lblYouWinLose.Caption = "SAMPEAN MATEK"
                blnGameOver = True
        End Select
    End If

    If blnGameOver Then
        mintGamesPlayed = mintGamesPlayed + 1
        lblGamesPlayed.Caption = CStr(mintGamesPlayed)
        lblGamesWon.Caption = Format$(mintGamesWon)
        lblYouWinLose.Visible = True
        tmrWinLoseTimer.Enabled = True
        For intX = 1 To 26
            cmdLetterPool(intX).Enabled = False
        Next
        For intX = 1 To 20
            If mstrLabelLetter(intX) <> " " Then
                lblPuzzleLetter(intX).Caption = mstrLabelLetter(intX)
            End If
            lblPuzzleLetter(intX).Enabled = False
        Next
        lblClue.Enabled = False
        lblPick.Enabled = False
        cmdSolve.Enabled = False
    End If

End Sub


'----------------------------------------------------------------------
Private Sub cmdExit_Click()
'----------------------------------------------------------------------

    Unload Me

End Sub




'**********************************************************************
'*                      Timer Event PROSEDUR                        *
'**********************************************************************

'----------------------------------------------------------------------
Private Sub tmrWinLoseTimer_Timer()
'----------------------------------------------------------------------

    lblYouWinLose.Visible = Not lblYouWinLose.Visible

End Sub

'**********************************************************************
'*                     PROGRAM PROSEDUR                      *
'**********************************************************************

'----------------------------------------------------------------------
Private Sub LoadGameData()
'----------------------------------------------------------------------

    Dim intFileNbr  As Integer
    
    intFileNbr = FreeFile
    Open (App.Path & "\HMPuzz.dat") For Input As #intFileNbr
    Do Until EOF(intFileNbr)
        mintNbrOfPuzzles = mintNbrOfPuzzles + 1
        ReDim Preserve maudtPuzzleRec(1 To mintNbrOfPuzzles)
        With maudtPuzzleRec(mintNbrOfPuzzles)
            Input #intFileNbr, .Puzzle, .Clue
        End With
    Loop
    Close #intFileNbr
    
    intFileNbr = FreeFile
    Open (App.Path & "\HMClues.dat") For Input As #intFileNbr
    Do Until EOF(intFileNbr)
        mintNbrOfClues = mintNbrOfClues + 1
        ReDim Preserve maudtClueRec(1 To mintNbrOfClues)
        With maudtClueRec(mintNbrOfClues)
            Input #intFileNbr, .ClueCode, .ClueDescription
        End With
    Loop
    Close #intFileNbr
    
    intFileNbr = FreeFile
    Open (App.Path & "\LastPuzz.dat") For Input As #intFileNbr
    Input #intFileNbr, mintLastPuzzle
    Close #intFileNbr
    
End Sub


'----------------------------------------------------------------------
Private Sub SetUpFormForNewGame()
'----------------------------------------------------------------------

    Dim intX    As Integer

    shpHead.Visible = False
    linArm1.Visible = False
    linArm2.Visible = False
    linTorso.Visible = False
    linLeg1.Visible = False
    linLeg2.Visible = False

    For intX = 1 To 20
        mstrLabelLetter(intX) = " "
        lblPuzzleLetter(intX).Caption = ""
        lblPuzzleLetter(intX).Enabled = True
    Next

    lblClue.Enabled = True
    lblPick.Enabled = True

    For intX = 1 To 26
        cmdLetterPool(intX).Enabled = True
    Next

    cmdSolve.Enabled = True

    tmrWinLoseTimer.Enabled = False
    lblYouWinLose.Visible = False

End Sub


'----------------------------------------------------------------------
Private Sub GetNextPuzzle()
'----------------------------------------------------------------------

    Dim intX    As Integer

    If mintLastPuzzle = UBound(maudtPuzzleRec) Then
        mintLastPuzzle = 1
    Else
        mintLastPuzzle = mintLastPuzzle + 1
    End If

    With maudtPuzzleRec(mintLastPuzzle)
        mstrPuzzle = .Puzzle & " "
        For intX = 1 To UBound(maudtClueRec)
            If maudtClueRec(intX).ClueCode = .Clue Then
                lblClue.Caption = "Clue: " _
                                & maudtClueRec(intX).ClueDescription
                Exit For
            End If
        Next
    End With
    
End Sub


'----------------------------------------------------------------------
Private Function LoadPuzzleIntoInterface() As Boolean
'----------------------------------------------------------------------
    
    Dim intPuzzLen      As Integer
    Dim intPuzzPos      As Integer
    Dim intLabelPos     As Integer
    Dim intSpace        As Integer
    Dim intWordLen      As Integer
    Dim strCurrWord     As String
    Dim intX            As Integer
    
    LoadPuzzleIntoInterface = True
    
    intPuzzLen = Len(mstrPuzzle)
    If intPuzzLen > 21 Then
        LoadPuzzleIntoInterface = False
        Exit Function
    End If
    
    intPuzzPos = 1
    intLabelPos = 1

    
    Do Until intPuzzPos > intPuzzLen
    
        intSpace = InStr(intPuzzPos, mstrPuzzle, " ")
        strCurrWord = Mid$(mstrPuzzle, _
                           intPuzzPos, _
                           intSpace - intPuzzPos)
        intWordLen = Len(strCurrWord)
        
        If intWordLen > 10 Then
            LoadPuzzleIntoInterface = False
            Exit Function
        End If
        
        If (intLabelPos + intWordLen) > 21 Then
            LoadPuzzleIntoInterface = False
            Exit Function
        End If
        
        If intLabelPos <= 10 Then
            If (intLabelPos + intWordLen) > 11 Then
                intLabelPos = 11
            End If
        End If
        
        For intX = intPuzzPos To (intPuzzPos + intWordLen - 1)
            If Not (UCase$(Mid$(mstrPuzzle, intX, 1)) Like "[A-Z]") Then
                LoadPuzzleIntoInterface = False
                Exit Function
            End If
            mstrLabelLetter(intLabelPos) = Mid$(mstrPuzzle, intX, 1)
            intLabelPos = intLabelPos + 1
        Next
        
        If intLabelPos <> 11 Then
            intLabelPos = intLabelPos + 1
        End If
        
        intPuzzPos = intSpace + 1
        
    Loop


End Function


'----------------------------------------------------------------------
Private Sub SaveLastPuzzleNumber()
'----------------------------------------------------------------------

    Dim intFileNbr  As Integer

    intFileNbr = FreeFile
    Open (App.Path & "\LastPuzz.dat") For Output As #intFileNbr
    Write #intFileNbr, mintLastPuzzle
    Close #intFileNbr

End Sub
