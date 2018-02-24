VERSION 5.00
Begin VB.Form frmCRMain 
   BackColor       =   &H000040C0&
   Caption         =   "Cara Memainkan"
   ClientHeight    =   4080
   ClientLeft      =   108
   ClientTop       =   432
   ClientWidth     =   6444
   Icon            =   "CRMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   6444
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   6360
      Top             =   4080
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Powered By Febrian Dwi Putra"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   372
      Left            =   120
      TabIndex        =   5
      Top             =   3840
      Width           =   3612
   End
   Begin VB.Label Label5 
      Caption         =   " 4. Kalau Sudah MATEK Klik Main Lagi"
      Height          =   252
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   3492
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFC0FF&
      Caption         =   " 3. Kalau Sudah Tahu Jawabannya Klik Langsung Tulis"
      Height          =   252
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   3972
   End
   Begin VB.Label Label3 
      Caption         =   " 2. Pilih Tombol A-Z Yang mengarah dalam Clue"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   3612
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Caption         =   " 1. Tebak Kata menurut Clue (Katagorinya)"
      Height          =   252
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3252
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "~ Cara Permainan :"
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   -120
      TabIndex        =   0
      Top             =   240
      Width           =   3972
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H000000FF&
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   4920
      Shape           =   3  'Circle
      Top             =   1200
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   4800
      Top             =   1440
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   4680
      Shape           =   2  'Oval
      Top             =   1200
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Line linGallows1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   600
      Y2              =   960
   End
   Begin VB.Line linLeg2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4920
      X2              =   5400
      Y1              =   2640
      Y2              =   3480
   End
   Begin VB.Line linLeg1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4920
      X2              =   4440
      Y1              =   2640
      Y2              =   3480
   End
   Begin VB.Line linArm2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   4920
      X2              =   5640
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line linArm1 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   4920
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line linTorso 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   4920
      X2              =   4920
      Y1              =   1680
      Y2              =   2640
   End
   Begin VB.Shape shpHead 
      BorderWidth     =   3
      Height          =   732
      Left            =   4200
      Shape           =   3  'Circle
      Top             =   960
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Line linGallowsBase 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   6240
      X2              =   5520
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line linGallows2 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5880
      X2              =   4920
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line linGallows3 
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   5880
      X2              =   5880
      Y1              =   600
      Y2              =   3960
   End
End
Attribute VB_Name = "frmCRMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub timer1_Timer()
    
    Static FebriDraw  As Integer
    
    Select Case FebriDraw
        Case 0
            linGallowsBase.Visible = True
        Case 1
           linGallows3.Visible = True
        Case 2
            linGallows2.Visible = True
        Case 3
            linGallows1.Visible = True
        Case 4
            shpHead.Visible = True
        Case 5
            Shape1.Visible = True
        Case 6
            Shape2.Visible = True
        Case 7
            Shape3.Visible = True
        Case 8
             linTorso.Visible = True
        Case 9
            linArm1.Visible = True
        Case 10
            linArm2.Visible = True
        Case 11
            linLeg1.Visible = True
        Case 12
            linLeg2.Visible = True
            Timer1.Interval = 2000
        Case 13
            'Dikosongi Durasi 2 Detik
      
    End Select
    
    FebriDraw = FebriDraw + 1
    
End Sub

