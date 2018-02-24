VERSION 5.00
Begin VB.Form frmHMSplash 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4044
   ClientLeft      =   4560
   ClientTop       =   2892
   ClientWidth     =   6348
   DrawMode        =   2  'Blackness
   Icon            =   "HMSPLASH.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4044
   ScaleWidth      =   6348
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Interval        =   1000
      Left            =   5880
      Top             =   3480
   End
   Begin VB.Label lblversi 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Versi 1.0"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   252
      Left            =   4440
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H0000FFFF&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      Height          =   12
      Left            =   3000
      Shape           =   2  'Oval
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00FF0000&
      DrawMode        =   8  'Xor Pen
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      DrawMode        =   4  'Mask Not Pen
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   132
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   720
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2012-2013 Febrian Dwi Putra "
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   3720
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Label lbltebak 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tebak Matek"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   732
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   6132
   End
   Begin VB.Line linLeg2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3360
      Y1              =   1800
      Y2              =   2400
   End
   Begin VB.Line linLeg1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   2640
      Y1              =   1800
      Y2              =   2400
   End
   Begin VB.Line linArm2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3480
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line linArm1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   2520
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line linTorso 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3000
      Y1              =   1080
      Y2              =   1800
   End
   Begin VB.Shape shpHead 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2400
      Shape           =   3  'Circle
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Line linGallows1 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3000
      X2              =   3000
      Y1              =   360
      Y2              =   600
   End
   Begin VB.Line linGallowsBase 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4080
      X2              =   3600
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Line linGallows2 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3840
      X2              =   3000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line linGallows3 
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   3840
      X2              =   3840
      Y1              =   360
      Y2              =   2520
   End
End
Attribute VB_Name = "frmHMSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterForm Me
    
End Sub


Private Sub tmrSplash_Timer()
    
    Static FebriDraw  As Integer
    
    Select Case FebriDraw
        Case 0
            
            linGallows1.Visible = True
            linGallows2.Visible = True
            linGallows3.Visible = True
            linGallowsBase.Visible = True
        Case 1
            shpHead.Visible = True
            Shape1.Visible = True
            Shape2.Visible = True
            Shape3.Visible = True
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
            tmrSplash.Interval = 2000
        Case 7
            lblversi.Visible = True
            lbltebak.Visible = True
        Case 8
            lblCopyright.Visible = True
        Case 9
            ' do nothing - keep form displayed for 2 more seconds
        Case 10
            tmrSplash.Enabled = False
            frmHMMain.Show
            Unload Me
    End Select
    
    FebriDraw = FebriDraw + 1
    
End Sub
