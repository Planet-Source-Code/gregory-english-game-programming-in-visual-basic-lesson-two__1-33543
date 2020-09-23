VERSION 5.00
Begin VB.Form frmShipChoose 
   BackColor       =   &H00000000&
   Caption         =   "Choose Your Ship!"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4845
   LinkTopic       =   "Form1"
   ScaleHeight     =   131
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   323
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label lblWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "54"
      Height          =   255
      Index           =   2
      Left            =   -360
      TabIndex        =   23
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblHeight 
      BackStyle       =   0  'Transparent
      Caption         =   "43"
      Height          =   255
      Index           =   2
      Left            =   -480
      TabIndex        =   22
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblHeight 
      BackStyle       =   0  'Transparent
      Caption         =   "33"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   20
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label lblWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "46"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   19
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblHeight 
      BackStyle       =   0  'Transparent
      Caption         =   "45"
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   18
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   8
      Left            =   3480
      TabIndex        =   17
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defense:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   3480
      TabIndex        =   16
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblStr 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   15
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblDef 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   14
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   13
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   3480
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   1800
      TabIndex        =   11
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defense:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblStr 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   9
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblDef 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   8
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1800
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Strength:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Defense:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblStr 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   3
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblDef 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   1
      Top             =   1680
      Width           =   495
   End
   Begin VB.Label lblMisc 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Speed:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Image imgShip 
      Height          =   645
      Index           =   2
      Left            =   3840
      Picture         =   "frmShipChoose.frx":0000
      Top             =   240
      Width           =   810
   End
   Begin VB.Image imgShip 
      Height          =   495
      Index           =   1
      Left            =   2160
      Picture         =   "frmShipChoose.frx":1BCE
      Top             =   360
      Width           =   750
   End
   Begin VB.Image imgShip 
      Height          =   675
      Index           =   0
      Left            =   480
      Picture         =   "frmShipChoose.frx":2FA8
      Top             =   240
      Width           =   690
   End
End
Attribute VB_Name = "frmShipChoose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub imgShip_Click(Index As Integer)
If MsgBox("Are you sure you want this ship?", vbYesNo) = vbYes Then
    Pilot.PilotShip = Index
    Pilot.PilotHealth = 100
    Pilot.PilotLives = 3
    Pilot.PilotLevel = 1
    Pilot.PilotScore = 0
    Pilot.Strength = Val(lblStr(Index).Caption)
    Pilot.Defense = Val(lblDef(Index).Caption)
    Pilot.Speed = Val(lblSpeed(Index).Caption)
    Pilot.ShipHeight = Val(lblHeight(Index).Caption)
    Pilot.ShipWidth = Val(lblWidth(Index).Caption)
    frmMain.Enabled = True
    Me.Hide
    GameRunning = 1
    frmMain.MainLoop
End If
End Sub
