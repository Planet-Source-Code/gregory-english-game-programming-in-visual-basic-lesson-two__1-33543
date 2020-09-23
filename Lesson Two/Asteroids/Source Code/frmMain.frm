VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form frmMain 
   Caption         =   "Asteroids!"
   ClientHeight    =   4980
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   332
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   479
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer tmrLevel 
      Enabled         =   0   'False
      Interval        =   20000
      Left            =   4320
      Top             =   4560
   End
   Begin VB.PictureBox picGame 
      BackColor       =   &H00000000&
      Height          =   4815
      Left            =   120
      ScaleHeight     =   317
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   301
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label lblLose 
         BackStyle       =   0  'Transparent
         Caption         =   "GAME OVER!"
         BeginProperty Font 
            Name            =   "Maiandra GD"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   0
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin MediaPlayerCtl.MediaPlayer mpMIDI 
      Height          =   375
      Left            =   4680
      TabIndex        =   11
      Top             =   3360
      Visible         =   0   'False
      Width           =   495
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -380
      WindowlessVideo =   0   'False
   End
   Begin VB.Label lblLives 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3360
      Width           =   2535
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2400
      Width           =   2535
   End
   Begin VB.Label lblHealth 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot Lives"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4680
      TabIndex        =   5
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot Level"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4680
      TabIndex        =   4
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot Score"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   4680
      TabIndex        =   3
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label lblMisc 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pilot Health"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label lblMisc 
      BackStyle       =   0  'Transparent
      Caption         =   "ASTEROIDS!"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.Menu mnuNew 
      Caption         =   "New Game"
   End
   Begin VB.Menu mnuMusic 
      Caption         =   "Music"
      Begin VB.Menu mnuTheme1 
         Caption         =   "Theme 1"
      End
      Begin VB.Menu mnuTheme2 
         Caption         =   "Theme 2"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'//ALL STAR FIELD STUFF
'//for our starfield
Private Type udtStar
    X As Integer
    Y As Integer
    Z As Single '//distance from us(aka drawwidth AND how fast it goes past us)
    Taken As Boolean
End Type

'//constants
Private Const NUM_MAX_STARS = 50
Private Stars(NUM_MAX_STARS - 1) As udtStar 'starfield

'//FPS
Private FrameCounter As Integer
Private FPSTimer As Long
Private FPS As Integer

Private Sub GetFPS()
If GetTickCount >= (FPSTimer + 1000) Then
    FPS = FrameCounter
    FrameCounter = 0
    FPSTimer = GetTickCount
Else
    FrameCounter = FrameCounter + 1
End If

frmMain.Caption = "Asteroids running at " & FPS & " frames per second"
End Sub

Private Sub InitStars()
Dim X As Long

For X = 0 To NUM_MAX_STARS - 1
    Stars(X).Taken = True
    Stars(X).X = Int(Rnd * picGame.Width)
    Stars(X).Y = Int(Rnd * picGame.Width)
    Stars(X).Z = Int(Rnd * 5) + 1
Next X
End Sub

Private Sub DoStars()
Dim X As Long

For X = 0 To NUM_MAX_STARS - 1
    If Stars(X).Taken = False Then
        Stars(X).Taken = True
        Stars(X).X = Int(Rnd * picGame.Width)
        Stars(X).Y = 0
        Stars(X).Z = Int(Rnd * 5) + 1
    Else
        DrawWidth = 1
        Stars(X).Y = Stars(X).Y + Stars(X).Z
        picGame.PSet (Stars(X).X, Stars(X).Y), vbWhite
        If Stars(X).Y > picGame.Height Then
            Stars(X).Taken = False
        End If
    End If
Next X
End Sub

Public Sub MainLoop()
'//tmp time
Dim tmpTime As Long

'//Initialize Game Data and graphics
InitStars 'Star field
InitSurfaces 'graphics
Pilot.X = (picGame.Width / 2) - (bbsShips(Pilot.PilotShip).Width / 2)
Pilot.Y = picGame.Height - bbsShips(Pilot.PilotShip).Height
Pilot.BulletDelay = 100
tmrLevel.Enabled = True
ReDim Asteroids((Pilot.PilotLevel * 5) - 1)
Do While GameRunning = 1 Or GameRunning = 2
    tmpTime = timeGetTime()
    
    If GameRunning = 1 Then
        '//clear the screen
        BitBlt bbsBackbuffer.hdc, 0, 0, picGame.Width, picGame.Height, bbsBackbuffer2.hdc, 0, 0, SRCCOPY
        
        DoStars '//update all the stars
        DoPilot '//update our pilot
        DoAsteroids '//update our asteroids
        DoBullets 'update our bullets
        '//blt backbuffer to main hdc
        BitBlt ScreenHDC, 0, 0, picGame.Width, picGame.Height, bbsBackbuffer.hdc, 0, 0, SRCCOPY
        
        '//check for deaths
        If Pilot.PilotHealth <= 0 Then
            '//any live left?
            If Pilot.PilotLevel > 0 Then
                Pilot.PilotLevel = Pilot.PilotLevel - 1
            End If
            GameRunning = 0
            lblLose.Visible = True
        End If
        '//set form data
        lblHealth.Caption = Pilot.PilotHealth
        lblScore.Caption = Pilot.PilotScore
        lblLives.Caption = Pilot.PilotLives
        lblLevel.Caption = Pilot.PilotLevel
        GetFPS
    End If
    '//frame cap limiter
    Do Until timeGetTime >= tmpTime + 30
    Loop
    DoEvents
Loop

End Sub

Private Sub Form_Load()
'//global variable for picGame.HDC, easier use
ScreenHDC = picGame.hdc
mpMIDI.FileName = App.Path & "\Audio\Theme1.mid"
mpMIDI.Play
End Sub

Private Sub Form_Unload(Cancel As Integer)
'//clear memory of all Device Contexts created
modEngine.DestroyHdcs
End Sub

Private Sub lblHealth_Click()
Pilot.PilotHealth = Pilot.PilotHealth + 1000
End Sub

Private Sub mnuNew_Click()
lblLose.Visible = False
Pilot.PilotName = InputBox("What is your Pilot's Name?", "Asteroids!")
If Pilot.PilotName <> "" Then
    frmShipChoose.Show
End If
End Sub

Private Sub mnuTheme1_Click()
mpMIDI.Stop
mpMIDI.FileName = App.Path & "\Audio\Theme1.mid"
mpMIDI.Play
End Sub

Private Sub mnuTheme2_Click()
mpMIDI.Stop
mpMIDI.FileName = App.Path & "\Audio\Theme2.mid"
mpMIDI.Play
End Sub

Private Sub tmrLevel_Timer()
Randomize Timer
Pilot.PilotLevel = Pilot.PilotLevel + 1
ReDim Preserve Asteroids((Pilot.PilotLevel * 5) - 1)
Me.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
End Sub
