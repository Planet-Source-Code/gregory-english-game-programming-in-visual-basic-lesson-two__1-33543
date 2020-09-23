Attribute VB_Name = "modGame"
Option Explicit

'//This sub loads all of our surfaces into a DC(like a picbox) for later use
'//We use the created udt BitBltSurface which has widths and height and the hdc of the bitmap
Public Sub InitSurfaces()
'//looping variables
Dim X As Long


'//ships
bbsShips(0) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Ship0.bmp", ScreenHDC, SURF_BITMAP, 46, 45)
bbsShips(1) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Ship1.bmp", ScreenHDC, SURF_BITMAP, 50, 33)
bbsShips(2) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Ship2.bmp", ScreenHDC, SURF_BITMAP, 54, 43)
bbsShipsMasks(0) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\ShipMask0.bmp", ScreenHDC, SURF_BITMAP, 46, 45)
bbsShipsMasks(1) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\ShipMask1.bmp", ScreenHDC, SURF_BITMAP, 50, 33)
bbsShipsMasks(2) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\ShipMask2.bmp", ScreenHDC, SURF_BITMAP, 54, 43)

'//bullet
bbsBullet = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Bullet.bmp", ScreenHDC, SURF_BITMAP, 54, 43)
bbsBulletMask = modEngine.CreateBitBltSurface(App.Path & "\Graphics\BulletMask.bmp", ScreenHDC, SURF_BITMAP, 54, 43)

'//back buffer
bbsBackbuffer = modEngine.CreateBitBltSurface("", ScreenHDC, SURF_BACKBUFFER, frmMain.picGame.Width, frmMain.picGame.Height)
bbsBackbuffer2 = modEngine.CreateBitBltSurface(App.Path & "\Graphics\backbuffer.bmp", ScreenHDC, SURF_BITMAP, frmMain.picGame.Width, frmMain.picGame.Height)

'//asteroids
bbsAsteroids(0) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Asteroid0.bmp", ScreenHDC, SURF_BITMAP, 40, 40)
bbsAsteroids(1) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Asteroid1.bmp", ScreenHDC, SURF_BITMAP, 40, 40)
bbsAsteroids(2) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\Asteroid2.bmp", ScreenHDC, SURF_BITMAP, 40, 40)
bbsAsteroidsMasks(0) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\AsteroidMask0.bmp", ScreenHDC, SURF_BITMAP, 40, 40)
bbsAsteroidsMasks(1) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\AsteroidMask1.bmp", ScreenHDC, SURF_BITMAP, 40, 40)
bbsAsteroidsMasks(2) = modEngine.CreateBitBltSurface(App.Path & "\Graphics\AsteroidMask2.bmp", ScreenHDC, SURF_BITMAP, 40, 40)

End Sub
'//Handles everything related to our pilot/ input and movement
Public Sub DoPilot()
'//Get pilot input
GetKeyboardInput Pilot.PilotInput
DoInput

'//draw the pilot
BitBlt bbsBackbuffer.hdc, Pilot.X, Pilot.Y, Pilot.ShipWidth, Pilot.ShipHeight, bbsShipsMasks(Pilot.PilotShip).hdc, 0, 0, SRCAND
BitBlt bbsBackbuffer.hdc, Pilot.X, Pilot.Y, Pilot.ShipWidth, Pilot.ShipHeight, bbsShips(Pilot.PilotShip).hdc, 0, 0, SRCINVERT

End Sub
Private Sub DoInput()
'//Check which buttons are true and if they are add the increment of the ships speed
'//which is currently in use
With Pilot
    '//movement
    If .PilotInput.btnUp = True Then
        If .Y - .Speed >= 0 Then
            .Y = .Y - .Speed
        End If
    ElseIf .PilotInput.btnDown = True Then
        If .Y + .Speed + .ShipHeight <= frmMain.picGame.Height Then
            .Y = .Y + .Speed
        End If
    End If
    
    If .PilotInput.btnLeft = True Then
        If .X - .Speed >= 0 Then
            .X = .X - .Speed
        End If
    ElseIf .PilotInput.btnRight = True Then
        If .X + .Speed + .ShipWidth <= frmMain.picGame.Width Then
            .X = .X + .Speed
        End If
    End If
    '//if control is down then create a new bullet
    If .PilotInput.btnControl = True Then
       CreateBullet
    End If
End With
End Sub
'//Handles our Asteroids
Public Sub DoAsteroids()
Randomize Timer
Dim PilotRect As RECT, AsteroidRect As RECT, X As Long

'//Create the pilot rectangle of where the pilot is
PilotRect = CreateRect(Pilot.X, Pilot.X + Pilot.ShipWidth, Pilot.Y, Pilot.Y + Pilot.ShipHeight)

'//Loop through all of the asteroids
For X = 0 To UBound(Asteroids)
    '//if the asteroid is blown up then create a new one at a random spot
    If Asteroids(X).Taken = False Then
        Asteroids(X).Taken = True
        Asteroids(X).X = Int(Rnd * frmMain.picGame.Width)
        Asteroids(X).Y = 0
        Asteroids(X).AsteroidNum = Int(Rnd * 3)
    Else '//if it already exists move it down and draw it
        Asteroids(X).Y = Asteroids(X).Y + (Asteroids(X).AsteroidNum + 1) * 2
        BitBlt bbsBackbuffer.hdc, Asteroids(X).X, Asteroids(X).Y, 40, 40, bbsAsteroidsMasks(Asteroids(X).AsteroidNum).hdc, 0, 0, SRCAND
        BitBlt bbsBackbuffer.hdc, Asteroids(X).X, Asteroids(X).Y, 40, 40, bbsAsteroids(Asteroids(X).AsteroidNum).hdc, 0, 0, SRCINVERT
        '//is the asteroid off the screen?
        If Asteroids(X).Y > frmMain.picGame.Height Then
            Asteroids(X).Taken = False
        End If
        '//did the asteroid hit the pilot?
        AsteroidRect = CreateRect(Asteroids(X).X, Asteroids(X).X + 40, Asteroids(X).Y, Asteroids(X).Y + 40)
        If Collide(AsteroidRect, PilotRect) = True Then
            Asteroids(X).Taken = False
            Pilot.PilotHealth = Pilot.PilotHealth - Asteroids(X).AsteroidNum
        End If
    End If
Next X
End Sub

Public Sub DoBullets()
'//Temporary Variables
Dim BulletRect As RECT, AsteroidRect As RECT
Dim X As Long, Y As Long

'//update bullet delay so pilot cant shoot 100000 bullets a second
Pilot.BulletDelay = Pilot.BulletDelay + 1
If Pilot.BulletDelay >= 10 Then
    Pilot.BulletDelay = 10
End If

'//loop through all of the bullets and update position
For X = 0 To 9
    If Pilot.Bullets(X).Used = True Then
        Pilot.Bullets(X).Y = Pilot.Bullets(X).Y - 5
        If Pilot.Bullets(X).Y <= 0 Then
            Pilot.Bullets(X).Used = False
        End If
        
        '//collision with an asteroid?
        '//the rect location of our bullet
        BulletRect = CreateRect(Pilot.Bullets(X).X, Pilot.Bullets(X).X + 10, Pilot.Bullets(X).Y, Pilot.Bullets(X).Y + 10)
        '//loop through all of the asteroids to see which one it hit or will hit
        For Y = 0 To UBound(Asteroids) - 1
            '//the rect location of the asteroid
            AsteroidRect = CreateRect(Asteroids(Y).X, Asteroids(Y).X + 40, Asteroids(Y).Y, Asteroids(Y).Y + 40)
            If Collide(BulletRect, AsteroidRect) = True Then
                Asteroids(Y).Taken = False
                Pilot.Bullets(X).Used = False
                Pilot.PilotScore = Pilot.PilotScore + ((Asteroids(Y).AsteroidNum + 1) * 100)
                '//play the catchy explosion sound
                PlayWave App.Path & "\Audio\Explode.wav", SND_ASYNC Or SND_NODEFAULT
            End If
        Next Y
        '//draw the bullet on the screen
        BitBlt bbsBackbuffer.hdc, Pilot.Bullets(X).X, Pilot.Bullets(X).Y, 10, 10, bbsBulletMask.hdc, 0, 0, SRCAND
        BitBlt bbsBackbuffer.hdc, Pilot.Bullets(X).X, Pilot.Bullets(X).Y, 10, 10, bbsBullet.hdc, 0, 0, SRCINVERT
    End If
Next X

End Sub

Private Sub CreateBullet()
'//Temporary Variables
Dim X As Long
'//this just creates a new bullet
'//but we first must check if the delay has been reached
If Pilot.BulletDelay >= 10 Then
    Pilot.BulletDelay = 0 '//reset the delay
        For X = 0 To 9
            If Pilot.Bullets(X).Used = False Then
                '//use simple math to align the bullet to the middle of the ship
                Pilot.Bullets(X).X = Pilot.X + (Pilot.ShipWidth / 2) - 5
                Pilot.Bullets(X).Y = Pilot.Y + 2
                Pilot.Bullets(X).Used = True
                '//play a nifty sound
                PlayWave App.Path & "\Audio\Laser.wav", SND_ASYNC Or SND_NODEFAULT
                Exit Sub
            End If
        Next X
End If
End Sub
