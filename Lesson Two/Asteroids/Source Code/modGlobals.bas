Attribute VB_Name = "modGlobals"
Option Explicit
'//This is just a module of all the needed declarations and such for the game

'//Constants
Public Const NUM_SHIPS = 3
Public Const NUM_ASTEROID_SIZES = 3
Public Const ASTEROID_WIDTH = 40

'//UDTs
Private Type udtBullet
    X As Long
    Y As Long
    Used As Boolean
End Type

Private Type udtPilot '//defines a player ship
    PilotName As String
    PilotShip As Byte
    PilotLevel As Byte '//determines max num of asteroids
    PilotLives As Byte
    PilotScore As Long
    PilotHealth As Long
    PilotInput As KeyboardInput
    '//Stats
    Strength As Integer
    Defense As Integer
    Speed As Integer
    '//location
    X As Long
    Y As Long
    '//ship size
    ShipWidth As Integer
    ShipHeight As Integer
    '//bullets
    Bullets(9) As udtBullet
    BulletDelay As Integer
End Type

Private Type udtAsteroid '//defined an asteroid
    AsteroidNum As Byte '//size and what image to use
    X As Long
    Y As Long
    Taken As Boolean
End Type

'//Graphics Declarations
'//hold all our bitmaps in these BitBltSurfaces(kinda setup like DX i think)
Public bbsBullet As BitBltSurface
Public bbsBulletMask As BitBltSurface
Public bbsShips(NUM_SHIPS - 1) As BitBltSurface
Public bbsShipsMasks(NUM_SHIPS - 1) As BitBltSurface
Public bbsAsteroids(NUM_ASTEROID_SIZES - 1) As BitBltSurface '3 sizes
Public bbsAsteroidsMasks(NUM_ASTEROID_SIZES - 1) As BitBltSurface '3 sizes
Public bbsBackbuffer As BitBltSurface
Public bbsBackbuffer2 As BitBltSurface

'//main Declarations
Public ScreenHDC As Long '//global variable for screen HDC
Public Pilot As udtPilot '//US
Public Asteroids() As udtAsteroid '//ENEMIES
'//GAME STATE(RUNNING, PAUSED, NO GAME)
Public GameRunning As Byte '0 = No Game, 1 = Running, 2 = paused

