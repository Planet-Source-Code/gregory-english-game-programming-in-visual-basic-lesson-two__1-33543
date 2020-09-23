Attribute VB_Name = "modEngine"
Option Explicit
'***************************
'Name:modEngine            *
'Desc: A Simple Game Engine*
'Started: 4/03/02          *
'***************************
'//Windows API Declarations
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function IntersectRect Lib "user32" (lpDestRect As RECT, lpSrc1Rect As RECT, lpSrc2Rect As RECT) As Long
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Dim MemHdc() As Long
Dim BitMapHdc() As Long
Dim TrashbmpHdc() As Long
Dim NumOfDcs As Integer

'//Enumerations
'/bitblt surface types
Public Enum BB_SURF_TYPE
    SURF_BITMAP = 1
    SURF_BACKBUFFER = 2
End Enum

'/sound flags
Public Enum SND_FLAGS
    SND_ASYNC = &H1 '//lets you play a new wav sound, interrupting another
    SND_LOOP = &H8 '//loops the wav sound
    SND_NODEFAULT = &H2 '//if wav file not there, then make sure NOTHING plays
    SND_SYNC = &H0 '//no control to program til wav is done playing
    SND_NOSTOP = &H10 '//if a wav file is already playing then it wont interrupt
End Enum

'//UDTs
'//for input
Public Type KeyboardInput
    btnDown As Boolean
    btnRight As Boolean
    btnUp As Boolean
    btnLeft As Boolean
    btnA As Boolean
    btnB As Boolean
    btnC As Boolean
    btnD As Boolean
    btnE As Boolean
    btnF As Boolean
    btnG As Boolean
    btnH As Boolean
    btnI As Boolean
    btnJ As Boolean
    btnK As Boolean
    btnL As Boolean
    btnM As Boolean
    btnN As Boolean
    btnO As Boolean
    btnP As Boolean
    btnQ As Boolean
    btnR As Boolean
    btnS As Boolean
    btnT As Boolean
    btnU As Boolean
    btnV As Boolean
    btnW As Boolean
    btnX As Boolean
    btnY As Boolean
    btnZ As Boolean
    btnAlt As Boolean
    btnControl As Boolean
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type BitBltSurface
    hdc As Long
    SurfType As BB_SURF_TYPE
    Width As Integer
    Height As Integer
End Type

'//Constants for bitblt rasterization options
Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCINVERT = &H660046
Public Const SRCPAINT = &HEE0086

'//********BITBLT ENGINE********\\'
'//THESE SUBS WERE GOTTEN FROM A BITBLT TUTORIAL AT VOODOOVB.THENEXUS.BC.CA great tutorial
Private Function CreateMemHdc(ScreenHDC As Long, Width As Integer, Height As Integer) As Long
ReDim Preserve MemHdc(NumOfDcs)
ReDim Preserve BitMapHdc(NumOfDcs)
ReDim Preserve TrashbmpHdc(NumOfDcs)

MemHdc(NumOfDcs) = CreateCompatibleDC(ScreenHDC)
    If MemHdc(NumOfDcs) Then
        BitMapHdc(NumOfDcs) = CreateCompatibleBitmap(ScreenHDC, Width, Height)
        If BitMapHdc(NumOfDcs) Then
            TrashbmpHdc(NumOfDcs) = SelectObject(MemHdc(NumOfDcs), BitMapHdc(NumOfDcs))
            CreateMemHdc = MemHdc(NumOfDcs)
        End If
    End If
NumOfDcs = NumOfDcs + 1
End Function

Private Sub LoadBmpToHdc(MHdc As Long, FileN As String)
Dim OrgBmp As Long
OrgBmp = SelectObject(MHdc, LoadPicture(FileN))
If OrgBmp Then
    DeleteObject (OrgBmp)
End If
End Sub

Public Sub DestroyHdcs()
Dim i As Integer
For i = 0 To NumOfDcs - 1
    BitMapHdc(i) = SelectObject(MemHdc(i), TrashbmpHdc(i))
    DeleteObject (BitMapHdc(i))
    DeleteDC (MemHdc(i))
Next i
End Sub

Public Function CreateBitBltSurface(FileName As String, ScreenHDC As Long, BBSurfType As BB_SURF_TYPE, intWidth As Integer, intHeight As Integer) As BitBltSurface
If BBSurfType = SURF_BITMAP Then '//Create us a bitmap hdc
    CreateBitBltSurface.hdc = CreateMemHdc(ScreenHDC, intWidth, intHeight) '//create the dc
    Call LoadBmpToHdc(CreateBitBltSurface.hdc, FileName) '//load in the bmp
    CreateBitBltSurface.Width = intWidth '//set some basic properties
    CreateBitBltSurface.Height = intHeight '//
    CreateBitBltSurface.SurfType = BBSurfType '//
ElseIf BBSurfType = SURF_BACKBUFFER Then '//Create us a backbuffer
    CreateBitBltSurface.hdc = CreateMemHdc(ScreenHDC, intWidth, intHeight)
    CreateBitBltSurface.Width = intWidth
    CreateBitBltSurface.Height = intHeight
    CreateBitBltSurface.SurfType = BBSurfType
End If
End Function


'//********INPUT ENGINE********\\'
'//gets input for all the buttons of our UDT KeyboardInput
'//you can easily add a button like Numbers by adding the var to the udt
'//and .VARNAME = GetAsyncKeyState(vbKeyWhatever)
Public Function GetKeyboardInput(Keyboard As KeyboardInput)

With Keyboard
    .btnDown = GetAsyncKeyState(vbKeyDown)
    .btnUp = GetAsyncKeyState(vbKeyUp)
    .btnRight = GetAsyncKeyState(vbKeyRight)
    .btnLeft = GetAsyncKeyState(vbKeyLeft)
    .btnA = GetAsyncKeyState(vbKeyA)
    .btnB = GetAsyncKeyState(vbKeyB)
    .btnC = GetAsyncKeyState(vbKeyC)
    .btnD = GetAsyncKeyState(vbKeyD)
    .btnE = GetAsyncKeyState(vbKeyE)
    .btnF = GetAsyncKeyState(vbKeyF)
    .btnG = GetAsyncKeyState(vbKeyG)
    .btnH = GetAsyncKeyState(vbKeyH)
    .btnI = GetAsyncKeyState(vbKeyI)
    .btnJ = GetAsyncKeyState(vbKeyJ)
    .btnK = GetAsyncKeyState(vbKeyK)
    .btnL = GetAsyncKeyState(vbKeyL)
    .btnM = GetAsyncKeyState(vbKeyM)
    .btnN = GetAsyncKeyState(vbKeyN)
    .btnO = GetAsyncKeyState(vbKeyO)
    .btnP = GetAsyncKeyState(vbKeyP)
    .btnQ = GetAsyncKeyState(vbKeyQ)
    .btnR = GetAsyncKeyState(vbKeyR)
    .btnS = GetAsyncKeyState(vbKeyS)
    .btnT = GetAsyncKeyState(vbKeyT)
    .btnU = GetAsyncKeyState(vbKeyU)
    .btnV = GetAsyncKeyState(vbKeyV)
    .btnW = GetAsyncKeyState(vbKeyW)
    .btnX = GetAsyncKeyState(vbKeyX)
    .btnY = GetAsyncKeyState(vbKeyY)
    .btnZ = GetAsyncKeyState(vbKeyZ)
    .btnControl = GetAsyncKeyState(vbKeyControl)
End With

End Function
'//********SOUND ENGINE********\\'
'//plays a wave sound
Public Function PlayWave(FileName As String, Flags As SND_FLAGS)
sndPlaySound FileName, Flags
End Function
'//stops a wave sound
Public Function StopWave()
sndPlaySound "", SND_NODEFAULT Or SND_ASYNC
End Function


'//********COLLISION DETECTION********\\'
'//help function for creating the RECT Type
Public Function CreateRect(X1 As Long, X2 As Long, Y1 As Long, Y2 As Long) As RECT
CreateRect.Left = X1
CreateRect.Right = X2
CreateRect.Top = Y1
CreateRect.Bottom = Y2
End Function
'//use intersectrect for collision
Public Function Collide(Rec1 As RECT, Rec2 As RECT) As Boolean
Dim EmptyRect As RECT
Collide = IntersectRect(EmptyRect, Rec1, Rec2)
End Function
