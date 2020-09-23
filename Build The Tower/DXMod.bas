Attribute VB_Name = "DXMod"
Option Explicit

'The Main Object of DirectX Components
Public DX7 As DirectX7
Public DxD As DirectDraw7

'The Variables for DirectX Object
Public LemmDX  As DirectDrawSurface7        'The Character Sprites
Public TexMouse As DirectDrawSurface7
Public Primary As DirectDrawSurface7        'The Screen We See
Public BackBuffer As DirectDrawSurface7     'The Backbuffer we need in Flipping Mode
Public Sprite As DirectDrawSurface7         'hold the sprites for other objects
Public SpriteHujan As DirectDrawSurface7    'hold the sprites for other objects
Public Display As DirectDrawSurface7
Public Layar As DirectDrawSurface7          'Layar tampilan untuk penggulungan
Public CuacaDX As DirectDrawSurface7
Public ToolBarDX As DirectDrawSurface7      'Layar Toolbar
Public IconWorker As DirectDrawSurface7
Public TempDX As DirectDrawSurface7
Public HujanDX As DirectDrawSurface7        'Hujan DX
Public MiniMapDX As DirectDrawSurface7      'MiniMap DX
Public SketsaDX As DirectDrawSurface7       'Sketsa Minimap
Public SplashDX As DirectDrawSurface7       'Splash Screen

'Untuk pengontrolan Gamma
Public mobjGammaControler As DirectDrawGammaControl    'The object that gets/sets gamma ramps
Public mudtGammaRamp As DDGAMMARAMP                    'The gamma ramp we'll use to alter the screen state
Public mudtOriginalRamp As DDGAMMARAMP                 'The gamma ramp we'll use to store the original screen state
Public mintRedVal As Integer                        'Store the currend red value w.r.t. original
Public mintGreenVal As Integer                      'Store the currend green value w.r.t. original
Public mintBlueVal As Integer                       'Store the currend blue value w.r.t. original
Public mblnGamma As Boolean                         'Do we have gamma support?
Public mblnFadeIn As Boolean                        'Should we fade back in?

'Untuk Status Box
Public BoxStatus As DirectDrawSurface7

'Program flow variables
Dim mlngFrameTime As Long                   'How long since last frame?
Dim mlngTimer As Long                       'How long since last FPS count update?
Dim mintFPSCounter As Integer               'Our FPS counter
Public mintFPS As Integer                   'Our FPS storage variable

Global Const MAX_FPS = 70

Public ddschar As DDSURFACEDESC2
Public ddsmap As DDSURFACEDESC2

Public Keyboard As New KeyboardCls
Public SFXMusik As New SoundCls

Public Sub SetGamma(intRed As Integer, intGreen As Integer, intBlue As Integer)

Dim i As Integer

    'Alter the gamma ramp to the percent given by comparing to original state
    'A value of zero ("0") for intRed, intGreen, or intBlue will result in the
    'gamma level being set back to the original levels. Anything ABOVE zero will
    'fade towards FULL colour, anything below zero will fade towards NO colour
    For i = 0 To 255
        If intRed < 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.red(i)) * (100 - Abs(intRed)) / 100)
        If intRed = 0 Then mudtGammaRamp.red(i) = mudtOriginalRamp.red(i)
        If intRed > 0 Then mudtGammaRamp.red(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.red(i))) * (100 - intRed) / 100))
        If intGreen < 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.green(i)) * (100 - Abs(intGreen)) / 100)
        If intGreen = 0 Then mudtGammaRamp.green(i) = mudtOriginalRamp.green(i)
        If intGreen > 0 Then mudtGammaRamp.green(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.green(i))) * (100 - intGreen) / 100))
        If intBlue < 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(ConvToUnSignedValue(mudtOriginalRamp.blue(i)) * (100 - Abs(intBlue)) / 100)
        If intBlue = 0 Then mudtGammaRamp.blue(i) = mudtOriginalRamp.blue(i)
        If intBlue > 0 Then mudtGammaRamp.blue(i) = ConvToSignedValue(65535 - ((65535 - ConvToUnSignedValue(mudtOriginalRamp.blue(i))) * (100 - intBlue) / 100))
    Next
    
    mobjGammaControler.SetGammaRamp DDSGR_DEFAULT, mudtGammaRamp

End Sub


Function Init() As Boolean
    On Error Resume Next
    
    Set DX7 = New DirectX7
    If DX7 Is Nothing Then
        MsgBox "Error Creating DirectX7 !", vbExclamation
        Exit Function
    End If
    
    Set DxD = DX7.DirectDrawCreate("")
    'set the cooperative level
    
    DxD.SetCooperativeLevel FrmPlay.hWnd, DDSCL_ALLOWMODEX Or DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    'resolusi 800 x 600 x 16 bit
    DxD.SetDisplayMode 800, 600, 16, 0, DDSDM_DEFAULT
    
    'set the primary and backbuffer for flipping chain
    ddsmap.lFlags = DDSD_BACKBUFFERCOUNT Or DDSD_CAPS
    ddsmap.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_VIDEOMEMORY
    ddsmap.lBackBufferCount = 1
    Set Primary = DxD.CreateSurface(ddsmap)
    Dim DD As DDSCAPS2
    DD.lCaps = DDSCAPS_BACKBUFFER
    Set BackBuffer = Primary.GetAttachedSurface(DD)
    
    BackBuffer.GetSurfaceDesc ddsmap
    
    Dim ddsd2 As DDSURFACEDESC2
    ddsd2.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd2.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_VIDEOMEMORY
    ddsd2.lWidth = 800
    ddsd2.lHeight = 600
    Set Display = DxD.CreateSurface(ddsd2)
    Set TempDX = DxD.CreateSurface(ddsd2)
    
    Dim ddsd3 As DDSURFACEDESC2
    ddsd3.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd3.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd3.lWidth = 828
    ddsd3.lHeight = 640
    Set Layar = DxD.CreateSurface(ddsd3)
    
    'ciptakan layar minimap
    Dim ddsd4 As DDSURFACEDESC2
    ddsd4.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT
    ddsd4.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    ddsd4.lWidth = 105
    ddsd4.lHeight = 105
    Set MiniMapDX = DxD.CreateSurface(ddsd4)
    
    'Make a new gamma controler
    Set mobjGammaControler = Primary.GetDirectDrawGammaControl

    'Fill out the original gamma ramps
    mobjGammaControler.GetGammaRamp DDSGR_DEFAULT, mudtOriginalRamp
    
    'Set our initial colour values to zero
    mintRedVal = 0
    mintGreenVal = 0
    mintBlueVal = 0
    
    Keyboard.InitDI
    Keyboard.aHwnd = FrmPlay.hWnd
    
    FrmPlay.Show
    
End Function
Function BoxRect(X, Y, X1, Y1) As RECT
    BoxRect.Top = Y
    BoxRect.Left = X
    BoxRect.Right = X1
    BoxRect.Bottom = Y1
End Function

Sub EndAll()
    On Error GoTo Keluar
    StillRunning = False
    SFXMusik.StopMusic
    SFXMusik.CloseDM
    SFXMusik.CloseDS
    Set SFXMusik = Nothing
    
    Set DX7 = Nothing
    Set DxD = Nothing
    Keyboard.CloseDI
Keluar:
    End
End Sub

Public Function LoadSprite(StrFileName As String, nColor As Long, ByRef Tex As DirectDrawSurface7) As Boolean
    On Error GoTo Keluar
    'loading file bitmap apakah berhasil atau tidak
    Dim ddsspr As DDSURFACEDESC2
    ddsspr.lFlags = DDSD_CAPS
    ddsspr.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    Set Tex = DxD.CreateSurfaceFromFile(StrFileName, ddsspr)
    
    'set the key color for the transparency thing
    Dim nColorKey As DDCOLORKEY
    nColorKey.high = nColor       'black
    nColorKey.low = nColor        'black
    Tex.SetColorKey DDCKEY_SRCBLT, nColorKey
    
    LoadSprite = True
    Exit Function
Keluar:
    MsgBox "Data Sprite tidak bisa diloading !", vbExclamation
    LoadSprite = False
End Function

Sub FPS()
    'Delay until specified FPS achieved
    Do While mlngFrameTime + (1000 \ MAX_FPS) > DX7.TickCount
        DoEvents
    Loop
    mlngFrameTime = DX7.TickCount

    'Count FPS
    If mlngTimer + 1000 <= DX7.TickCount Then
        mlngTimer = DX7.TickCount
        mintFPS = mintFPSCounter + 1
        mintFPSCounter = 0
    Else
        mintFPSCounter = mintFPSCounter + 1
    End If
End Sub

