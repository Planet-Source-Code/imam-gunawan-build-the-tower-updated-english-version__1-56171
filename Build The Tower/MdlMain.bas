Attribute VB_Name = "MdlMain"
Option Explicit

Public StillRunning As Boolean

Private Type udtGedung
    bWalkable As Boolean
    bSpace As Boolean
    bGround As Boolean
    bTiang As Boolean
    bCor As Boolean
    bytSprX As Byte
    bytSprY As Byte
    bLadder As Boolean
    bCrane As Boolean
    bCat As Boolean
    bWall As Boolean
    bCockpit As Boolean
End Type

Private Type udtUangLayar
    bytX As Byte
    bytY As Byte
    JumlahUang As String
    Arah As Byte
    bytLong As Integer
End Type

Private Type udtWorkShop
    bytJendela As Byte
    bytPintu As Byte
    bytTangga As Byte
    intTangga As Integer
    intJendela As Integer
    intPintu As Integer
    bytBata As Byte
    intBata As Integer
    bytPasir As Byte
    intPasir As Integer
    bytSemen As Byte
    intSemen As Integer
    bytCat As Byte
    intCat As Integer
    bytKayu As Byte
    intKayu As Integer
    bytBesi As Byte
    intBesi As Integer
    bytKaca As Byte
    intKaca As Integer
    
    'nilai sementara
    bytTmpJendela As Byte
    bytTmpPintu As Byte
    bytTmpTangga As Byte
    bytTmpBata As Byte
    bytTmpSemen As Byte
    bytTmpPasir As Byte
    bytTmpCat As Byte
    bytTmpKayu As Byte
    bytTmpBesi As Byte
    bytTmpKaca As Byte
End Type

Private Type udtTotalRaw
    bytTmpJendela As Integer
    bytTmpPintu As Integer
    bytTmpTangga As Integer
    bytTmpBata As Integer
    bytTmpSemen As Integer
    bytTmpPasir As Integer
    bytTmpCat As Integer
    bytTmpKayu As Integer
    bytTmpBesi As Integer
    bytTmpKaca As Integer
End Type

Private Type udtTiangProg
    bytX As Byte
    bytY As Byte
    bytProgress As Byte
    bytCurrentProg As Byte
    bytLong As Byte
End Type

Public ArGedung(0 To MAXGEDUNGX, 0 To MAXGEDUNGY) As udtGedung

Public MoneyScreen() As udtUangLayar

Public Workshop As udtWorkShop

Public TiangProg(50) As udtTiangProg

Public TotalWorkshop As udtTotalRaw

Private Type udtRGB
    R As Integer
    G As Integer
    B As Integer
End Type

Private Type udtWorker
    Active As Boolean            'kondisi (non)aktif worker
    NoLemm As Integer            'No worker ID
    Nama As String               'Nama Worker
    Honor As Integer             'Honor worker
    Stability As Byte            'stabilitas worker
    Tolerance As Byte            'toleransi terhadap pekerjaan
    Progress As Byte             'progress kemajuan pekerjaan

    Frame As Byte                'posisi frame worker
    Way As Byte                  '
    Job As Byte
    JobX As Byte
    JobY As Byte
    Status As Byte               'status worker (walk,work,idle,rest,etc.)
    XPosition As Integer         'posisi xposisi worker di layar
    YPosition As Integer         'posisi yposisi worker di layar
    Perubah As Integer
    XTileOnArray As Byte         'posisi worker di array
    YTileOnArray As Byte         'posisi worker di array
    WorkSpeed As Byte            'nilai kecepatan perubahan antar frame dalam bekerja
    WorkNow As Byte              'nilai kecepatan pekerja saat sekarang
    CurrentSpeed As Byte         'kecepatan perubahan frame worker

    WalkSpeed As Byte            'kecepatan jalan worker
    IsMoving As Boolean          'nilai boolean menyatakan kondisi worker
    WayMove As Byte              'Arah jalan worker
    SearchMove As Byte           'Arah yang dicari worker
    MoveDuration As Byte
    XHeadSmooth As Integer       'nilai penambah untuk posisi x
    YHeadSmooth As Integer       'nilai penambah untuk posisi y

    XPoint As Byte               'nilai X yang diclick
    YPoint As Byte               'nilai Y yang diclick
End Type

Private Type udtCuaca
    Suhu As Byte             'suhu cuaca
    Status As Byte           'status cuaca
    YMove As Byte
    Panjang As Byte

    JamMulai As Byte
    MenitMulai As Byte
    JamSelesai As Byte
    MenitSelesai As Byte
    JamSekarang As Byte
    MenitSekarang As Byte

    HariIni As Date          'hari ini dalam pekerjaan

    Delayment As Byte        'bekerja dalam satu delay setiap max_fps
    EachFPS As Integer       'nilai perubah untuk max_fps
End Type

Private Type udtCrane
    BaseX As Byte
    BaseY As Byte
    HeightCrane As Byte
    LengthCrane As Byte

    'bagian cockpit crane
    XCock As Integer
    YCock As Integer
    XCockpitOnArray As Byte
    YCockpitOnArray As Byte

    'bagian kepala crane
    XHead As Byte
    YHead As Byte
    YKait As Byte
    LengthChain As Integer
    CraneSpeed As Integer
    XPositionHead As Integer
    YPositionHead As Integer
    XHeadSmooth As Integer
    WayMove As Byte

    MoveDuration As Byte
    CraneAction As Byte

    LastPlaceCraneX As Byte
    LastPlaceCraneY As Byte

    CranePilih As Boolean
    Occupied As Boolean
    LebarTiang As Byte
    XTiangOnArray As Byte
    YTiangOnArray As Byte
    PointerDraw As Boolean
    PointerLong As Byte

    PickStuff As Byte

    IsCraneMoving As Boolean
    XPoint As Byte
    YPoint As Byte

    StartMoving As Boolean
End Type

Public RGBColor As udtRGB
Public InvertRGB As Integer

Public Worker(1000) As New LemCls
Private RecWorker(1000) As udtWorker
Private RecCuaca As udtCuaca
Private InvHujan As Integer
Private HujanFPS As Byte
Private RecCrane As udtCrane

Public Cuaca As New CuacaCls
Public WorkerName(10) As String
Public intLemm As Integer
Public SelNo As Integer                 'no worker yang dipilih
Public SelNoDana As Integer             'no worker di dalam dana pekerja window
Public Selected As Boolean              'selected worker or not
Public GlobalSel As Boolean             'global selected or not
Public SelectedToolbar As Byte
Public ArSpr(0 To 7, 0 To 12) As RECT

Private intVar As Integer               'variabel private untuk dipakai
Private intVar2 As Integer

Public Gedung As New GedungCls
Public Screen As New ScreenCls
Public Crane As New CraneCls


Sub CheckTiangProg()
    'rutin untuk melakukan update terhadap progress pengecoran tiang
    
    If TiangArray <= 0 Or PAUSE_GAME Then Exit Sub
    For intVar2 = 1 To TiangArray
        With TiangProg(intVar2)
            .bytLong = .bytLong + 1
            If .bytLong > 10 Then
            .bytLong = 0
            .bytCurrentProg = .bytCurrentProg + 1
            If .bytCurrentProg >= .bytProgress Then
                'maka progress tiang siap, dan tiang selesai dicor
                ArGedung(.bytX, .bytY).bytSprX = 5
                ArGedung(.bytX, .bytY).bytSprY = 3
                ArGedung(.bytX, .bytY).bTiang = True
                ArGedung(.bytX, .bytY).bCor = True
                Screen.Redraw = True
                RedrawMiniMap = True
            End If
            End If
        End With
    Next intVar2
        
    'cek array pertama bernilai nol atau tidak (FIFO)
    If TiangArray > 0 Then
    If TiangProg(1).bytCurrentProg >= 100 Then
        'copy array menjadi ubound-1
        For intVar2 = 1 To TiangArray - 1
            TiangProg(intVar2).bytX = TiangProg(intVar2 + 1).bytX
            TiangProg(intVar2).bytY = TiangProg(intVar2 + 1).bytY
            TiangProg(intVar2).bytCurrentProg = TiangProg(intVar2 + 1).bytCurrentProg
            TiangProg(intVar2).bytLong = TiangProg(intVar2 + 1).bytLong
            TiangProg(intVar2).bytProgress = TiangProg(intVar2 + 1).bytProgress
        Next intVar2
        'redim sekali lagi
        TiangArray = TiangArray - 1
    End If
    End If
    
End Sub

Sub Crane2Record()
With RecCrane
    .BaseX = Crane.BaseX
    .BaseY = Crane.BaseY
    .HeightCrane = Crane.HeightCrane
    .LengthCrane = Crane.LengthCrane

    'bagian cockpit crane
    .XCock = Crane.XCock
    .YCock = Crane.YCock
    .XCockpitOnArray = Crane.XCockpitOnArray
    .YCockpitOnArray = Crane.YCockpitOnArray

    'bagian kepala crane
    .XHead = Crane.XHead
    .YHead = Crane.YHead
    .YKait = Crane.YKait
    .LengthChain = Crane.LengthChain
    .CraneSpeed = Crane.CraneSpeed
    .XPositionHead = Crane.XPositionHead
    .YPositionHead = Crane.YPositionHead
    .XHeadSmooth = Crane.XHeadSmooth
    .WayMove = Crane.WayMove

    .MoveDuration = Crane.MoveDuration
    .CraneAction = Crane.CraneAction

    .LastPlaceCraneX = Crane.LastPlaceCraneX
    .LastPlaceCraneY = Crane.LastPlaceCraneY

    .CranePilih = Crane.CranePilih
    .Occupied = Crane.Occupied
    .LebarTiang = Crane.LebarTiang
    .XTiangOnArray = Crane.XTiangOnArray
    .YTiangOnArray = Crane.YTiangOnArray
    
    .PointerDraw = Crane.PointerDraw
    .PointerLong = Crane.PointerLong

    .PickStuff = Crane.PickStuff

    .IsCraneMoving = Crane.IsCraneMoving
    .XPoint = Crane.XPoint
    .YPoint = Crane.YPoint

    .StartMoving = Crane.StartMoving
End With
End Sub

Sub Cuaca2Record()
    With RecCuaca
        .Suhu = Cuaca.Suhu
        .Status = Cuaca.Status
        .YMove = Cuaca.YMove
        .Panjang = Cuaca.Panjang

        .JamMulai = Cuaca.JamMulai
        .MenitMulai = Cuaca.MenitMulai
        .JamSelesai = Cuaca.JamSelesai
        .MenitSelesai = Cuaca.MenitSelesai
        .JamSekarang = Cuaca.JamSekarang
        .MenitSekarang = Cuaca.MenitSekarang

        .HariIni = Cuaca.HariIni

        .Delayment = Cuaca.Delayment
        .EachFPS = Cuaca.EachFPS
    End With
End Sub

Private Sub DrawMainMenuGedung()
    'rutin untuk menggambar gedung ke dalam layar menu
    Dim XG As Byte
    Dim YG As Byte
    For XG = 1 To 20
    For YG = 0 To 1
        BackBuffer.BltFast (XG * 14) + 100, 329 - (YG * 20), Sprite, BoxRect(0, 0, 14, 20), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    Next YG
    Next XG
    For YG = 0 To 1
        BackBuffer.BltFast 394, 329 - (YG * 20), Sprite, BoxRect(70, 60, 84, 80), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        BackBuffer.BltFast 394, 269 - (YG * 20), Sprite, BoxRect(70, 60, 84, 80), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        BackBuffer.BltFast 113, 329 - (YG * 20), Sprite, BoxRect(70, 60, 84, 80), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        
    Next YG
    For YG = 0 To 3
        BackBuffer.BltFast 550, 329 - (YG * 20), Sprite, BoxRect(70, 60, 84, 80), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    Next YG
    
    For XG = 1 To 21
        BackBuffer.BltFast (XG * 14) + 100, 289, Sprite, BoxRect(0, 80, 14, 100), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    Next XG
    
End Sub

Public Sub DrawMoneyOnScreen(bytX As Byte, bytY As Byte, Money As String, Way As Byte)
    NoArray = NoArray + 1
    ReDim Preserve MoneyScreen(NoArray)
    MoneyScreen(NoArray).bytX = bytX
    MoneyScreen(NoArray).bytY = bytY
    MoneyScreen(NoArray).JumlahUang = Money
    MoneyScreen(NoArray).Arah = Way
    MoneyScreen(NoArray).bytLong = 30
End Sub

Public Sub ErrorBox(SError As String)
    'rutin ini akan menampilkan kotak Error berikut pesannya
    
    If ShowError.RedrawError Then
    
    ShowError.RedrawError = False
    
    'pindahkan backbuffer ke tempdx
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.BltFast 0, 0, BackBuffer, BoxRect(0, 0, 800, 600), DDBLTFAST_WAIT
    
    'bayangan
    TempDX.BltColorFill BoxRect(250, 250, 550, 330), QBColor(0)
    'kotak
    TempDX.DrawBox 245, 245, 545, 325
    TempDX.BltColorFill BoxRect(246, 246, 544, 324), RGB(10, 100, 100)
    TempDX.BltColorFill BoxRect(246, 246, 544, 260), RGB(50, 0, 150)
    
    With stFont
        .Name = "Comic Sans MS"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    'gambar OK
    TempDX.SetForeColor RGB(0, 0, 0)
    TempDX.DrawRoundedBox 360, 300, 430, 320, 4, 4
    TempDX.BltColorFill BoxRect(361, 301, 429, 319), RGB(15, 100, 15)
    TempDX.DrawText 385, 302, "O K", False
    
    With stFont
        .Name = "Comic Sans MS"
        .Size = 8
        .Bold = False
    End With
    TempDX.SetFont stFont
    
    TempDX.SetForeColor RGB(255, 255, 255)
    'tuliskan hasil kata ke layar
    TempDX.DrawText 260, 265, ShowError.StrKata, False
    'tuliskan titlebar
    TempDX.DrawText 330, 245, "Pesan Status Permainan", False
    
    End If
    
    BackBuffer.BltFast 245, 245, TempDX, BoxRect(245, 245, 550, 330), DDBLTFAST_WAIT
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 360 To 430         'tombol OK
        Select Case CursorY
        Case 300 To 320
            'tombol OK ditekan
            ShowError.ErrorWindow = False
            PAUSE_GAME = False
        End Select
    End Select
    End If
End Sub

Sub FillProgToTiang(bytX As Byte, bytY As Byte)
    TiangArray = TiangArray + 1
    TiangProg(TiangArray).bytX = bytX
    TiangProg(TiangArray).bytY = bytY
    TiangProg(TiangArray).bytCurrentProg = 0
    TiangProg(TiangArray).bytProgress = 100
End Sub

Public Function HitungSkor() As Integer
    Dim TmpSkor As Integer
    Dim JumlahBataObj As Long
    Dim JlhBata As Long
    Dim ArObjek(1 To MAXGEDUNGX, 1 To MAXGEDUNGY) As udtGedung
    TmpSkor = 0
    HitungSkor = TmpSkor
    
    'gunakan arobjek sebagai isi objektif
    'isikan objek bata
    Dim JlhX As Byte
    Dim JlhY As Byte
    For JlhX = 50 To 120
    For JlhY = 76 To 146
        ArObjek(JlhX, JlhY).bWall = True
    Next JlhY
    Next JlhX
    For JlhX = 80 To 120
    For JlhY = 36 To 75
        ArObjek(JlhX, JlhY).bWall = True
    Next JlhY
    Next JlhX
    For JlhX = 81 To 91
    For JlhY = 26 To 35
        ArObjek(JlhX, JlhY).bWall = True
    Next JlhY
    Next JlhX
    
    JumlahBataObj = (70 * 70) + (40 * 39) + (10 * 9)
    
    Dim JumlahJendela As Integer
    Dim ShouldJendela As Integer
    
    JlhBata = 0
    JumlahJendela = 0
    For JlhX = 1 To MAXGEDUNGX
    For JlhY = 1 To MAXGEDUNGY
        If ArGedung(JlhX, JlhY).bWall = ArObjek(JlhX, JlhY).bWall Then JlhBata = JlhBata + 1
        If ArGedung(JlhX, JlhY).bytSprX = 4 And ArGedung(JlhX, JlhY).bytSprY = 3 Then JumlahJendela = JumlahJendela + 1
    Next JlhY
    Next JlhX
    ShouldJendela = Int(JumlahBataObj * 0.01)
    
    'hitung jumlah persen
    Dim Persen As Single
    Dim PersenJendela As Single
    PersenJendela = (JumlahJendela / ShouldJendela) '* 100
    Persen = (JlhBata / JumlahBataObj) '* 100
    HitungSkor = Int(((PersenJendela + Persen) * 800) / 100)
End Function
Public Sub InitNameWorker()
    WorkerName(1) = "JANSEN"
    WorkerName(2) = "MICHAEL"
    WorkerName(3) = "OLSEN"
    WorkerName(4) = "HANS"
    WorkerName(5) = "MCCLANE"
    WorkerName(6) = "GRUBER"
    WorkerName(7) = "IRVIN"
    WorkerName(8) = "EDWARD"
    WorkerName(9) = "IRIS"
    WorkerName(10) = "CORTINO"
End Sub


Public Sub LoadGame()
    'rutin ini akan melakukan penyimpanan terhadap game
    Open App.Path & "\Save\Default.Jat" For Binary Access Read As #1
    
    '1.Peta permainan
    Get #1, 1, ArGedung
    
    '2.Jumlah Lemmings
    Get #1, , intLemm
    
    '3.Array Lemmings
    Get #1, , RecWorker
    
    'Jalankan konversi kelas Worker ke dalam Record
    Call Record2Worker
    
    '4.Workshop
    Get #1, , Workshop
    
    '5.Statusgame
    Get #1, , StatusGame
    
    '6.Graph
    Get #1, , Graph
    
    '7. Misc
    Dim bytXGedung As Integer
    Dim bytYGedung As Integer
    Dim bytXScroll As Integer
    Dim bytYScroll As Integer
    
    Get #1, , bytXGedung
    Get #1, , bytYGedung
    Get #1, , bytXScroll
    Get #1, , bytYScroll
    
    '8.Cuaca
    Get #1, , RecCuaca
    Call Record2Cuaca
    
    '9.Crane
    Get #1, , RecCrane
    Call Record2Crane
    
    '10. Objektif
    Get #1, , Objektif
    
    '11.TiangProg Progress dan tiangarray
    Get #1, , TiangArray
    Get #1, , TiangProg
    
    '12.Total Workshop
    Get #1, , TotalWorkshop
    Close #1
    
    Gedung.XGedung = bytXGedung
    Gedung.YGedung = bytYGedung
    Screen.XScroll = bytXScroll
    Screen.YScroll = bytYScroll
End Sub

Sub Main()
    'Jalankan menu utama terlebih dahulu
    StillRunning = True
    DXMod.Init
    
    Call InitNameWorker
    'Call ApplyObjective
    Call LoadDataSprites
    Call MainMenu
    
    StillRunning = True
    
    Call GameMod.SkalaMiniMap
    SelectedToolbar = 1
    MousePointer = MOUSE_DEFAULT
    Crane.XPoint = Crane.XHead
    Crane.YPoint = Crane.YHead
    PAUSE_GAME = False
    Screen.Redraw = True
    RedrawMiniMap = True
    Tampakbangunan = TAMPAK_DALAM
    Gedung.SwapArr2Ged
    
    ShowMap = True
    
    Cuaca.Delayment = 3
    
    'mulai jalankan musik dan suara
    If SFXMusik.InitDM Then
        If SFXMusik.LoadMusic(App.Path & "\Music\hotsteel.Mid") Then
            SFXMusik.PlayMusic
        End If
    End If
    
    Render
End Sub

Sub InitMainMenuWorker()
    Dim Xf As Byte
    'arah kiri kek kanan
    For Xf = 1 To 5
        Worker(Xf).Active = True
        Worker(Xf).Job = Walk
        Worker(Xf).Status = Walk
        Worker(Xf).WalkSpeed = Int(Rnd * 3) + 1
        Worker(Xf).XPosition = Int(Rnd * 5) + 50
        Worker(Xf).YPosition = 329
        Worker(Xf).WayMove = MOVE_RIGHT
    Next Xf
    'arah kanan ke kiri
    For Xf = 5 To 10
        Worker(Xf).Active = True
        Worker(Xf).Job = Walk
        Worker(Xf).Status = Walk
        Worker(Xf).WalkSpeed = Int(Rnd * 3) + 1
        Worker(Xf).XPosition = 730
        Worker(Xf).YPosition = 329
        Worker(Xf).WayMove = MOVE_LEFT
    Next Xf
    'bagian yang bekerja
    For Xf = 11 To 15
        Worker(Xf).Active = True
        Worker(Xf).Job = BUILDER
        Worker(Xf).Status = WORK
        Worker(Xf).XPosition = 300
        Worker(Xf).YPosition = 329
        Worker(Xf).WorkSpeed = 20
    Next Xf
    Worker(11).XPosition = 150
    Worker(12).Job = WELD_RIGHT
    Worker(12).XPosition = 500
    Worker(13).Job = WELD_LEFT
    Worker(13).XPosition = 350
    Worker(13).YPosition = 269
    Worker(14).YPosition = 269
    Worker(15).Job = DIG_LEFT
    Worker(15).XPosition = 600
End Sub
Sub MainMenu()
    'Menu Interface Awal Permainan
    'Init variable directdraw
    Dim LogoDX As DirectDrawSurface7
    Dim MenuDX(1 To 3) As DirectDrawSurface7
    
    'loading data interface
    If Not LoadSprite(App.Path & "\graphics\logo.bmp", 0, LogoDX) Then
        MsgBox "Error Loading Image !", vbExclamation
        Exit Sub
    End If
    If Not LoadSprite(App.Path & "\graphics\MenuBaru.bmp", 0, MenuDX(1)) Then
        MsgBox "Error Loading Image !", vbExclamation
        Exit Sub
    End If
    If Not LoadSprite(App.Path & "\graphics\MenuLanjut.bmp", 0, MenuDX(2)) Then
        MsgBox "Error Loading Image !", vbExclamation
        Exit Sub
    End If
    If Not LoadSprite(App.Path & "\graphics\MenuKeluar.bmp", 0, MenuDX(3)) Then
        MsgBox "Error Loading Image !", vbExclamation
        Exit Sub
    End If
    
    'later used to enhanced menu pointing
    InvertRGB = 20
    
    'isikan terlebih dahulu data di tempdx
    TempDX.BltColorFill BoxRect(0, 0, 800, 600), QBColor(0)
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.BltFast 100, 50, LogoDX, BoxRect(0, 0, 600, 31), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    TempDX.BltFast 300, 400, MenuDX(1), BoxRect(0, 0, 200, 12), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    TempDX.BltFast 300, 450, MenuDX(2), BoxRect(0, 0, 200, 12), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    TempDX.BltFast 300, 500, MenuDX(3), BoxRect(0, 0, 200, 12), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    TempDX.DrawBox 50, 100, 750, 350
    TempDX.BltColorFill BoxRect(51, 101, 749, 349), RGB(60, 50, 50)
    TempDX.DrawText 650, 550, "Ver 1.05.04.04", False
    
    Call InitMainMenuWorker
    
    Do While StillRunning
        BackBuffer.BltFast 0, 0, TempDX, BoxRect(0, 0, 800, 600), DDBLTFAST_WAIT
        
        Call DrawMainMenuGedung
        Call UpdateMainMenuWorker
        BackBuffer.SetForeColor RGB(255, 255, 255)
        
        Call HandleMenuMouse

        Call HandleMenuKeys
        
        'Show Mouse
        BackBuffer.BltFast CursorX, CursorY, TexMouse, BoxRect(0, 0, 15, 15), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Primary.Flip Nothing, DDFLIP_WAIT
        FPS
        DoEvents
    Loop
End Sub


Sub NewGame()
    'rutin ini akan melakukan inisialisasi permainan baru
    Gedung.InitGedung

    Call InitWorker
    Workshop.intJendela = 30
    Workshop.intPintu = 5
    Workshop.intTangga = 15
    
    Call GameMod.InitHargaBahan
    Call Cuaca.Initialize
    Cuaca.Delayment = 15
    Graph(1).StrBln = Format(Cuaca.HariIni, "mmm")
End Sub
Sub Record2Crane()
With Crane
    .BaseX = RecCrane.BaseX
    .BaseY = RecCrane.BaseY
    .HeightCrane = RecCrane.HeightCrane
    .LengthCrane = RecCrane.LengthCrane

    'bagian cockpit crane
    .XCock = RecCrane.XCock
    .YCock = RecCrane.YCock
    .XCockpitOnArray = RecCrane.XCockpitOnArray
    .YCockpitOnArray = RecCrane.YCockpitOnArray

    'bagian kepala crane
    .XHead = RecCrane.XHead
    .YHead = RecCrane.YHead
    .YKait = RecCrane.YKait
    .LengthChain = RecCrane.LengthChain
    .CraneSpeed = RecCrane.CraneSpeed
    .XPositionHead = RecCrane.XPositionHead
    .YPositionHead = RecCrane.YPositionHead
    .XHeadSmooth = RecCrane.XHeadSmooth
    .WayMove = RecCrane.WayMove

    .MoveDuration = RecCrane.MoveDuration
    .CraneAction = RecCrane.CraneAction

    .LastPlaceCraneX = RecCrane.LastPlaceCraneX
    .LastPlaceCraneY = RecCrane.LastPlaceCraneY

    .CranePilih = RecCrane.CranePilih
    .Occupied = RecCrane.Occupied
    .LebarTiang = RecCrane.LebarTiang
    .XTiangOnArray = RecCrane.XTiangOnArray
    .YTiangOnArray = RecCrane.YTiangOnArray
    
    .PointerDraw = RecCrane.PointerDraw
    .PointerLong = RecCrane.PointerLong

    .PickStuff = RecCrane.PickStuff

    .IsCraneMoving = RecCrane.IsCraneMoving
    .XPoint = RecCrane.XPoint
    .YPoint = RecCrane.YPoint

    .StartMoving = RecCrane.StartMoving
End With

End Sub

Sub Record2Cuaca()
    With Cuaca
        .Suhu = RecCuaca.Suhu
        .Status = RecCuaca.Status
        .YMove = RecCuaca.YMove
        .Panjang = RecCuaca.Panjang

        .JamMulai = RecCuaca.JamMulai
        .MenitMulai = RecCuaca.MenitMulai
        .JamSelesai = RecCuaca.JamSelesai
        .MenitSelesai = RecCuaca.MenitSelesai
        .JamSekarang = RecCuaca.JamSekarang
        .MenitSekarang = RecCuaca.MenitSekarang

        .HariIni = RecCuaca.HariIni

        .Delayment = RecCuaca.Delayment
        .EachFPS = RecCuaca.EachFPS
    End With
End Sub

Sub Record2Worker()
    'rutin ini untuk melakukan konversi ke dalam record
    Dim intI As Integer
    For intI = 1 To 1000
        With Worker(intI)
            .Active = RecWorker(intI).Active
            .NoLemm = RecWorker(intI).NoLemm
            .Nama = RecWorker(intI).Nama
            .Honor = RecWorker(intI).Honor
            .Stability = RecWorker(intI).Stability
            .Tolerance = RecWorker(intI).Tolerance
            .Progress = RecWorker(intI).Progress
            .Frame = RecWorker(intI).Frame
            .Way = RecWorker(intI).Way
            .Job = RecWorker(intI).Job
            .JobX = RecWorker(intI).JobX
            .JobY = RecWorker(intI).JobY
            .Status = RecWorker(intI).Status
            .XPosition = RecWorker(intI).XPosition
            .YPosition = RecWorker(intI).YPosition
            .Perubah = RecWorker(intI).Perubah
            .XTileOnArray = RecWorker(intI).XTileOnArray
            .YTileOnArray = RecWorker(intI).YTileOnArray
            .WorkSpeed = RecWorker(intI).WorkSpeed
            .WorkNow = RecWorker(intI).WorkNow
            .CurrentSpeed = RecWorker(intI).CurrentSpeed

            .WalkSpeed = RecWorker(intI).WalkSpeed
            .IsMoving = RecWorker(intI).IsMoving
            .WayMove = RecWorker(intI).WayMove
            .SearchMove = RecWorker(intI).SearchMove
            .MoveDuration = RecWorker(intI).MoveDuration
            .XHeadSmooth = RecWorker(intI).XHeadSmooth
            .YHeadSmooth = RecWorker(intI).YHeadSmooth

            .XPoint = RecWorker(intI).XPoint
            .YPoint = RecWorker(intI).YPoint
        End With
    Next intI
End Sub

Sub Render()
    
    Do While StillRunning
        'TempDX.BltColorFill BoxRect(0, 0, 800, 600), QBColor(0)
        'BackBuffer.BltColorFill BoxRect(0, 500, 800, 600), QBColor(0)
        
        'update layar
        Screen.DrawGedung Gedung.XGedung, Gedung.YGedung, Screen.XScroll, Screen.YScroll
        
        'If Cuaca.Status = HUJAN Then
        '    Call UpdateHujan
        'End If
        
        'bagian updating Crane
        Call GameMod.CheckCraneMovement
        Call Crane.DrawHead
        Call GameMod.UpdateCraneCockpit
        
        'bagian updating worker terhadap posisi
        Call GameMod.UpdateWorker
        
        'bagian HUD permainan
        Call GameMod.DrawStatus
        Call GameMod.UpdateStatusBar
        
        'bagian pengatur cuaca dan waktu
        Call GameMod.AturCuacadanWaktu
        
        If ShowMap Then
            Call GameMod.RefreshMiniMap
            Call GameMod.UpdateMiniMap
        End If
        
        Call WriteMoney
        Call CheckTiangProg
        Call GameMod.DrawToolBar
        
        'Tampilkan informasi permainan
        Call ShowInformation

        'bagian pelacak window on atau off
        If WindowMaker = True Then
            ShowWindowMaker
        ElseIf WindowBelanjaBata = True Then
            ShowBelanjaBata
        ElseIf WindowBelanjaPasir = True Then
            ShowBelanjaPasir
        ElseIf WindowBelanjaSemen = True Then
            ShowBelanjaSemen
        ElseIf WindowBelanjaCat = True Then
            ShowBelanjaCat
        ElseIf WindowBelanjaKayu = True Then
            ShowBelanjaKayu
        ElseIf WindowBelanjaBesi = True Then
            ShowBelanjaBesi
        ElseIf WindowBelanjaKaca = True Then
            ShowBelanjaKaca
        ElseIf WindowGraph = True Then
            ShowWindowGraph
        ElseIf WindowSketch = True Then
            ShowSketch Objektif.dSketsa
        ElseIf WindowTiang = True Then
            ShowWindowTiang
        ElseIf WindowHelp = True Then
            ShowWindowHelp
        ElseIf WindowTips = True Then
            ShowWindowTips
        ElseIf WindowWorkShop = True Then
            ShowWindowWorkshop
        'ElseIf WindowSplash = True Then
        '    ShowSplashScreen
        ElseIf ShowError.ErrorWindow = True Then
            Call ErrorBox(ShowError.StrKata)
        End If
        
        'bagian pengatur handling mouse dan keyboard
        Call HandleMouse
        Call HandleKeys
        
        'bagian pengatur penggulungan layar
        Call Screen.CheckScroll
        
        If GlobalSel And SelNo > 0 Then
            'non aktifkan crane
            Crane.CranePilih = False
            'maka tulis identitas worker yang terpilih
            Call WriteWorkerID(SelNo)
            If Mouse_Button1 = True Then Mouse_Button1 = False
        End If
        
        If Crane.CranePilih And Not GlobalSel Then
            Call WriteCraneID
        End If
        
        'apakah pointer perlu digambar
        'With Crane
        'If .PointerDraw Then
        '    If .PointerLong < 10 Then
        '        .PointerDraw = False
        '    Else
        '        .PointerLong = .PointerLong - 1
        '        BackBuffer.BltFast ((.XPoint - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll, ((.YPoint - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll, TexMouse, BoxRect(45, 0, 59, 14), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        '    End If
        'End If
        'End With
        
        BackBuffer.SetForeColor RGB(0, 0, 0)
        'tampilkan sprite batu bata
        If CursorY < 520 Then
        If PASANGBATU Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(0, 0, 13, 19), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf PasangTiang Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(56, 40, 69, 59), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf PasangTangga Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(70, 40, 84, 59), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf Penghancur Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(42, 60, 55, 79), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf PasangJendela Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(56, 60, 69, 79), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        ElseIf PasangCat Then
            BackBuffer.BltFast ((CursorX \ TILEWIDTH) * TILEWIDTH) + Screen.XScroll, ((CursorY \ TILEHEIGHT) * TILEHEIGHT) + Screen.YScroll, Sprite, BoxRect(70, 80, 84, 99), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
        End If
        End If
        
        BackBuffer.BltFast CursorX, CursorY, TexMouse, BoxRect(MousePointer * 15, 0, (MousePointer * 15) + 14, 15), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        
        'BackBuffer.DrawText 100, 300, ArGedung(50, 144).bytSprX, False
        'BackBuffer.DrawText 100, 315, TmpSpr(50, 144).bytSprX, False
        'BackBuffer.DrawText 100, 300, Worker(5).YTileOnArray, False
        'BackBuffer.DrawText 200, 300, ArGedung(Screen.GetTileX, Screen.GetTileY).bLadder, False
        'BackBuffer.DrawText 300, 300, ArGedung(Screen.GetTileX, Screen.GetTileY).bGround, False
        'BackBuffer.DrawText 100, 200, Worker(1).WalkSpeed, False
        'BackBuffer.DrawText 200, 200, Worker(1).SearchMove, False
        'BackBuffer.DrawText 200, 300, Worker(1).WayMove, False
        'BackBuffer.DrawText 300, 300, ArGedung(Screen.GetTileX, Screen.GetTileY).bWalkable, False
        
        Primary.Flip Nothing, DDFLIP_WAIT
        FPS
        DoEvents
    Loop
End Sub


Sub HandleKeys()
    If Not Keyboard.SuccessAcquire Then Exit Sub
    If Keyboard.KeyStatus(DIK_LEFT) Then
        'gulung kiri
        ScrollScreen = True
        ScrollWay = SCROLL_LEFT
        Screen.Redraw = True
        Call Screen.CheckScroll
    ElseIf Keyboard.KeyStatus(DIK_RIGHT) Then
        ScrollScreen = True
        ScrollWay = SCROLL_RIGHT
        Screen.Redraw = True
        Call Screen.CheckScroll
    ElseIf Keyboard.KeyStatus(DIK_UP) Then
        ScrollScreen = True
        ScrollWay = SCROLL_UP
        Screen.Redraw = True
        Call Screen.CheckScroll
    ElseIf Keyboard.KeyStatus(DIK_DOWN) Then
        ScrollScreen = True
        ScrollWay = SCROLL_DOWN
        Screen.Redraw = True
        Call Screen.CheckScroll
    End If
    
    If Keyboard.KeyStatus(DIK_M) Then
        ShowMap = Not ShowMap
        Keyboard.ClearDI
    'pengaturan tampak dari bangunan
    ElseIf Keyboard.KeyStatus(DIK_1) Then
        If Tampakbangunan = TAMPAK_DALAM Then Exit Sub
        Keyboard.ClearDI
        Tampakbangunan = TAMPAK_DALAM
        PAUSE_GAME = False
        Call Gedung.ViewGedung(Tampakbangunan)
        Screen.Redraw = True
    ElseIf Keyboard.KeyStatus(DIK_2) Then
        If Tampakbangunan = TAMPAK_LUAR Then Exit Sub
        Keyboard.ClearDI
        Tampakbangunan = TAMPAK_LUAR
        PAUSE_GAME = True
        Call Gedung.ViewGedung(Tampakbangunan)
        Screen.Redraw = True
    ElseIf Keyboard.KeyStatus(DIK_SPACE) Then       'bagian pausing game
        'maka hentikan game sejenak
        If Tampakbangunan <> TAMPAK_LUAR Then
            PAUSE_GAME = Not PAUSE_GAME
            Keyboard.ClearDI
        End If
    ElseIf Keyboard.KeyStatus(DIK_O) Then
        If Not Crane.CranePilih Then
            Crane.CranePilih = True
            GlobalSel = False
        End If
    ElseIf Keyboard.KeyStatus(DIK_S) Then
        RedrawWindowSplash = True
        WindowSplash = True
        PAUSE_GAME = True
    ElseIf Keyboard.KeyStatus(DIK_W) Then
        RedrawWindowWorkshop = True
        WindowWorkShop = True
        PAUSE_GAME = True
    End If
    
    'Giliran Crane terpilih
    If Crane.CranePilih And Not GlobalSel Then
        'cek apakah tombol yang ditekan
        'Keyboard.ClearDI
        If Keyboard.KeyStatus(DIK_EQUALS) Then
            'naikkan crane ke atas
            Crane.RaiseBase
            Keyboard.ClearDI
            Screen.Redraw = True
            RedrawMiniMap = True
        ElseIf Keyboard.KeyStatus(DIK_MINUS) Then
            'turunkan crane ke bawah
            Crane.LowerBase
            Keyboard.ClearDI
            Screen.Redraw = True
            RedrawMiniMap = True
        ElseIf Keyboard.KeyStatus(DIK_RBRACKET) Then
            'panjangkan tangan crane
            Crane.LengthenArm
            Keyboard.ClearDI
            Screen.Redraw = True
            RedrawMiniMap = True
        ElseIf Keyboard.KeyStatus(DIK_LBRACKET) Then
            'pendekkan tangan crane
            Crane.ShortenArm
            Keyboard.ClearDI
            Screen.Redraw = True
            RedrawMiniMap = True
        ElseIf Keyboard.KeyStatus(DIK_P) Then
            'memasangkan tiang vertikal ke daerah destination
            If Crane.Occupied And Crane.TiangLetak Then    'maka tiang bisa diletakkan
                Crane.GlueTiang
                Crane.Occupied = False
                If Crane.LengthChain > 1 Then Crane.LengthChain = Crane.LengthChain - 1
                Screen.Redraw = True
                RedrawMiniMap = True
            End If
        End If
    End If
    
    'bagian pengontrol bunga pinjaman
    If WindowBudget Then
        With StatusGame
        'cek tombol pageup dan pagedown
        If Keyboard.KeyStatus(DIK_PAGEUP) Then
            Keyboard.ClearDI
            .Pinjaman = .Pinjaman + 10000
            RedrawWindowBudget = True
        End If
        
        If Keyboard.KeyStatus(DIK_PAGEDOWN) And .Pinjaman > 0 Then
            Keyboard.ClearDI
            .Pinjaman = .Pinjaman - 10000
            RedrawWindowBudget = True
        End If
        End With
    End If
    
    'bagian add/minus lebar tiang
    If WindowTiang Then
        'cek tombol pageup dan pagedown
        If Keyboard.KeyStatus(DIK_PAGEUP) And Crane.LebarTiang < 12 Then
            Keyboard.ClearDI
            Crane.LebarTiang = Crane.LebarTiang + 1
            RedrawWindowTiang = True
        End If
        
        If Keyboard.KeyStatus(DIK_PAGEDOWN) And Crane.LebarTiang > 4 Then
            Keyboard.ClearDI
            Crane.LebarTiang = Crane.LebarTiang - 1
            RedrawWindowTiang = True
        End If
        
    End If
    
    'bagian pembelanjaan
    With Workshop
    If WindowBelanjaBata Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpBata = .bytTmpBata + 1
            RedrawWindowBata = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpBata > 0 Then .bytTmpBata = .bytTmpBata - 1
            RedrawWindowBata = True
        End If
    ElseIf WindowBelanjaPasir Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpPasir = .bytTmpPasir + 1
            RedrawWindowPasir = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpPasir > 0 Then .bytTmpPasir = .bytTmpPasir - 1
            RedrawWindowPasir = True
        End If
    ElseIf WindowBelanjaSemen Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpSemen = .bytTmpSemen + 1
            RedrawWindowSemen = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpSemen > 0 Then .bytTmpSemen = .bytTmpSemen - 1
            RedrawWindowSemen = True
        End If
    ElseIf WindowBelanjaKayu Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpKayu = .bytTmpKayu + 1
            RedrawWindowKayu = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpKayu > 0 Then .bytTmpKayu = .bytTmpKayu - 1
            RedrawWindowKayu = True
        End If
    ElseIf WindowBelanjaBesi Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpBesi = .bytTmpBesi + 1
            RedrawWindowBesi = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpBesi > 0 Then .bytTmpBesi = .bytTmpBesi - 1
            RedrawWindowBesi = True
        End If
    ElseIf WindowBelanjaKaca Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpKaca = .bytTmpKaca + 1
            RedrawWindowKaca = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpKaca > 0 Then .bytTmpKaca = .bytTmpKaca - 1
            RedrawWindowKaca = True
        End If
    ElseIf WindowBelanjaCat Then
        If Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            .bytTmpCat = .bytTmpCat + 1
            RedrawWindowCat = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If .bytTmpCat > 0 Then .bytTmpCat = .bytTmpCat - 1
            RedrawWindowCat = True
        End If
    ElseIf WindowDanaPekerja Then
        If Keyboard.KeyStatus(DIK_PAGEUP) Then
            Keyboard.ClearDI
            If SelNoDana > 1 Then SelNoDana = SelNoDana - 1
            RedrawWindowDana = True
        ElseIf Keyboard.KeyStatus(DIK_PAGEDOWN) Then
            Keyboard.ClearDI
            If SelNoDana < intLemm Then SelNoDana = SelNoDana + 1
            RedrawWindowDana = True
        ElseIf Keyboard.KeyStatus(DIK_ADD) Then
            Keyboard.ClearDI
            Worker(SelNoDana).Honor = Worker(SelNoDana).Honor + 100
            RedrawWindowDana = True
        ElseIf Keyboard.KeyStatus(DIK_SUBTRACT) Then
            Keyboard.ClearDI
            If Worker(SelNoDana).Honor > 0 Then Worker(SelNoDana).Honor = Worker(SelNoDana).Honor + 100
            RedrawWindowDana = True
        End If
    End If
    End With
End Sub

Sub HandleMenuKeys()
    'DxInput.CheckKeys
    'If DxInput.aKeys(DIK_DOWN) Then
    '    ScrollScreen = True
    '    ScrollWay = SCROLL_DOWN
    '    Call Screen.CheckScroll
    'End If
End Sub

Function RndName() As String
    'create a name randomly
    Dim intName As Byte
    Randomize
    intName = Int(Rnd * 10) + 1
    RndName = WorkerName(intName)
End Function

Sub RubahWarnaBox(Color As Byte)
    Select Case Color
    Case 1
        RGBColor.R = RGBColor.R + InvertRGB
        RGBColor.G = 0
        RGBColor.B = 0
        If RGBColor.R > 254 Or RGBColor.R < 1 Then
            InvertRGB = InvertRGB * -1
            RGBColor.R = 0
        End If
    Case 2
        RGBColor.G = RGBColor.G + InvertRGB
        RGBColor.R = 0
        RGBColor.B = 0
        If RGBColor.G > 254 Or RGBColor.G < 1 Then
            InvertRGB = InvertRGB * -1
            RGBColor.G = 0
        End If
    Case 3
        RGBColor.B = RGBColor.B + InvertRGB
        RGBColor.R = 0
        RGBColor.G = 0
        If RGBColor.B > 254 Or RGBColor.B < 1 Then
            InvertRGB = InvertRGB * -1
            RGBColor.B = 0
        End If
    End Select
    BackBuffer.SetForeColor RGB(RGBColor.R, RGBColor.G, RGBColor.B)
End Sub



Sub LoadDataSprites()
    LoadSprite App.Path & "\graphics\Worker.bmp", 0, LemmDX
    LoadSprite App.Path & "\graphics\pointers.bmp", 0, TexMouse
    LoadSprite App.Path & "\graphics\sprite.bmp", 0, Sprite
    LoadSprite App.Path & "\graphics\box.bmp", 0, BoxStatus
    LoadSprite App.Path & "\graphics\cuaca.bmp", RGB(255, 255, 255), CuacaDX
    LoadSprite App.Path & "\graphics\tool.bmp", 0, ToolBarDX
    LoadSprite App.Path & "\graphics\MukaWorker.bmp", 0, IconWorker
    'LoadSprite App.Path & "\graphics\Hujan.bmp", 0, HujanDX
    'LoadSprite App.Path & "\graphics\spriteHujan.bmp", 0, SpriteHujan
    
    Dim Xf As Byte
    'Init Frame Posisi Berjalan Worker
    For Xf = 0 To 7
        ArSpr(Xf, WALK_LEFT) = BoxRect(Xf * 14, 0, (Xf + 1) * 14, 19)
        ArSpr(Xf, WALK_RIGHT) = BoxRect(Xf * 14, 20, (Xf + 1) * 14, 39)
        ArSpr(Xf, PANIC) = BoxRect(Xf * 14, 40, (Xf + 1) * 14, 59)
        ArSpr(Xf, DIG_LEFT) = BoxRect(Xf * 14, 60, (Xf + 1) * 14, 79)
        ArSpr(Xf, DIG_RIGHT) = BoxRect(Xf * 14, 80, (Xf + 1) * 14, 99)
        ArSpr(Xf, PUSH_LEFT) = BoxRect(Xf * 14, 100, (Xf + 1) * 14, 119)
        ArSpr(Xf, PUSH_RIGHT) = BoxRect(Xf * 14, 120, (Xf + 1) * 14, 139)
        ArSpr(Xf, STAND_LEFT) = BoxRect(Xf * 14, 140, (Xf + 1) * 14, 159)
        ArSpr(Xf, STAND_RIGHT) = BoxRect(Xf * 14, 160, (Xf + 1) * 14, 179)
        ArSpr(Xf, WELD_LEFT) = BoxRect(Xf * 14, 180, (Xf + 1) * 14, 199)
        ArSpr(Xf, WELD_RIGHT) = BoxRect(Xf * 14, 200, (Xf + 1) * 14, 219)
        ArSpr(Xf, UP_DOWN) = BoxRect(Xf * 14, 220, (Xf + 1) * 14, 239)
        ArSpr(Xf, BUILDER) = BoxRect(Xf * 14, 240, (Xf + 1) * 14, 259)
    Next Xf
End Sub

Sub HandleMouse()
    'mouse digunakan untuk melakukan kontrol terhadap permainan
    'baik itu penggulungan dan sebagainya, dengan nilai cursorX/Y
    ScrollScreen = False
    ScrollWay = SCROLL_NONE
    MousePointer = MOUSE_DEFAULT
    
    If CursorX <= 0 Or CursorX >= 775 Or (CursorY >= 500 And CursorY <= 520) Or CursorY <= 0 Then
    
    'then check for which way the screen should scroll
    Select Case CursorX
    Case Is <= 0
        ScrollWay = SCROLL_LEFT
        ScrollScreen = True
        MousePointer = MOUSE_SCROLL_LEFT
        Screen.Redraw = True
    Case Is >= 775
        ScrollWay = SCROLL_RIGHT
        ScrollScreen = True
        MousePointer = MOUSE_SCROLL_RIGHT
        Screen.Redraw = True
    End Select
    Select Case CursorY
    Case 500 To 520
        ScrollWay = SCROLL_DOWN
        ScrollScreen = True
        MousePointer = MOUSE_SCROLL_DOWN
        Screen.Redraw = True
    Case Is <= 0
        ScrollWay = SCROLL_UP
        ScrollScreen = True
        MousePointer = MOUSE_SCROLL_UP
        Screen.Redraw = True
    End Select
    
    
    Else
        'this part is not for scrolling stuff
        If CursorY >= 520 Then
            'we are in toolbar zone
            Call GameMod.HandleToolbarZone
        Else
            'we are in game/screen zone
            Call GameMod.HandleGameZone
        End If
    End If
    
End Sub

Sub HandleMenuMouse()
    'Subroutine untuk melacak pemindahan mouse ke objek menu
    'nilai cursorX dan cursorY
    Dim MenuPilih As Byte
    
    MenuPilih = 0
    
    Select Case CursorX
    Case 301 To 499
        Select Case CursorY
        Case 401 To 413 'menu pertama baru
            'Call RubahWarnaBox(1)
            BackBuffer.DrawBox 290, 390, 510, 423
            MenuPilih = 1
        Case 451 To 463
            'Call RubahWarnaBox(2)
            BackBuffer.DrawBox 290, 440, 510, 473
            MenuPilih = 2
        Case 501 To 513
            'Call RubahWarnaBox(3)
            BackBuffer.DrawBox 290, 490, 510, 523
            MenuPilih = 3
        End Select
    End Select
    
    If Mouse_Button0 Then   'mouse tombol kiri tertekan
    Select Case MenuPilih
    Case 1  'Permainan Baru
        Call ApplyObjective
        StillRunning = False
        Call NewGame
        Mouse_Button0 = False
    Case 2  'Lanjut Permainan
        'mulai load permainan dari savegame
        Call LoadGame
        StillRunning = False
        Mouse_Button0 = False
    Case 3  'Keluar
        End
    End Select
    End If
End Sub


Public Sub SaveGame()
    'rutin ini akan melakukan penyimpanan terhadap game
    Open App.Path & "\Save\Default.Jat" For Binary Access Write As #1
    
    '1.Peta permainan
    Put #1, 1, ArGedung
    
    '2.Jumlah Lemmings
    Put #1, , intLemm
    
    'Jalankan konversi kelas Worker ke dalam Record
    Call Worker2Record
    
    '3.Array Lemmings
    Put #1, , RecWorker
    
    '4.Workshop
    Put #1, , Workshop
    
    '5.Statusgame
    Put #1, , StatusGame
    
    '6.Graph
    Put #1, , Graph
    
    '7.Misc
    Put #1, , Gedung.XGedung
    Put #1, , Gedung.YGedung
    Put #1, , Screen.XScroll
    Put #1, , Screen.YScroll
    
    '8.Cuaca
    Call Cuaca2Record
    Put #1, , RecCuaca
    
    '9.Variabel Crane
    Call Crane2Record
    Put #1, , RecCrane
    
    '10.Objektif permainan
    Put #1, , Objektif
    
    '11.Artiang Progress dan TiangArray
    Put #1, , TiangArray
    Put #1, , TiangProg
    
    '12.Total Workshop (added later to fix bug)
    Put #1, , TotalWorkshop
    Close #1
End Sub


Sub ShowInformation()
    'rutin ini untuk menampilkan informasi permainan
    'yang akan ditampilkan berupa
    
    With stFont
        .Name = "Comic Sans MS"
        .Size = 9
    End With
    BackBuffer.SetFont stFont
    BackBuffer.SetForeColor RGB(0, 0, 0)
    '1.Posisi Mouse terhadap permainan
    BackBuffer.DrawText 5, 76, "Position : " & Screen.GetTileX & "," & Screen.GetTileY, False
    '2.Frame Per Second
    BackBuffer.DrawText 5, 91, "FPS : " & mintFPS, False
    '3.Tampak Bangunan
    BackBuffer.DrawText 5, 106, "View : " & IIf(Tampakbangunan = 0, "Outside View", "Inside View"), False
    '4. Jumlah Pekerja
    BackBuffer.DrawText 5, 121, "Worker : " & intLemm, False
    
    BackBuffer.SetForeColor RGB(150, 150, 150)
    '1.Posisi Mouse terhadap permainan
    BackBuffer.DrawText 4, 75, "Position : " & Screen.GetTileX & "," & Screen.GetTileY, False
    '2.Frame Per Second
    BackBuffer.DrawText 4, 90, "FPS : " & mintFPS, False
    '3.Tampak Bangunan
    BackBuffer.DrawText 4, 105, "View : " & IIf(Tampakbangunan = 0, "Outside View", "Inside View"), False
    '4.Jumlah Pekerja
    BackBuffer.DrawText 4, 120, "Worker : " & intLemm, False
End Sub

Sub Temp()
        BackBuffer.DrawText 600, 0, ScrollScreen, False
        BackBuffer.DrawText 600, 10, SelNo, False
        BackBuffer.DrawText 600, 20, intLemm, False
        BackBuffer.DrawText 600, 30, Crane.Visible, False
        BackBuffer.DrawText 600, 40, Crane.CranePilih, False
        BackBuffer.DrawText 600, 30, Worker(SelNo).XTileOnArray, False
        BackBuffer.DrawText 600, 50, mintFPS, False
        BackBuffer.DrawText 400, 0, Crane.XPositionHead, False
        BackBuffer.DrawText 400, 10, Crane.YPositionHead, False
        BackBuffer.DrawText 400, 20, Crane.XHead, False
        BackBuffer.DrawText 400, 30, Crane.YHead + Crane.LengthChain, False
        BackBuffer.DrawText 400, 40, Crane.VisibleHead(Crane.XHead, Crane.YHead), False
        BackBuffer.DrawText 400, 60, Crane.IsArrivedX, False
        BackBuffer.DrawText 400, 50, Crane.XTiangOnArray, False
        BackBuffer.DrawText 400, 70, Worker(1).XTileOnArray, False
        BackBuffer.DrawText 400, 80, Worker(1).YTileOnArray, False
        BackBuffer.DrawText 400, 90, Worker(1).Status, False
        BackBuffer.DrawText 400, 100, Worker(1).Visible, False
        BackBuffer.DrawText 400, 110, Worker(1).XPosition, False
        BackBuffer.DrawText 400, 120, Worker(1).YPosition, False
        BackBuffer.DrawText 400, 130, Crane.IsArrivedX, False
        BackBuffer.DrawText 400, 150, Crane.IsArrivedY, False
        
        BackBuffer.DrawText 500, 0, Gedung.XGedung, False
        BackBuffer.DrawText 500, 20, Gedung.YGedung, False
        BackBuffer.DrawText 500, 30, Crane.XPositionHead, False
        BackBuffer.DrawText 500, 40, Crane.YPositionHead, False
        BackBuffer.DrawText 500, 50, Screen.GetTileX, False
        BackBuffer.DrawText 500, 60, Screen.GetTileY, False
        BackBuffer.DrawText 500, 70, Crane.PickCockpit, False
        BackBuffer.DrawText 0, 0, CursorX, False
        BackBuffer.DrawText 100, 0, CursorY, False
        BackBuffer.DrawText 150, 0, Screen.XScroll, False
        BackBuffer.DrawText 150, 100, Screen.YScroll, False
End Sub

Sub UpdateHujan()
    Dim intvarx As Byte
    Dim intvary As Byte
    For intvarx = 0 To 10
    For intvary = 0 To 13
        If intvary Mod 2 = 0 Then
            BackBuffer.BltFast (intvarx * 79) + 40, (intvary * 40) + InvHujan, CuacaDX, BoxRect(0, 0, 1, 4), DDBLTFAST_WAIT
        Else
            BackBuffer.BltFast intvarx * 79, (intvary * 40) + InvHujan, CuacaDX, BoxRect(0, 0, 1, 4), DDBLTFAST_WAIT
        End If
    Next intvary
    Next intvarx
    
    If HujanFPS > 5 Then
        'inversi nilai penurunan hujan
        HujanFPS = 0
        InvHujan = InvHujan + 2
        If InvHujan > 79 Then InvHujan = 0
    Else
        HujanFPS = HujanFPS + 1
    End If
End Sub

Sub UpdateMainMenuWorker()
    'rutin ini untuk melakukan tampilan para worker
    'di main menu waktu pilihan
    Dim bytX As Byte
    
    For bytX = 1 To 15
        With Worker(bytX)
        'jika worker nampak maka update posisi dan frame worker
        If .Active Then
            .UpdateFrame
        End If
        
        If .XPosition > 736 Or .XPosition < 50 Then
            If .WayMove = MOVE_RIGHT Then
                .WayMove = MOVE_LEFT
            Else
                .WayMove = MOVE_RIGHT
            End If
        End If
        
        If .WayMove = MOVE_RIGHT Then
            .XPosition = .XPosition + .WalkSpeed
        ElseIf .WayMove = MOVE_LEFT Then
            .XPosition = .XPosition - .WalkSpeed
        End If
        
        Select Case .Status
        Case Walk
            '.UpdatePosition
            If .Visible Then
            Select Case .WayMove
            Case MOVE_RIGHT
                BackBuffer.BltFast .XPosition, .YPosition, LemmDX, ArSpr(.Frame, WALK_LEFT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Case MOVE_LEFT
                BackBuffer.BltFast .XPosition, .YPosition, LemmDX, ArSpr(.Frame, WALK_RIGHT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End Select
            End If
        Case WORK
            BackBuffer.BltFast .XPosition, .YPosition, LemmDX, ArSpr(.Frame, .Job), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        End Select
        End With
    Next bytX
End Sub
Public Function VisibleMoney(X As Byte, Y As Byte) As Boolean
    VisibleMoney = False
    If ((X + MAXTILEX > Gedung.XGedung) And (X > Gedung.XGedung)) And ((Y < Gedung.YGedung + MAXTILEY) And (Y > Gedung.YGedung)) Then
        VisibleMoney = True
    End If
End Function

Function VisiblePointer(intX As Integer) As Boolean
    If Worker(intX).Visible And Worker(intX).YPosition < 520 Then
        VisiblePointer = True
    Else
        VisiblePointer = False
    End If
End Function


Sub Worker2Record()
    'rutin ini untuk melakukan konversi ke dalam record
    Dim intI As Integer
    For intI = 1 To 1000
        With RecWorker(intI)
            .Active = Worker(intI).Active
            .NoLemm = Worker(intI).NoLemm
            .Nama = Worker(intI).Nama
            .Honor = Worker(intI).Honor
            .Stability = Worker(intI).Stability
            .Tolerance = Worker(intI).Tolerance
            .Progress = Worker(intI).Progress
            .Frame = Worker(intI).Frame
            .Way = Worker(intI).Way
            .Job = Worker(intI).Job
            .JobX = Worker(intI).JobX
            .JobY = Worker(intI).JobY
            .Status = Worker(intI).Status
            .XPosition = Worker(intI).XPosition
            .YPosition = Worker(intI).YPosition
            .Perubah = Worker(intI).Perubah
            .XTileOnArray = Worker(intI).XTileOnArray
            .YTileOnArray = Worker(intI).YTileOnArray
            .WorkSpeed = Worker(intI).WorkSpeed
            .WorkNow = Worker(intI).WorkNow
            .CurrentSpeed = Worker(intI).CurrentSpeed

            .WalkSpeed = Worker(intI).WalkSpeed
            .IsMoving = Worker(intI).IsMoving
            .WayMove = Worker(intI).WayMove
            .SearchMove = Worker(intI).SearchMove
            .MoveDuration = Worker(intI).MoveDuration
            .XHeadSmooth = Worker(intI).XHeadSmooth
            .YHeadSmooth = Worker(intI).YHeadSmooth

            .XPoint = Worker(intI).XPoint
            .YPoint = Worker(intI).YPoint
        End With
    Next intI
End Sub

Public Sub WriteAlignRight(PosisiX As Integer, PosisiY As Integer, Number As Variant, ddScreen As DirectDrawSurface7)
    Dim bytVarX As Integer
    Dim Tulisan As String
    Tulisan = Format$(Number, "#,#0.00")
    
    For bytVarX = 0 To Len(Tulisan) - 1
    With ddScreen
        .DrawText PosisiX - (7 * bytVarX), PosisiY, Mid(Tulisan, IIf(Len(Tulisan) - bytVarX > 0, Len(Tulisan) - bytVarX, Len(Tulisan)), 1), False
    End With
    Next bytVarX
End Sub

Public Sub WriteMoney()
    'rutin untuk mengupdate layar dengan jumlah uang tertera
    If NoArray <= 0 Then Exit Sub
    For intVar = 1 To NoArray
        With MoneyScreen(intVar)
            If VisibleMoney(.bytX, .bytY) And .bytLong >= 0 Then
                With stFont
                    .Bold = False
                    .Name = "Comic Sans MS"
                    .Size = 8
                End With
                BackBuffer.SetFont stFont
                BackBuffer.SetForeColor RGB(0, 0, 0)
                
                Select Case .Arah
                Case MOVE_LEFT
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll - (.bytLong - 10) + 2, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll + 2, .JumlahUang, False
                Case MOVE_RIGHT
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll + .bytLong + 2, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll + 2, .JumlahUang, False
                Case MOVE_UP
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll + 2, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll - (.bytLong - 10) + 2, .JumlahUang, False
                Case MOVE_DOWN
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll + 2, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll + .bytLong + 2, .JumlahUang, False
                End Select
            
                If CInt(.JumlahUang) < 0 Then
                    BackBuffer.SetForeColor RGB(255, 150, 150)
                Else
                    BackBuffer.SetForeColor RGB(0, 255, 0)
                End If
                Select Case .Arah
                Case MOVE_LEFT
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll - (.bytLong - 10), ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll, .JumlahUang, False
                Case MOVE_RIGHT
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll + .bytLong, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll, .JumlahUang, False
                Case MOVE_UP
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll - (.bytLong - 10), .JumlahUang, False
                Case MOVE_DOWN
                    BackBuffer.DrawText ((.bytX - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll, ((.bytY - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll + .bytLong, .JumlahUang, False
                End Select
                .bytLong = .bytLong - 1
            End If
        End With
    Next intVar
        
    'cek array pertama bernilai nol atau tidak (FIFO)
    If NoArray > 0 Then
    If MoneyScreen(1).bytLong <= 0 Then
        'copy array menjadi ubound-1
        For intVar = 1 To NoArray - 1
            MoneyScreen(intVar).bytX = MoneyScreen(intVar + 1).bytX
            MoneyScreen(intVar).bytY = MoneyScreen(intVar + 1).bytY
            MoneyScreen(intVar).bytLong = MoneyScreen(intVar + 1).bytLong
            MoneyScreen(intVar).JumlahUang = MoneyScreen(intVar + 1).JumlahUang
            MoneyScreen(intVar).Arah = MoneyScreen(intVar + 1).Arah
        Next intVar
        'redim sekali lagi
        NoArray = NoArray - 1
        ReDim Preserve MoneyScreen(NoArray)
    End If
    End If
End Sub

Private Sub WriteWorkerID(SelNo As Integer)
    'rutin untuk menulis identitas worker
    
    'tampilkan captions
    With stFont
        .Bold = False
        .Size = 7
        .Name = "Arial"
    End With
    'set to backbuffer
    BackBuffer.SetFont stFont
    BackBuffer.SetForeColor RGB(255, 255, 255)
    BackBuffer.DrawText 715, 525, "No. : " & Worker(SelNo).NoLemm, False
    BackBuffer.DrawText 715, 535, "Nama : " & Worker(SelNo).Nama, False
    BackBuffer.DrawText 715, 545, "Job :" & Worker(SelNo).Job, False
    BackBuffer.DrawText 715, 555, "Honor : " & Worker(SelNo).Honor, False
    BackBuffer.DrawText 715, 565, "Stabilitas : " & Worker(SelNo).Stability, False
    BackBuffer.DrawText 715, 575, "Toleransi : " & Worker(SelNo).Tolerance, False
    BackBuffer.DrawBox 715, 586, 790, 593
    'update progress bar
    BackBuffer.BltColorFill BoxRect(716, 587, 716 + Int((Worker(SelNo).Progress / 100) * 73), 592), RGB(100, 250, 0)
    
    BackBuffer.BltFast 660, 540, IconWorker, BoxRect(0, 0, 39, 44), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    'gambar sprite panah di atas worker
    Worker(SelNo).UpdatePosition
    If VisiblePointer(SelNo) And Not PAUSE_GAME Then BackBuffer.BltFast Worker(SelNo).XPosition + Worker(SelNo).XHeadSmooth, Worker(SelNo).YPosition - 17, TexMouse, BoxRect(15, 0, 30, 15), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End Sub


Public Function ConvToSignedValue(lngValue As Long) As Integer

    'Cheezy method for converting to signed integer
    If lngValue <= 32767 Then
        ConvToSignedValue = CInt(lngValue)
        Exit Function
    End If
    
    ConvToSignedValue = CInt(lngValue - 65535)

End Function

Public Function ConvToUnSignedValue(intValue As Integer) As Long

    'Cheezy method for converting to unsigned integer
    If intValue >= 0 Then
        ConvToUnSignedValue = intValue
        Exit Function
    End If
    
    ConvToUnSignedValue = intValue + 65535

End Function

Private Sub WriteCraneID()
    'tulis identitas crane di daerah selected
    'tampilkan captions
    With stFont
        .Bold = False
        .Size = 7
        .Name = "Arial"
    End With
    'set to backbuffer
    BackBuffer.SetFont stFont
    BackBuffer.SetForeColor RGB(255, 255, 255)
    BackBuffer.DrawText 715, 525, "CRANE", False
    BackBuffer.DrawText 715, 535, "Occupied : " & Crane.Occupied, False
    BackBuffer.DrawText 715, 545, "Height : " & Crane.HeightCrane * 1.6 & " m", False
    BackBuffer.DrawText 715, 555, "Length : " & Crane.LengthCrane * 1.25 & "m", False
    BackBuffer.BltFast 660, 527, IconWorker, BoxRect(40, 0, 77, 44), DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    
    'gambar sprite panah di atas crane
    If VisibleCrane And Not PAUSE_GAME Then BackBuffer.BltFast Crane.XCock, Crane.YCock - 18, TexMouse, BoxRect(15, 0, 30, 15), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
End Sub

Private Function VisibleCrane() As Boolean
    If Crane.Visible Then
        VisibleCrane = True
    Else
        VisibleCrane = False
    End If
End Function
