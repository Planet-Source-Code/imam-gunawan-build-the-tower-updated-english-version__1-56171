Attribute VB_Name = "GameMod"
Option Explicit

'Module ini mengatur segala perangkat mengenai permainan
'atau yang mengendalikan flow game

Type udtStatusGame
    lFixBudget As Long
    Budget As Long
    Pinjaman As Long
    bytBunga As Single
        
    lOngkos As Long         'dana ongkos dan gaji
    lPendirian As Long      'uang pendirian bangunan
    
    Cuaca As Byte
    dHariIni As Date
    dWaktu As Date
End Type

Public StatusGame As udtStatusGame
Public stFont As New StdFont
Public SkalaMap As Single
Public WidthMap As Single
Public HeightMap As Single

Type udtGraphGame
    StrBln As String
    intSkorGame As Integer
    lMoney As Long
    bSudah As Boolean
End Type

Type udtObjektif
    dHariSelesai As Date
    dSketsa As Byte
    bytStart As Byte
    bytEnd As Byte
End Type

Type udtTmpbytSpr
    bytSprX As Byte
    bytSprY As Byte
End Type

Public TipsGame(10) As String
Public NoTips As Byte

Public Graph(12) As udtGraphGame
Public Objektif As udtObjektif

Private TooltipStr As String

Public NoArray As Byte
Public TiangArray As Byte
Public Tampakbangunan As Byte

'Variabel window on/off
Public WindowMaker As Boolean
Public RedrawWindowMaker As Boolean
Public WindowBelanjaBata As Boolean
Public RedrawWindowBata As Boolean
Public WindowBelanjaPasir As Boolean
Public RedrawWindowPasir As Boolean
Public WindowBelanjaSemen As Boolean
Public RedrawWindowSemen As Boolean
Public WindowBelanjaCat As Boolean
Public RedrawWindowCat As Boolean
Public WindowBelanjaKayu As Boolean
Public RedrawWindowKayu As Boolean
Public WindowBelanjaBesi As Boolean
Public RedrawWindowBesi As Boolean
Public WindowBelanjaKaca As Boolean
Public RedrawWindowKaca As Boolean
Public WindowBudget As Boolean
Public RedrawWindowBudget As Boolean
Public WindowGraph As Boolean
Public RedrawWindowGraph As Boolean
Public WindowDanaPekerja As Boolean
Public RedrawWindowDana As Boolean
Public WindowSketch As Boolean
Public RedrawWindowSketch As Boolean
Public WindowTiang As Boolean
Public RedrawWindowTiang As Boolean
Public WindowHelp As Boolean
Public RedrawWindowHelp As Boolean
Public WindowTips As Boolean
Public RedrawWindowTips As Boolean
Public WindowWorkShop As Boolean
Public RedrawWindowWorkshop As Boolean
Public WindowSplash As Boolean
Public RedrawWindowSplash As Boolean

Public RedrawMiniMap As Boolean

Private Type udtError
    ErrorWindow As Boolean
    RedrawError As Boolean
    StrKata As String
End Type

Public ShowError As udtError

'Variabel penentu pekerjaan worker
Public PASANGBATU As Boolean
Public PasangTiang As Boolean
Public PasangTangga As Boolean
Public Penghancur As Boolean
Public PasangJendela As Boolean
Public PasangCat As Boolean

Private JarakX As Single
Private JarakY As Single
Private X As Byte
Private Y As Byte
Private bHighlight As Byte
Private intLimitX As Integer
Private intX As Integer
Private intvarx As Integer
Private intvary As Integer

'status HUD on screen game
Public PAUSE_GAME As Boolean
Public MousePointer As Byte
Public ShowMap As Boolean
Public MusikOn As Boolean

Public Const MAXGEDUNGX = 150
Public Const MAXGEDUNGY = 150

Public Const MAXTILEY = 28
Public Const MAXTILEX = 57

Public Const MOUSESPEED = 2
Public Const SCROLLSPEED_LEFT_RIGHT = 7     'harus dibagi habis terhadap 14
Public Const SCROLLSPEED_UP_DOWN = 5        'harus dibagi habis terhadap 20
Public Const TILEWIDTH = 14
Public Const TILEHEIGHT = 20

Public Const SCROLL_LEFT = 0
Public Const SCROLL_RIGHT = 1
Public Const SCROLL_UP = 2
Public Const SCROLL_DOWN = 3
Public Const SCROLL_NONE = 4

Public Const MOVE_NONE = 0
Public Const MOVE_LEFT = 1
Public Const MOVE_RIGHT = 2
Public Const MOVE_UP = 3
Public Const MOVE_DOWN = 4

Public Const CRANE_MOVE = 1
Public Const CRANE_RISE = 2
Public Const CRANE_LOWER = 3

Public Const MOUSE_DEFAULT = 0
Public Const MOUSE_POINT = 2
Public Const MOUSE_SCROLL_DOWN = 4
Public Const MOUSE_SCROLL_LEFT = 5
Public Const MOUSE_SCROLL_UP = 6
Public Const MOUSE_SCROLL_RIGHT = 7

Public Const TAMPAK_LUAR = 0
Public Const TAMPAK_DALAM = 1

Public ScrollScreen As Boolean

Public ScrollWay As Byte

'Something to hold our position of mouse
Public CursorX As Long
Public CursorY As Long
Public Mouse_Button0 As Boolean
Public Mouse_Button1 As Boolean
Public Mouse_Button2 As Boolean
Public Mouse_Button3 As Boolean

'Sprite calculating in Y
Public Const WALK_LEFT = 0
Public Const WALK_RIGHT = 1
Public Const PANIC = 2
Public Const DIG_LEFT = 3
Public Const DIG_RIGHT = 4
Public Const PUSH_LEFT = 5
Public Const PUSH_RIGHT = 6
Public Const STAND_LEFT = 7
Public Const STAND_RIGHT = 8
Public Const WELD_LEFT = 9
Public Const WELD_RIGHT = 10
Public Const UP_DOWN = 11
Public Const BUILDER = 12

'Job for workers
Public Const STAND = 0
Public Const DIGGER = 2
Public Const PASANG_TANGGA = 3
Public Const TUKANG_LAS = 4
Public Const PASANG_BATU = 5
Public Const PASANG_JENDELA = 6
Public Const HANCUR = 7
Public Const PASANG_CAT = 8

'Status for workers
Public Const IDLE = 0           'worker dalam keadaan diam
Public Const Walk = 1           'worker dalam keadaan berjalan
Public Const WORK = 2           'worker dalam keadaan bekerja
Public Const REST = 3           'worker dalam keadaan istirahat
Public Const CLIMB = 4          'worker dalam keadaan memanjat
Public Const WELD = 5           'worker dalam keadaan las

'Jam Mulai bekerja dan Selesai
Public Const JAM_MULAI = 8
Public Const MENIT_MULAI = 0
Public Const JAM_SELESAI = 17
Public Const MENIT_SELESAI = 0

'Temporary Array Gedung untuk perubahan View
Public TmpSpr(1 To MAXGEDUNGX, 1 To MAXGEDUNGY) As udtTmpbytSpr
Sub ApplyObjective()
    'rutin ini untuk menampilkan objektif beserta
    'gambar gedung yang diinginkan
    Objektif.dHariSelesai = #8/15/2000#
    Objektif.dSketsa = 1
    Select Case Objektif.dSketsa
    Case 1
        Objektif.bytStart = 50
        Objektif.bytEnd = 120
    End Select
End Sub


Function BangunJendela() As Boolean
    BangunJendela = False
    With Workshop
        If .bytTmpKayu >= 1 And .bytTmpKaca >= 1 Then
            BangunJendela = True
            .bytTmpKayu = .bytTmpKayu - 1
            .bytTmpKaca = .bytTmpKaca - 1
        End If
    End With
End Function


Function BangunPintu() As Boolean
    'komposisi 2 kayu
    BangunPintu = False
    With Workshop
        If .bytTmpKayu >= 2 Then
            BangunPintu = True
            .bytTmpKayu = .bytTmpKayu - 2
        End If
    End With
End Function


Function BangunTangga() As Boolean
    'komposisi 1 bata+2 pasir +1 semen
    BangunTangga = False
    With Workshop
        If .bytTmpBesi >= 2 Then
            BangunTangga = True
            .bytTmpBesi = .bytTmpBesi - 2
        End If
    End With
End Function

Function BangunTembok() As Boolean
    'rutin untuk mengecek apakah komposisi tembok cukup
    'dengan komposisi 1 bata+2 pasir +1 semen
    BangunTembok = False
    With Workshop
        If .bytBata >= 1 And .bytPasir >= 2 And .bytSemen >= 1 Then
            BangunTembok = True
            'potong jumlah tembok
            .bytBata = .bytBata - 1
            .bytPasir = .bytPasir - 2
            .bytSemen = .bytSemen - 1
        End If
    End With
End Function

Sub InitHargaBahan()
    'rutin ini melakukan init terhadap harga awal setiap bahan
    'yang digunakan untuk membangun bangunan
    With Workshop
        .intBata = 125
        .intBesi = 225
        .intCat = 75
        .intKayu = 80
        .intKaca = 80
        .intPasir = 50
        .intSemen = 200
    End With
End Sub
Sub InitTips()
    'rutin untuk menginisialisasi tips
    TipsGame(0) = "Click on Minimap to faster navigation"
    TipsGame(1) = "Use window Tips to assist you"
    TipsGame(2) = "Use the keypad button to navigate screen"
    TipsGame(3) = "Press Space to Pause the Game"
    TipsGame(4) = "Bla Bla Bla"
    TipsGame(5) = "Efficience your worker to maximize job"
    TipsGame(6) = "Higher honor to make your worker happy"
    TipsGame(7) = "Press M to toggle Minimap"
    TipsGame(8) = "La La La La La La...."
    TipsGame(9) = "Rome doesn't built in one day"
    TipsGame(10) = "Ehm Ehm Ehm..."
End Sub

Sub LetakAwan(ArX As Byte, ArY As Byte)
    'rutin ini untuk mengambil sprite awan dan meletakkan ke layar
    'dengan parameter ArX,ArY sebagai titik koordinat awal
    ArGedung(ArX, ArY).bytSprX = 2
    ArGedung(ArX, ArY).bytSprY = 4
    
    ArGedung(ArX + 1, ArY).bytSprX = 3
    ArGedung(ArX + 1, ArY).bytSprY = 4
    
    ArGedung(ArX + 2, ArY).bytSprX = 4
    ArGedung(ArX + 2, ArY).bytSprY = 4
End Sub

Function LolosDinding(TileX As Byte, TileY As Byte) As Boolean
    LolosDinding = True
    If ((TileX - 1) > 1 And (TileX + 1) < MAXGEDUNGX) And ((TileY + 1 < MAXGEDUNGY) And (TileY - 1 > 1)) Then
    If Not (ArGedung(TileX, TileY - 1).bWall Or ArGedung(TileX + 1, TileY).bWall Or _
    ArGedung(TileX - 1, TileY).bWall Or ArGedung(TileX + 1, TileY).bCor Or _
    ArGedung(TileX - 1, TileY).bCor) Then
        LolosDinding = False
    End If
    End If
End Function

Sub RefreshMiniMap()
    'rutin ini hanya dijalankan apabila Minimap akan direfresh
    'dengan data baru
    
    If RedrawMiniMap Then
    
    MiniMapDX.BltColorFill BoxRect(0, 0, 105, 105), RGB(0, 0, 0)
    'melakukan penggambaran terhadap minimap dengan pengisian setiap pixel dengan nilai array
    
    Call GameMod.DrawOrangeBox(0, 0, 14, 14, MiniMapDX)
    For X = 1 To 105
        For Y = 1 To 105
            'cek terhadap crane
            If ArGedung(Int(X * SkalaMap), Int(Y * SkalaMap)).bCrane = True Then
                MiniMapDX.BltFast X, Y, Sprite, BoxRect(58, 0, 59, 1), DDBLTFAST_WAIT
                
            'cek terhadap wall
            ElseIf ArGedung(Int(X * SkalaMap), Int(Y * SkalaMap)).bWall = True Then
                MiniMapDX.BltFast X, Y, Sprite, BoxRect(6, 0, 7, 1), DDBLTFAST_WAIT
                
            'cek terhadap tiang
            ElseIf ArGedung(Int(X * SkalaMap), Int(Y * SkalaMap)).bTiang = True Then
                MiniMapDX.BltFast X, Y, Sprite, BoxRect(28, 40, 29, 41), DDBLTFAST_WAIT
            
            'cek terhadap Cor
            ElseIf ArGedung(Int(X * SkalaMap), Int(Y * SkalaMap)).bCor = True Then
                MiniMapDX.BltFast X, Y, Sprite, BoxRect(28, 40, 29, 41), DDBLTFAST_WAIT
            
            End If
        Next Y
    Next X
    
    RedrawMiniMap = False
    End If
    
    'transfer ke backbuffer
    BackBuffer.BltFast 693, 0, MiniMapDX, BoxRect(0, 0, 105, 105), DDBLTFAST_WAIT
            
End Sub
Sub ShowBelanjaBata()
    'rutin ini untuk menampilkan tampilan window belanja Bata
    'untuk bayangan
    
    If RedrawWindowBata Then
    
    RedrawWindowBata = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Buy Brick", False
    TempDX.DrawText 210, 230, "Qty Brick : ", False
    TempDX.DrawText 210, 270, "Rem. Brick : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Cancel", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytBata & " pcs", False
    TempDX.DrawText 390, 230, .bytTmpBata & " pcs", False
    TempDX.DrawText 430, 230, " x " & Format(.intBata, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpBata * .intBata, "#,#0"))) & Format(.bytTmpBata * .intBata, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            .bytTmpBata = .bytTmpBata + 1
            RedrawWindowBata = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            If .bytTmpBata > 0 Then .bytTmpBata = .bytTmpBata - 1
            RedrawWindowBata = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        TotalWorkshop.bytTmpBata = TotalWorkshop.bytTmpBata + .bytTmpBata
        WindowBelanjaBata = False
        PAUSE_GAME = False
        .bytBata = .bytBata + .bytTmpBata
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpBata * .intBata) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        'tombol Cancel ditekan
        WindowBelanjaBata = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With
End Sub


Sub ShowWindowWorkshop()
    'rutin ini untuk menampilkan tampilan isi workshop
    
    If RedrawWindowWorkshop Then
    
    RedrawWindowWorkshop = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 455), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 451
    TempDX.BltColorFill BoxRect(201, 201, 599, 450), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Tahoma"
        .Size = 8
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "W O R K S H O P", False
    TempDX.DrawText 210, 230, "Qty Brick : ", False
    TempDX.DrawText 210, 245, "Qty Sand : ", False
    TempDX.DrawText 210, 260, "Qty Cement : ", False
    TempDX.DrawText 210, 275, "Qty Lumber : ", False
    TempDX.DrawText 210, 290, "Qty Iron : ", False
    TempDX.DrawText 210, 305, "Qty Paint : ", False
    TempDX.DrawText 210, 320, "Qty Glass : ", False
    TempDX.DrawText 400, 230, "Qty Door : ", False
    TempDX.DrawText 400, 245, "Qty Ladder : ", False
    TempDX.DrawText 400, 260, "Qty Window : ", False
    TempDX.DrawText 210, 360, "To see the composition mixture , refer to help menu ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 420, 500, 440, 4, 4
    TempDX.SetFont stFont
    TempDX.DrawText 460, 422, "OK", False
    With Workshop
        'tampilkan result workshop
        TempDX.DrawText 330, 230, .bytBata, False
        TempDX.DrawText 330, 245, .bytPasir, False
        TempDX.DrawText 330, 260, .bytSemen, False
        TempDX.DrawText 330, 275, .bytKayu, False
        TempDX.DrawText 330, 290, .bytBesi, False
        TempDX.DrawText 330, 305, .bytCat, False
        TempDX.DrawText 330, 320, .bytKaca, False
        TempDX.DrawText 520, 230, .bytPintu, False
        TempDX.DrawText 520, 245, .bytTangga, False
        TempDX.DrawText 520, 260, .bytJendela, False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 455), DDBLTFAST_WAIT
    
    If Mouse_Button0 Then
        Select Case CursorX
        Case 440 To 500     'tombol OK
        Select Case CursorY
        Case 420 To 440
            WindowWorkShop = False
            PAUSE_GAME = False
        End Select
        End Select
    End If
End Sub

Sub ShowBelanjaBesi()
    'rutin ini untuk menampilkan tampilan window belanja Bata
    'untuk bayangan
    
    If RedrawWindowBesi Then
    
    RedrawWindowBesi = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Buy Iron", False
    TempDX.DrawText 210, 230, "Qty Iron : ", False
    TempDX.DrawText 210, 270, "Rem. Iron : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytBesi & " btg", False
    TempDX.DrawText 390, 230, .bytTmpBesi & " btg", False
    TempDX.DrawText 430, 230, " x " & Format(.intBesi, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpBesi * .intBesi, "#,#0"))) & Format(.bytTmpBesi * .intBesi, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            .bytTmpBesi = .bytTmpBesi + 1
            RedrawWindowBesi = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            If .bytTmpBesi > 0 Then .bytTmpBesi = .bytTmpBesi - 1
            RedrawWindowBesi = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        'tombol OK ditekan
        TotalWorkshop.bytTmpBesi = TotalWorkshop.bytTmpBesi + .bytTmpBesi
        'tombol OK ditekan
        WindowBelanjaBesi = False
        PAUSE_GAME = False
        .bytBesi = .bytBesi + .bytTmpBesi
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpBesi * .intBesi) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        'tombol Cancel ditekan
        WindowBelanjaBesi = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With

End Sub

Sub ShowBelanjaCat()
    'rutin ini untuk menampilkan tampilan window belanja Bata
    'untuk bayangan
    
    If RedrawWindowCat Then
    
    RedrawWindowCat = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Buy Paint", False
    TempDX.DrawText 210, 230, "Qty Paint : ", False
    TempDX.DrawText 210, 270, "Rem Paint: ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytCat & " can", False
    TempDX.DrawText 390, 230, .bytTmpCat & " can", False
    TempDX.DrawText 430, 230, " x " & Format(.intCat, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpCat * .intCat, "#,#0"))) & Format(.bytTmpCat * .intCat, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            .bytTmpCat = .bytTmpCat + 1
            RedrawWindowCat = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            If .bytTmpCat > 0 Then .bytTmpCat = .bytTmpCat - 1
            RedrawWindowCat = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        'tombol OK ditekan
        TotalWorkshop.bytTmpCat = TotalWorkshop.bytTmpCat + .bytTmpCat
        WindowBelanjaCat = False
        PAUSE_GAME = False
        .bytCat = .bytCat + .bytTmpCat
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpCat * .intCat) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        'tombol Cancel ditekan
        WindowBelanjaCat = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With

End Sub

Sub ShowBelanjaKaca()
    'rutin ini untuk menampilkan tampilan window belanja Kaca
    
    If RedrawWindowKaca Then
    
    RedrawWindowKaca = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Buy Glass", False
    TempDX.DrawText 210, 230, "Qty Glass : ", False
    TempDX.DrawText 210, 270, "Rem. Glass : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytKaca & " bh", False
    TempDX.DrawText 390, 230, .bytTmpKaca & " bh", False
    TempDX.DrawText 430, 230, " x " & Format(.intKaca, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpKaca * .intKaca, "#,#0"))) & Format(.bytTmpKaca * .intKaca, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin kaca
            .bytTmpKaca = .bytTmpKaca + 1
            RedrawWindowKaca = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin kaca
            If .bytTmpKaca > 0 Then .bytTmpKaca = .bytTmpKaca - 1
            RedrawWindowKaca = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        TotalWorkshop.bytTmpKaca = TotalWorkshop.bytTmpKaca + .bytTmpKaca
        WindowBelanjaKaca = False
        PAUSE_GAME = False
        .bytKaca = .bytKaca + .bytTmpKaca
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpKaca * .intKaca) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        'tombol Cancel ditekan
        WindowBelanjaKaca = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With
End Sub

Sub ShowBelanjaKayu()
    'rutin ini untuk menampilkan tampilan window belanja Bata
    'untuk bayangan
    
    If RedrawWindowKayu Then
    
    RedrawWindowKayu = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Belanja Kayu", False
    TempDX.DrawText 210, 230, "Jumlah Kayu : ", False
    TempDX.DrawText 210, 270, "Sisa Kayu : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytKayu & " btg", False
    TempDX.DrawText 390, 230, .bytTmpKayu & " btg", False
    TempDX.DrawText 430, 230, " x " & Format(.intKayu, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpKayu * .intKayu, "#,#0"))) & Format(.bytTmpKayu * .intKayu, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            .bytTmpKayu = .bytTmpKayu + 1
            RedrawWindowKayu = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin cat
            If .bytTmpKayu > 0 Then .bytTmpKayu = .bytTmpKayu - 1
            RedrawWindowKayu = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        TotalWorkshop.bytTmpKayu = TotalWorkshop.bytTmpKayu + .bytTmpKayu
        WindowBelanjaKayu = False
        PAUSE_GAME = False
        .bytKayu = .bytKayu + .bytTmpKayu
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpKayu * .intKayu) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        'tombol Cancel ditekan
        WindowBelanjaKayu = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With
End Sub
Sub ShowBelanjaPasir()
    'rutin ini untuk menampilkan tampilan window belanja Bata
    'untuk bayangan
    
    If RedrawWindowPasir Then
    
    RedrawWindowPasir = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Belanja Pasir", False
    TempDX.DrawText 210, 230, "Jumlah Pasir : ", False
    TempDX.DrawText 210, 270, "Sisa Pasir : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytPasir & " m^3", False
    TempDX.DrawText 390, 230, .bytTmpPasir & " m^3", False
    TempDX.DrawText 430, 230, " x " & Format(.intPasir, "#,#0"), False
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpPasir * .intPasir, "#,#0"))) & Format(.bytTmpPasir * .intPasir, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            .bytTmpPasir = .bytTmpPasir + 1
            RedrawWindowPasir = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            If .bytTmpPasir > 0 Then .bytTmpPasir = .bytTmpPasir - 1
            RedrawWindowPasir = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        TotalWorkshop.bytTmpPasir = TotalWorkshop.bytTmpPasir + .bytTmpPasir
        WindowBelanjaPasir = False
        PAUSE_GAME = False
        .bytPasir = .bytPasir + .bytTmpPasir
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpPasir * .intPasir) * -1, "#,#0"), MOVE_DOWN)
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        WindowBelanjaPasir = False
        PAUSE_GAME = False
    End Select
    End Select
    
    End If
    End With
End Sub
Sub ShowBelanjaSemen()
    'rutin ini untuk menampilkan tampilan window belanja Semen
    'untuk bayangan
    
    If RedrawWindowSemen Then
    
    RedrawWindowSemen = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 355), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 351
    TempDX.BltColorFill BoxRect(201, 201, 599, 350), RGB(100, 250, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(150, 200, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 350, 202, "Belanja Semen", False
    TempDX.DrawText 210, 230, "Jumlah Semen : ", False
    TempDX.DrawText 210, 270, "Sisa Semen : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 300, 500, 320, 4, 4
    TempDX.DrawRoundedBox 510, 300, 570, 320, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 302, "OK", False
    TempDX.DrawText 524, 302, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 300, 270, .bytSemen & " sak", False
    TempDX.DrawText 390, 230, .bytTmpSemen & " sak", False
    TempDX.DrawText 430, 230, " x " & Format(.intSemen, "#,#0"), False
    'Call WriteAlignRight(460, 230, .bytSemen * .intSemen, TempDX)
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTmpSemen * .intSemen, "#,#0"))) & Format(.bytTmpSemen * .intSemen, "#,#0"), False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 355), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            .bytTmpSemen = .bytTmpSemen + 1
            RedrawWindowSemen = True
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            If .bytTmpSemen > 0 Then .bytTmpSemen = .bytTmpSemen - 1
            RedrawWindowSemen = True
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 300 To 320
        TotalWorkshop.bytTmpSemen = TotalWorkshop.bytTmpSemen + .bytTmpSemen
        WindowBelanjaSemen = False
        PAUSE_GAME = False
        .bytSemen = .bytSemen + .bytTmpSemen
        Call DrawMoneyOnScreen(148, 140, Format((.bytTmpSemen * .intSemen) * -1, "#,#0"), MOVE_DOWN)
    End Select
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 300 To 320
        WindowBelanjaSemen = False
        PAUSE_GAME = False
    End Select
    End Select
    End If
    End With
End Sub
Sub ShowSketch(NomorSketch As Byte)
    'rutin untuk menampilkan objektif
    If RedrawWindowSketch Then
        RedrawWindowSketch = False
        
        TempDX.SetForeColor RGB(255, 255, 255)
        TempDX.BltFast 0, 0, BackBuffer, BoxRect(0, 0, 800, 600), DDBLTFAST_WAIT
        
        TempDX.DrawBox 174, 29, 626, 481
        
        Select Case NomorSketch
        Case 1
            LoadSprite App.Path & "\sketsa\skets1.bmp", 0, SketsaDX
            TempDX.BltFast 175, 30, SketsaDX, BoxRect(0, 0, 450, 450), DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
        End Select
    End If
    
    BackBuffer.BltFast 174, 29, TempDX, BoxRect(174, 29, 626, 481), DDBLTFAST_WAIT
    
    If Mouse_Button0 Then
        Mouse_Button0 = False
        WindowSketch = False
        PAUSE_GAME = False
    End If
End Sub
Sub ShowWindowGraph()
    'rutin ini akan melakukan update tampilan terhadap grafik permainan
    If RedrawWindowGraph Then
    
    RedrawWindowGraph = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.BltFast 0, 0, BackBuffer, BoxRect(0, 0, 800, 600), DDBLTFAST_WAIT
    
    TempDX.BltColorFill BoxRect(105, 105, 705, 465), QBColor(0)
    TempDX.DrawBox 100, 100, 700, 460
    TempDX.BltColorFill BoxRect(101, 101, 699, 464), RGB(0, 110, 50)
    TempDX.BltColorFill BoxRect(101, 101, 699, 120), RGB(150, 200, 150)
    With stFont
        .Name = "Comic Sans MS"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    TempDX.SetForeColor RGB(0, 0, 0)
    TempDX.DrawText 352, 103, "Grafik Permainan", False
    TempDX.DrawText 562, 434, "OK", False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.DrawText 350, 101, "Grafik Permainan", False
    TempDX.DrawText 560, 432, "OK", False
    
    'gambar grafik line
    TempDX.DrawLine 150, 140, 150, 390
    TempDX.DrawLine 150, 390, 670, 390
    'mulai tulis nilai uang di garis vertikal
    Dim bytVert As Integer
    For bytVert = 8 To 1 Step -1
        TempDX.SetForeColor RGB(0, 0, 0)
        TempDX.DrawText 122, 382 - (bytVert * 30), bytVert * 100, False
        TempDX.SetForeColor RGB(255, 255, 255)
        TempDX.DrawText 120, 380 - (bytVert * 30), bytVert * 100, False
    Next bytVert
    'garis horizontal
    For bytVert = 1 To 12
        TempDX.SetForeColor RGB(0, 0, 0)
        TempDX.DrawText 152 + (bytVert * 40), 395, Format(DateSerial(1, Val(Format("01 " & Graph(1).StrBln, "mm")) + (bytVert - 1), 1), "mmm"), False
        TempDX.SetForeColor RGB(255, 255, 255)
        TempDX.DrawText 150 + (bytVert * 40), 393, Format(DateSerial(1, Val(Format("01 " & Graph(1).StrBln, "mm")) + (bytVert - 1), 1), "mmm"), False
    Next bytVert
    
    'gambar OK dan Cancel
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.DrawRoundedBox 540, 430, 600, 450, 4, 4
    
    TempDX.SetForeColor RGB(0, 0, 0)
    'mulai gambarkan garis grafik di dalam kotak
    Dim bytSkalaPrev As Single
    Dim bytSkalaNext As Single
    'gambarkan garis skor ke grafiknya
    For bytVert = 1 To 12
        If Graph(bytVert).bSudah = False Then Exit For
        'gambarkan garis skor
        bytSkalaPrev = (Graph(bytVert - 1).intSkorGame * 8) * (250 / 800)
        bytSkalaNext = (Graph(bytVert).intSkorGame * 8) * (250 / 800)
        TempDX.SetForeColor RGB(150, 0, 120)
        TempDX.DrawLine ((bytVert - 1) * 40) + 150, 390 - Int(bytSkalaPrev), ((bytVert) * 40) + 150, 390 - Int(bytSkalaNext)
    Next bytVert
    End If
    BackBuffer.BltFast 100, 100, TempDX, BoxRect(100, 100, 705, 465), DDBLTFAST_WAIT
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    
    Select Case CursorX
    Case 540 To 600     'tombol OK
    Select Case CursorY
    Case 430 To 450
        WindowGraph = False
        PAUSE_GAME = False
    End Select
    
    End Select
    End If
End Sub
Sub ShowWindowMaker()
    'rutin ini untuk menampilkan tampilan window
    'untuk bayangan
    Dim lTotal As Long
    If RedrawWindowMaker Then
    
    RedrawWindowMaker = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 405), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 400
    TempDX.BltColorFill BoxRect(201, 201, 599, 399), RGB(0, 150, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(0, 100, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.SetFontTransparency True
    TempDX.DrawText 350, 202, "Pembuatan", False
    TempDX.DrawText 210, 230, "Buat Tangga  : ", False
    TempDX.DrawText 210, 250, "Buat Jendela : ", False
    TempDX.DrawText 210, 270, "Buat Pintu   : ", False
    TempDX.DrawText 320, 230, " +   -   : ", False
    TempDX.DrawText 320, 250, " +   -   : ", False
    TempDX.DrawText 320, 270, " +   -   : ", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 370, 500, 390, 4, 4
    TempDX.DrawRoundedBox 510, 370, 570, 390, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 372, "OK", False
    TempDX.DrawText 522, 372, "Batal", False
    
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    
    With Workshop
    'tampilkan result pemilihan objek
    TempDX.DrawText 390, 230, .bytTmpTangga & " bh", False
    TempDX.DrawText 390, 250, .bytTmpJendela & " bh", False
    TempDX.DrawText 390, 270, .bytTmpPintu & " bh", False
    
    TempDX.DrawText 430, 230, " x " & Format(.intTangga, "#,#0"), False
    TempDX.DrawText 430, 250, " x " & Format(.intJendela, "#,#0"), False
    TempDX.DrawText 430, 270, " x " & Format(.intPintu, "#,#0"), False
    
    TempDX.DrawText 460, 230, "  = " & Space(10 - Len(Format(.bytTangga * .intTangga, "#,#0"))) & Format(.bytTmpTangga * .intTangga, "#,#0"), False
    TempDX.DrawText 460, 250, "  = " & Space(10 - Len(Format(.bytJendela * .intJendela, "#,#0"))) & Format(.bytTmpJendela * .intJendela, "#,#0"), False
    TempDX.DrawText 460, 270, "  = " & Space(10 - Len(Format(.bytPintu * .intPintu, "#,#0"))) & Format(.bytTmpPintu * .intPintu, "#,#0"), False
    TempDX.SetForeColor RGB(0, 0, 0)
    TempDX.DrawLine 210, 290, 594, 290
    TempDX.SetForeColor RGB(255, 255, 255)
    
    lTotal = (.bytTmpJendela * .intJendela) + (.bytTmpPintu * .intPintu) + (.bytTmpTangga * .intTangga)
    TempDX.DrawText 420, 300, "Total   =" & Space(11 - Len(Format(lTotal, "#,#0"))) & Format(lTotal, "#,#0"), False
    
    'tampilkan jumlah bahan baku
    TempDX.DrawText 210, 310, "Jumlah Kayu =  " & .bytTmpKayu, False
    TempDX.DrawText 210, 325, "Jumlah Kaca =  " & .bytTmpKaca, False
    TempDX.DrawText 210, 340, "Jumlah Besi =  " & .bytTmpBesi, False
    End With
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 405), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    With Workshop
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 329 To 333         'pada posisi plus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            If BangunTangga Then
                .bytTmpTangga = .bytTmpTangga + 1
                RedrawWindowMaker = True
            End If
        Case 250 To 262     'bikin jendela
            If BangunJendela Then
                .bytTmpJendela = .bytTmpJendela + 1
                RedrawWindowMaker = True
            End If
        Case 276 To 281     'bikin pintu
            If BangunPintu Then
                .bytTmpPintu = .bytTmpPintu + 1
                RedrawWindowMaker = True
            End If
        End Select
        'MousePointer = 2
    Case 355 To 361     'pada posisi minus
        Select Case CursorY
        Case 235 To 241     'bikin tangga
            If .bytTmpTangga > 0 Then
                .bytTmpTangga = .bytTmpTangga - 1
                RedrawWindowMaker = True
                'kembalikan ke bahan baku
                .bytTmpBesi = .bytTmpBesi + 2
            End If
        Case 250 To 262     'bikin jendela
            If .bytTmpJendela > 0 Then
                .bytTmpJendela = .bytTmpJendela - 1
                RedrawWindowMaker = True
                .bytTmpKaca = .bytTmpKaca + 1
                .bytTmpKayu = .bytTmpKayu + 1
            End If
        Case 276 To 281     'bikin pintu
            If .bytTmpPintu > 0 Then
                .bytTmpPintu = .bytTmpPintu - 1
                RedrawWindowMaker = True
                .bytTmpKayu = .bytTmpKayu + 2
            End If
        End Select
        'MousePointer = 2
        
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 370 To 390
        'cek apakah jumlah uang mencukupi
        lTotal = (.bytTmpJendela * .intJendela) + (.bytTmpPintu * .intPintu) + (.bytTmpTangga * .intTangga)
        .bytTangga = .bytTmpTangga
        .bytJendela = .bytTmpJendela
        .bytPintu = .bytTmpPintu
        
        .bytBesi = .bytTmpBesi
        .bytKayu = .bytTmpKayu
        .bytKaca = .bytTmpKaca
        
        'tombol OK ditekan
        WindowMaker = False
        PAUSE_GAME = False
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 370 To 390
        'tombol Cancel ditekan
        WindowMaker = False
        PAUSE_GAME = False
    End Select
    End Select
    End If
    End With
End Sub

Sub ShowWindowTiang()
    'rutin ini untuk menampilkan tampilan window
    'untuk bayangan
    
    If RedrawWindowTiang Then
    
    RedrawWindowTiang = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 405), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 400
    TempDX.BltColorFill BoxRect(201, 201, 599, 399), RGB(0, 150, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(0, 100, 200)
    With stFont
        .Name = "Courier New"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.SetFontTransparency True
    TempDX.SetForeColor RGB(0, 0, 0)
    TempDX.DrawText 350, 202, "Tiang Horizontal", False
    TempDX.DrawText 210, 230, "Lebar Tiang  : " & Crane.LebarTiang, False
    TempDX.DrawText 210, 250, "Biaya : " & (Crane.LebarTiang * 100), False
    TempDX.DrawText 210, 285, "(Min. 4 dan Maks. 12)", False
    TempDX.DrawText 210, 300, "Gunakan tombol (PgUp) dan (PgDn) untuk lebar tiang", False
    
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 440, 370, 500, 390, 4, 4
    TempDX.DrawRoundedBox 510, 370, 570, 390, 4, 4
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.DrawText 460, 372, "OK", False
    TempDX.DrawText 522, 372, "Batal", False
    
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 405), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 440 To 500     'tombol OK
    Select Case CursorY
    Case 370 To 390
        'tombol OK ditekan
        'periksa apakah dana mencukupi
        With Crane
            .Occupied = True       'head telah terisi
            'kurangi panjang chain menjadi
            .LengthChain = .LengthChain - 2
            .YPoint = .YHead
            'masukkan data tiang
            .XTiangOnArray = .XHead - (.LebarTiang \ 2)
            .YTiangOnArray = .YHead + .LengthChain + 1
        End With

        PAUSE_GAME = False
        WindowTiang = False
        
        'lakukan tampilan lebar tiang
    End Select
    
    Case 510 To 570     'tombol Cancel
    Select Case CursorY
    Case 370 To 390
        'tombol Cancel ditekan
        With Crane
            .Occupied = False 'head tidak terisi
            'kurangi panjang chain menjadi
            .LengthChain = .LengthChain - 2
            .YPoint = .YHead
        End With
        
        PAUSE_GAME = False
        WindowTiang = False
    End Select
    End Select
    End If
End Sub


Sub ShowWindowHelp()
    'rutin ini untuk menampilkan tampilan window
    'untuk bayangan
    
    If RedrawWindowHelp Then
    
    RedrawWindowHelp = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(55, 55, 755, 455), QBColor(0)
    TempDX.DrawBox 50, 50, 750, 450
    TempDX.BltColorFill BoxRect(51, 51, 749, 449), RGB(0, 150, 100)
    TempDX.BltColorFill BoxRect(52, 52, 748, 70), RGB(0, 100, 200)
    With stFont
        .Name = "Tahoma"
        .Size = 8
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.SetForeColor RGB(0, 0, 0)
    
    'help text
    With TempDX
        .DrawText 60, 55, "Help", False
        .DrawText 60, 75, "Crane Function", False
        .DrawText 60, 90, "= (Equal) : Raise Crane", False
        .DrawText 60, 105, "- (Minus) : Lower Crane", False
        .DrawText 60, 120, "[ (Left Bracket) : Shorten Hand", False
        .DrawText 60, 135, "] (Right Bracket) : Lengthen Hand", False
        .DrawText 60, 150, "P (P Button) : Place vertical pole to spot", False
        
        .DrawText 60, 180, "Screen Navigation", False
        .DrawText 60, 195, "Up,Down,Left,Right: Screen Scrolling", False
        .DrawText 60, 210, "M : Toogle On/Off Minimap", False
        .DrawText 60, 225, "1 : Inside View", False
        .DrawText 60, 240, "2 : Outside View", False
        
        .DrawText 60, 270, "Keyboard", False
        .DrawText 60, 285, "O : Select Crane", False
        '.DrawText 60, 300, "S : Melihat Splash Screen", False
        .DrawText 60, 315, "W : Show Material Qty", False
        
        .DrawText 400, 75, "Material Composition", False
        .DrawText 400, 90, "Wall: 1 Brick + 2 Sand + 1 Cement", False
        .DrawText 400, 105, "Door: 2 Lumber", False
        .DrawText 400, 120, "Window: 1 Lumber + 1 Glass", False
        .DrawText 400, 135, "Ladder: 2 Iron", False
    End With
    
    
    TempDX.SetForeColor RGB(0, 0, 0)
    'gambar OK dan Cancel
    TempDX.DrawRoundedBox 350, 415, 420, 440, 4, 4
    
    TempDX.DrawText 375, 420, "OK", False
    
    End If
    
    BackBuffer.BltFast 50, 50, TempDX, BoxRect(50, 50, 755, 455), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 350 To 420     'tombol OK
    Select Case CursorY
    Case 416 To 440
        'tombol OK ditekan
        PAUSE_GAME = False
        WindowHelp = False
    End Select
    End Select
    End If
End Sub

Sub AturCuacadanWaktu()
    'pengontrol waktu
    'atur delayment
    Dim intJ As Integer
    Dim XPos As Byte
    If Cuaca.EachFPS >= (Cuaca.Delayment * MAX_FPS) Then
        Cuaca.EachFPS = 0
        If Cuaca.MenitSekarang + 1 >= 60 Then
            Cuaca.JamSekarang = Cuaca.JamSekarang + 1
            Cuaca.MenitSekarang = 0
        Else
            Cuaca.MenitSekarang = Cuaca.MenitSekarang + 1
        End If
    Else
        Cuaca.EachFPS = Cuaca.EachFPS + 1
    End If
    
    'pengontrol hari
    If (Cuaca.JamSekarang = Cuaca.JamSelesai) And (Cuaca.MenitSekarang = Cuaca.MenitSelesai) Then
        'tambahan satu hari ke kalendar, dan bersihkan semuanya
        Cuaca.JamSekarang = Cuaca.JamMulai
        Cuaca.MenitSekarang = Cuaca.MenitMulai
        
        Dim TmpHariIni As String
        Dim EsokHari As String
        TmpHariIni = Format(Cuaca.HariIni, "mmm")
        Cuaca.HariIni = CDate(Cuaca.HariIni) + 1
        EsokHari = Format(Cuaca.HariIni, "mmm")
        'cek apakah waktu objektif sudah tercapai
        With Objektif
            If .dHariSelesai = (CDate(Cuaca.HariIni) - 1) Then
                'permainan berakhir
                ShowError.ErrorWindow = True
                ShowError.StrKata = "The target day is expired.. you lose !"
                ShowError.RedrawError = True
                PAUSE_GAME = True
                'hitung skor
                For XPos = 1 To 12
                    If Graph(XPos).bSudah = False Then Exit For
                Next XPos
                If XPos = 12 Then
                    If Graph(12).bSudah = True Then
                        'mundurkan semua isi array ke belakang dan sisakan terakhir
                        For XPos = 1 To 12
                            Graph(XPos - 1).lMoney = Graph(XPos).lMoney
                            Graph(XPos - 1).StrBln = Graph(XPos).StrBln
                            Graph(XPos - 1).intSkorGame = Graph(XPos).intSkorGame
                        Next XPos
                        'pada bagian terakhir sisipkan hasil terakhir
                        Graph(12).bSudah = True
                        Graph(12).StrBln = TmpHariIni
                        Graph(12).intSkorGame = HitungSkor()
                    End If
                Else
                    'masukkan ke bagian terakhir
                    Graph(XPos).bSudah = True
                    Graph(XPos).intSkorGame = HitungSkor()
                    Graph(XPos).StrBln = TmpHariIni
                End If
            End If
        End With
        
        'lakukan pemeriksaan terhadap hariini dengan mengambil 3 huruf
        If TmpHariIni <> EsokHari Or (Objektif.dHariSelesai = (CDate(Cuaca.HariIni) - 1)) Then
            'bagian ini mengisi array graph dengan memeriksa bsudah
            'selalu mengisi ke bagian terakhir dari array
            For XPos = 1 To 12
                If Graph(XPos).bSudah = False Then Exit For
            Next XPos
            If XPos = 12 Then
                If Graph(12).bSudah = True Then
                    'mundurkan semua isi array ke belakang dan sisakan terakhir
                    For XPos = 1 To 12
                        Graph(XPos - 1).lMoney = Graph(XPos).lMoney
                        Graph(XPos - 1).StrBln = Graph(XPos).StrBln
                        Graph(XPos - 1).intSkorGame = Graph(XPos).intSkorGame
                    Next XPos
                    'pada bagian terakhir sisipkan hasil terakhir
                    Graph(12).bSudah = True
                    Graph(12).StrBln = TmpHariIni
                    Graph(12).intSkorGame = HitungSkor()
                End If
            Else
                'masukkan ke bagian terakhir
                Graph(XPos).bSudah = True
                Graph(XPos).intSkorGame = HitungSkor()
                Graph(XPos).StrBln = TmpHariIni
            End If
        End If
    End If
    
    'bagian pengatur waktu untuk ditampilkan ke status bar
    With StatusGame
        BackBuffer.SetForeColor RGB(0, 0, 0)
        BackBuffer.DrawText 9, 542, "Date :" & Format(Cuaca.HariIni, "ddd - dd,mmm,yyyy"), False
        BackBuffer.DrawText 9, 557, "Time :" & Format(CDate(Cuaca.JamSekarang & ":" & Cuaca.MenitSekarang), "hh:mm"), False
    
        BackBuffer.SetForeColor RGB(255, 255, 255)
        BackBuffer.DrawText 7, 540, "Date :" & Format(Cuaca.HariIni, "ddd - dd,mmm,yyyy"), False
        BackBuffer.DrawText 7, 555, "Time :" & Format(CDate(Cuaca.JamSekarang & ":" & Cuaca.MenitSekarang), "hh:mm"), False
    End With
End Sub
Sub CheckCraneMovement()
    'rutin ini untuk melacak pergerakan crane
    'berdasarkan penekanan tombol
    'periksa nilai xhead dan lakukan update
    If PAUSE_GAME Then Exit Sub
    
    With Crane
        If Not .IsArrivedX Then
            'crane belum berjalan, maka jalankan
            If Not .IsCraneMoving Then
                .IsCraneMoving = True
                Select Case .WayMove
                Case MOVE_LEFT
                    .MoveCrane MOVE_LEFT
                Case MOVE_RIGHT
                    .MoveCrane MOVE_RIGHT
                End Select
            ElseIf .IsCraneMoving Then
                .ContinueMovingX
                'cek apakah crane memuat tiang
                If .Occupied Then .XTiangOnArray = .XHead - (.LebarTiang \ 2)
            End If
            
        ElseIf (Not .IsArrivedY) Then
        
            'posisi X telah tiba maka giliran posisi Y
            'cek apakah pergerakan ke atas atau ke bawah
            If .YPoint > (.BaseY - .HeightCrane) + .LengthChain Then
                .LengthChain = .LengthChain + 1
            ElseIf .YPoint < (.BaseY - .HeightCrane) + .LengthChain Then
                .LengthChain = .LengthChain - 1
            ElseIf .YPoint = (.BaseY - .HeightCrane) + .LengthChain Then
                .YPoint = .YHead
            End If
            
            'cek apakah ukuran lengthchain telah mencapai ujung
            If .LengthChain > .HeightCrane Then
                .LengthChain = .HeightCrane
                .YPoint = .YHead
            ElseIf .LengthChain < 2 Then
                .LengthChain = 2
                .YPoint = .YHead
            End If
            
            If .Occupied Then .YTiangOnArray = .YHead + .LengthChain + 1
                
            'periksa apakah kepala berada di posisi baru
            If (.XHead = 18 Or .XHead = 17) And (.YHead + .LengthChain) = 145 Then
                'maka tampilkan rutin pemesanan vertikal beam
                PAUSE_GAME = True
                RedrawWindowTiang = True
                WindowTiang = True
                Crane.LebarTiang = 4
            End If
        End If
    End With
End Sub

Public Sub DrawHiasBox(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer)
    'untuk menggambar boks hias dengan parameter
    'X1,Y1      - Pojok Kiri atas
    'X2,Y2      - Pojok Kanan Bawah
    
    'pojok kiri atas
    BackBuffer.BltFast X1 * 7, Y1 * 7, BoxStatus, BoxRect(0, 0, 7, 7), DDBLTFAST_WAIT
    'garis horisontal atas dan bawah
    For X = X1 + 1 To X2 - 1
        BackBuffer.BltFast X * 7, Y1 * 7, BoxStatus, BoxRect(8, 0, 15, 7), DDBLTFAST_WAIT
        BackBuffer.BltFast X * 7, Y2 * 7, BoxStatus, BoxRect(8, 16, 15, 23), DDBLTFAST_WAIT
    Next X
    'pojok kanan atas
    BackBuffer.BltFast X2 * 7, Y1 * 7, BoxStatus, BoxRect(16, 0, 23, 7), DDBLTFAST_WAIT
    For X = Y1 + 1 To Y2 - 1
        BackBuffer.BltFast X1 * 7, (X * 7), BoxStatus, BoxRect(0, 8, 7, 15), DDBLTFAST_WAIT
        BackBuffer.BltFast X2 * 7, (X * 7), BoxStatus, BoxRect(16, 8, 23, 15), DDBLTFAST_WAIT
    Next X
    'pojok kiri bawah
    BackBuffer.BltFast X1 * 7, Y2 * 7, BoxStatus, BoxRect(0, 16, 7, 23), DDBLTFAST_WAIT
    'pojok kanan bawah
    BackBuffer.BltFast X2 * 7, Y2 * 7, BoxStatus, BoxRect(16, 16, 23, 23), DDBLTFAST_WAIT
    'mulai fill semua daerah kosong dengan warna orange
    BackBuffer.BltColorFill BoxRect(X1 * 7 + 2, Y1 * 7 + 2, X2 * 7, Y2 * 7), RGB(255, 0, 255)
End Sub

Public Sub DrawStatus()
    'Menggambar status bar beserta keterangannya
    'pojok kiri atas
    BackBuffer.BltFast 0, 521, BoxStatus, BoxRect(0, 0, 7, 7), DDBLTFAST_WAIT
    'garis horisontal atas dan bawah
    For X = 1 To 112
        BackBuffer.BltFast X * 7, 521, BoxStatus, BoxRect(8, 0, 15, 7), DDBLTFAST_WAIT
        BackBuffer.BltFast X * 7, 592, BoxStatus, BoxRect(8, 16, 15, 23), DDBLTFAST_WAIT
    Next X
    For X = 1 To 10
        BackBuffer.BltFast 0, (X * 7) + 521, BoxStatus, BoxRect(0, 8, 7, 15), DDBLTFAST_WAIT
        BackBuffer.BltFast 791, (X * 7) + 521, BoxStatus, BoxRect(16, 8, 23, 15), DDBLTFAST_WAIT
    Next X
    'pojok kanan atas
    BackBuffer.BltFast 791, 521, BoxStatus, BoxRect(16, 0, 23, 7), DDBLTFAST_WAIT
    'pojok kiri bawah
    BackBuffer.BltFast 0, 592, BoxStatus, BoxRect(0, 16, 7, 23), DDBLTFAST_WAIT
    'pojok kanan bawah
    BackBuffer.BltFast 791, 592, BoxStatus, BoxRect(16, 16, 23, 23), DDBLTFAST_WAIT
    'mulai fill semua daerah kosong dengan warna biru
    BackBuffer.BltColorFill BoxRect(7, 525, 792, 592), RGB(255, 0, 255)
End Sub


Public Sub DrawToolBar()
    'draw all toolbar to toolbar zone
    BackBuffer.BltFast 150, 530, ToolBarDX, BoxRect(0, (SelectedToolbar - 1) * 65, 500, SelectedToolbar * 65), DDBLTFAST_WAIT
    
    'cek apakah tab order berada di tab ke 3 (objektif)
    If SelectedToolbar = 3 Then
        BackBuffer.SetFont stFont
        BackBuffer.SetForeColor RGB(255, 255, 255)
        BackBuffer.DrawText 210, 559, "The Tower done before " & Format(Objektif.dHariSelesai, "dd-mmm-yyyy"), False
    End If
End Sub


Public Sub HandleToolbarZone()
    'berada di daerah toolbar
    TooltipStr = ""
    'Mouse_Button0 = False
    BackBuffer.SetForeColor RGB(255, 255, 255)
    'berada di daerah tab choice
    If CursorX > 165 And CursorX < 484 And CursorY < 551 Then
        If Mouse_Button0 Then
            'diclick left maka change tab order
            SelectedToolbar = ((CursorX - 150) \ 64) + 1
            'paksa menjadi nilai 5 jika
            SelectedToolbar = IIf(SelectedToolbar >= 6, 4, SelectedToolbar)
        End If
    ElseIf CursorX > 165 And CursorX < 484 And CursorY > 551 Then
        'berada di daerah pilihan icon
        Select Case SelectedToolbar
        Case 1  'tab tugas
            intLimitX = 486
        Case 2  'tab belanja
            intLimitX = 414
        Case 3  'tab objektif
            intLimitX = 198
        Case 4  'tab
            intLimitX = 270
        Case 5  'tab sistem
            intLimitX = 341
        End Select
        
        If CursorX > 180 And CursorX < intLimitX And CursorY > 559 And CursorY < 575 Then
            'highlight icon
            bHighlight = ((CursorX - 180) \ 18)
            'jika bernilai genap maka tampilkan box,
            'karena nilai ganjil berada di daerah antara icon
            If bHighlight Mod 2 = 0 Then
                BackBuffer.DrawBox (bHighlight * 18) + 176, 557, (bHighlight * 18) + 196, 579
                Select Case SelectedToolbar
                Case 1      'tab tugas
                    Select Case bHighlight
                    Case 0
                        TooltipStr = "Stop the worker"
                    Case 2
                        TooltipStr = "Make the required stuff in workshop"
                    Case 4
                        TooltipStr = "Build the Wall"
                    Case 6
                        TooltipStr = "Build Ladder"
                    Case 8
                        TooltipStr = "Build Door"
                    Case 10
                        TooltipStr = "Build Window"
                    Case 12
                        TooltipStr = "Paint the spot"
                    Case 14
                        TooltipStr = "Demolish the spot"
                    Case 16
                        TooltipStr = "Build foundation"
                    End Select
                Case 2      'tab Belanja
                    Select Case bHighlight
                    Case 0
                        TooltipStr = "Buy Brick"
                    Case 2
                        TooltipStr = "Buy Sand"
                    Case 4
                        TooltipStr = "Buy Cement"
                    Case 6
                        TooltipStr = "Buy Paint"
                    Case 8
                        TooltipStr = "Buy Lumber"
                    Case 10
                        TooltipStr = "Buy Iron"
                    Case 12
                        TooltipStr = "Buy Glass"
                    End Select
                Case 3
                    Select Case bHighlight
                    Case 0
                        TooltipStr = "Show the Objective"
                    End Select
                Case 4      'budget
                    Select Case bHighlight
                    Case 0
                        TooltipStr = "Unuseable"
                    Case 2
                        TooltipStr = "See the budget"
                    Case 4
                        TooltipStr = "Unuseable"
                    End Select
                Case 5
                    Select Case bHighlight
                    Case 0
                        TooltipStr = "Save the progressing game"
                    Case 2
                        TooltipStr = "Exit the game"
                    Case 4
                        TooltipStr = "Show hints"
                    Case 6
                        TooltipStr = "See the tutorial"
                    Case 8
                        TooltipStr = "Toggle On/Off sound effect and music"
                    End Select
                End Select
                'tuliskan tooltip ke layar backbuffer
                
                BackBuffer.DrawText 0, 505, TooltipStr, False
                
            End If
            'periksa penekanan tombol kiri mouse
            If Mouse_Button0 Then
                Mouse_Button0 = False
                Select Case SelectedToolbar
                Case 1      'tab tugas
                    Select Case bHighlight
                    Case 0      'tombol STOP
                        If GlobalSel Then       'ada pekerja terpilih
                            'hentikan pekerja dari pekerjaan apapun sekarang
                            If Worker(SelNo).Status = WORK Or Worker(SelNo).Status = Walk Then
                                Worker(SelNo).XPoint = Worker(SelNo).XTileOnArray
                                Worker(SelNo).YPoint = Worker(SelNo).YTileOnArray
                                Worker(SelNo).Job = STAND
                                Worker(SelNo).Status = IDLE
                                Worker(SelNo).Progress = 0
                            End If
                        End If
                    Case 2      'tombol pembuatan
                        PAUSE_GAME = True
                        WindowMaker = True
                        RedrawWindowMaker = True
                        Workshop.bytTmpTangga = 0
                        Workshop.bytTmpJendela = 0
                        Workshop.bytTmpPintu = 0
                        Workshop.bytTmpBesi = Workshop.bytBesi
                        Workshop.bytTmpKaca = Workshop.bytKaca
                        Workshop.bytTmpKayu = Workshop.bytKayu
                    Case 4      'tombol batu bata
                        'set false terhadap segala bentuk pemasangan barang lain
                        PasangTangga = False
                        PasangTiang = False
                        Penghancur = False
                        PasangJendela = False
                        PasangCat = False
                        
                        'tombol ini akan memilih pekerja dengan melacak pemilihan globalsel
                        If Not GlobalSel Then
                            'tidak ada worker terpilih
                            PAUSE_GAME = True
                            ShowError.ErrorWindow = True
                            ShowError.RedrawError = True
                            ShowError.StrKata = "Tidak Ada Worker terpilih !!"
                            Exit Sub
                        End If
                        
                        'maka tampilkan sprite tembok dan pilih click di daerah untuk dikerjakan pekerja
                        PASANGBATU = True
                        
                    Case 6      'tombol tangga
                        'set false terhadap segala bentuk pekerjaan barang lain
                        PasangTiang = False
                        PASANGBATU = False
                        Penghancur = False
                        PasangJendela = False
                        PasangCat = False
                        
                        If Not GlobalSel Then
                            'tidak ada worker terpilih
                            PAUSE_GAME = True
                            ShowError.ErrorWindow = True
                            ShowError.RedrawError = True
                            ShowError.StrKata = "Tidak Ada Worker terpilih !!"
                            Exit Sub
                        End If
                        
                        'maka tampilkan sprite tangga  dan pilih click di daerah untuk dikerjakan pekerja
                        PasangTangga = True
                        
                    Case 10     'Jendela
                        'set false terhadap segala bentuk pekerjaan barang lain
                        PasangTiang = False
                        PASANGBATU = False
                        Penghancur = False
                        PasangTangga = False
                        PasangCat = False
                        
                        If Not GlobalSel Then
                            'tidak ada worker terpilih
                            PAUSE_GAME = True
                            ShowError.ErrorWindow = True
                            ShowError.RedrawError = True
                            ShowError.StrKata = "Tidak Ada Worker terpilih !!"
                            Exit Sub
                        End If
                        
                        'maka tampilkan sprite tangga  dan pilih click di daerah untuk dikerjakan pekerja
                        PasangJendela = True
                    Case 12     'cat
                        'stel yang lain menjadi false
                        PASANGBATU = False
                        PasangTangga = False
                        PasangTiang = False
                        Penghancur = False
                        PasangJendela = False
                        
                        If Not GlobalSel Then
                            'tidak ada worker terpilih
                            PAUSE_GAME = True
                            ShowError.ErrorWindow = True
                            ShowError.RedrawError = True
                            ShowError.StrKata = "Tidak ada worker terpilih !!"
                            Exit Sub
                        End If
                        
                        PasangCat = True
                        
                    Case 14     'penghancur
                        'stel yang lain menjadi false
                        PASANGBATU = False
                        PasangTangga = False
                        PasangTiang = False
                        PasangCat = False
                        PasangJendela = False
                        Penghancur = False
                        
                        'If Not GlobalSel Then
                            'tidak ada worker terpilih
                            PAUSE_GAME = True
                            ShowError.ErrorWindow = True
                            ShowError.RedrawError = True
                            ShowError.StrKata = "Temporary disabled !!"
                            Exit Sub
                        'End If
                        
                        'Penghancur = True
                    
                    Case 16     'Pasang Tiang
                        'para pekerja yang terpilih dinonaktifkan terlebih dahulu
                        GlobalSel = False
                        SelNo = 0
                        'tidak perlu memerlukan pekerja, maka langsung pasang tiang secara vertikal
                        PasangTangga = False
                        Penghancur = False
                        PASANGBATU = False
                        PasangJendela = False
                        PasangTiang = True
                        
                    End Select
                    
                Case 2      'tab Belanja
                    Select Case bHighlight
                    Case 0          'belanja batu bata
                        PAUSE_GAME = True
                        WindowBelanjaBata = True
                        RedrawWindowBata = True
                        Workshop.bytTmpBata = 0
                    Case 2          'belanja pasir
                        PAUSE_GAME = True
                        WindowBelanjaPasir = True
                        RedrawWindowPasir = True
                        Workshop.bytTmpPasir = 0
                    Case 4          'belanja semen
                        PAUSE_GAME = True
                        WindowBelanjaSemen = True
                        RedrawWindowSemen = True
                        Workshop.bytTmpSemen = 0
                    Case 6          'belanja Cat
                        PAUSE_GAME = True
                        WindowBelanjaCat = True
                        RedrawWindowCat = True
                        Workshop.bytTmpCat = 0
                    Case 8          'belanja Kayu
                        PAUSE_GAME = True
                        WindowBelanjaKayu = True
                        RedrawWindowKayu = True
                        Workshop.bytTmpKayu = 0
                    Case 10          'belanja besi
                        PAUSE_GAME = True
                        WindowBelanjaBesi = True
                        RedrawWindowBesi = True
                        Workshop.bytTmpBesi = 0
                    Case 12
                        PAUSE_GAME = True
                        WindowBelanjaKaca = True
                        RedrawWindowKaca = True
                        Workshop.bytTmpKaca = 0
                    End Select
                    
                Case 3      'tab objektif
                    Select Case bHighlight
                    Case 0
                        PAUSE_GAME = True
                        RedrawWindowSketch = True
                        WindowSketch = True
                    End Select
                Case 4      'tab
                    Select Case bHighlight
                    Case 2      'window grafik
                        PAUSE_GAME = True
                        RedrawWindowGraph = True
                        WindowGraph = True
                    End Select
                
                Case 5
                    Select Case bHighlight
                    Case 0      'save
                        'prosedure penyimpanan permainan
                        Call SaveGame
                        'tampilkan pesan sponsor
                        PAUSE_GAME = True
                        ShowError.ErrorWindow = True
                        ShowError.RedrawError = True
                        ShowError.StrKata = "Game has been saved "
                    Case 2      'Exit Game
                        Call EndAll
                    Case 4      'Tips permainan
                        PAUSE_GAME = True
                        RedrawWindowTips = True
                        WindowTips = True
                        Call InitTips
                        NoTips = 0
                    Case 6      'Help Permainan
                        PAUSE_GAME = True
                        RedrawWindowHelp = True
                        WindowHelp = True
                    Case 8
                        MusikOn = Not MusikOn
                        If MusikOn Then
                            If SFXMusik.InitDM Then
                                If SFXMusik.LoadMusic(App.Path & "\music\hotsteel.mid") Then
                                    SFXMusik.PlayMusic
                                End If
                            End If
                        Else
                            SFXMusik.StopMusic
                        End If
                    End Select
                End Select
            End If
        End If
        
    End If
End Sub

Public Sub HandleGameZone()
    Selected = False
    If PAUSE_GAME Then Exit Sub
    
    For intX = 0 To intLemm
        If Worker(intX).Visible Then
        If Worker(intX).XPosition <= CursorX And Worker(intX).XPosition + TILEWIDTH >= CursorX And Worker(intX).YPosition <= CursorY And Worker(intX).YPosition + TILEHEIGHT >= CursorY Then
            BackBuffer.SetForeColor RGB(255, 255, 255)
            BackBuffer.setDrawWidth 1
            BackBuffer.DrawBox Worker(intX).XPosition - 2 + Worker(intX).XHeadSmooth, Worker(intX).YPosition - 2, Worker(intX).XPosition + TILEWIDTH + 2 + Worker(intX).XHeadSmooth, Worker(intX).YPosition + TILEHEIGHT + 2
            
            'ganti bentuk kursor
            MousePointer = MOUSE_POINT
            
            'selected worker on
            Selected = True
            
            'tampilkan captions
            With stFont
                .Bold = False
                .Size = 7
                .Name = "Arial"
            End With
            'set to backbuffer
            BackBuffer.SetFont stFont
            BackBuffer.SetForeColor RGB(255, 255, 255)
            BackBuffer.DrawText 4, 3, "No : " & Worker(intX).NoLemm, False
            BackBuffer.DrawText 4, 13, "Name : " & Worker(intX).Nama, False
            BackBuffer.DrawText 4, 23, "Job : " & Worker(intX).Job, False
            BackBuffer.DrawText 4, 33, "Honor : " & Worker(intX).Honor, False
            BackBuffer.DrawText 4, 43, "Stability : " & Worker(intX).Stability, False
            BackBuffer.DrawText 4, 53, "Tolerance : " & Worker(intX).Tolerance, False
            BackBuffer.DrawBox 4, 66, 84, 72
            'update progress bar
            BackBuffer.BltColorFill BoxRect(5, 67, 5 + Int((Worker(intX).Progress / 100) * 78), 71), RGB(100, 250, 0)

            'keluar dari looping agar speed tidak berkurang
            Exit For
        End If
        End If
    Next intX
    
    'tidak ada worker terpilih , maka ganti mouse
    If intX > intLemm Then
        MousePointer = MOUSE_DEFAULT
    End If
    
    'BAGIAN PEKERJA
    If Mouse_Button0 And Selected Then
        'mouse kiri ditekan dan pilihan adalah pekerja
        'there are worker selected, so point the exact pointer up the object
        GlobalSel = True
        SelNo = intX
    ElseIf Mouse_Button0 And Not Selected Then
        'yang terpilih adalah daerah kosong
        GlobalSel = False
        SelNo = 0
    'bagian worker bekerja dengan menggunakan tombol kanan mouse
    ElseIf Mouse_Button1 And GlobalSel And (Worker(SelNo).Status = IDLE) Then
        'mouse kanan ditekan
        'SelNo = intX
        'maka masukkan nilai variabel untuk pergerakan karakter
        Dim TmpX As Byte
        Dim TmpY As Byte
        
        TmpX = Screen.GetTileX
        TmpY = Screen.GetTileY
        
        If Worker(SelNo).XTileOnArray < TmpX Then
            If Worker(SelNo).MoveRight(Worker(SelNo).XTileOnArray, Worker(SelNo).YTileOnArray) Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = Walk
                Worker(SelNo).WayMove = MOVE_RIGHT
            ElseIf Worker(SelNo).YTileOnArray < TmpY Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = CLIMB
                Worker(SelNo).WayMove = MOVE_DOWN
            ElseIf Worker(SelNo).YTileOnArray > TmpY Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = CLIMB
                Worker(SelNo).WayMove = MOVE_UP
            End If
            
        ElseIf Worker(SelNo).XTileOnArray > TmpX Then
            If Worker(SelNo).MoveLeft(Worker(SelNo).XTileOnArray, Worker(SelNo).YTileOnArray) Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = Walk
                Worker(SelNo).WayMove = MOVE_LEFT
            ElseIf Worker(SelNo).YTileOnArray < TmpY Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = CLIMB
                Worker(SelNo).WayMove = MOVE_DOWN
            ElseIf Worker(SelNo).YTileOnArray > TmpY Then
                Worker(SelNo).XPoint = Screen.GetTileX
                Worker(SelNo).YPoint = Screen.GetTileY
                Worker(SelNo).Status = CLIMB
                Worker(SelNo).WayMove = MOVE_UP
            End If
        End If
        
        'lakukan pencarian move
        If Worker(SelNo).YTileOnArray < Worker(SelNo).YPoint Then
            Worker(SelNo).SearchMove = MOVE_DOWN
        ElseIf Worker(SelNo).YTileOnArray > Worker(SelNo).YPoint Then
            Worker(SelNo).SearchMove = MOVE_UP
        ElseIf Worker(SelNo).XTileOnArray > Worker(SelNo).XPoint Then
            Worker(SelNo).SearchMove = MOVE_LEFT
        ElseIf Worker(SelNo).XTileOnArray < Worker(SelNo).XPoint Then
            Worker(SelNo).SearchMove = MOVE_RIGHT
        'ElseIf Worker(SelNo).YTileOnArray = Worker(SelNo).YPoint Then
            'tidak perlu dicari langkah
        '    Worker(SelNo).SearchMove = MOVE_NONE
        End If
    End If
    
    'BAGIAN CRANE
    If Mouse_Button0 And Crane.PickCockpit Then
        'mouse kiri ditekan dan kepala crane yang terpilih
        Crane.CranePilih = True
        'CraneSel = True
    ElseIf Mouse_Button1 And Crane.CranePilih Then
        Mouse_Button1 = False
        
        If (Crane.XHead < (Crane.LengthCrane + Crane.BaseX) And (Screen.GetTileX > Crane.XHead)) Or ((Crane.XHead > Crane.BaseX + 2) And (Screen.GetTileX < Crane.XHead)) Then
        
            Crane.XPoint = Screen.GetTileX
            Crane.YPoint = Screen.GetTileY
            Crane.PointerDraw = True
            Crane.PointerLong = 50
            
            'cek dulu apakah kepala sudah mencapai ujung kiri maupun kanan
            If Crane.XHead < Crane.XPoint Then
                Crane.WayMove = MOVE_RIGHT
                If Crane.XPoint >= Crane.BaseX + Crane.LengthCrane Then
                    Crane.XPoint = Crane.BaseX + Crane.LengthCrane
                End If
            Else
                Crane.WayMove = MOVE_LEFT
                If Crane.XPoint <= Crane.BaseX + 2 Then
                    Crane.XPoint = Crane.BaseX + 2
                End If
            End If
        End If
    End If
    
    'Bagian informasi Workshop
    If Mouse_Button0 And Screen.GetTileY = 142 And (Screen.GetTileX = 146 Or Screen.GetTileX = 147 Or Screen.GetTileX = 148) Then
        'Tampilkan informasi di dalam workshop
        PAUSE_GAME = True
        WindowWorkShop = True
        RedrawWindowWorkshop = True
    End If
    
    'BAGIAN PEMASANGAN PEKERJAAN
    If Mouse_Button0 And PasangTiang Then
        intvarx = Screen.GetTileX
        intvary = Screen.GetTileY
        
        If Not ((ArGedung(intvarx, intvary).bSpace = True And ArGedung(intvarx, intvary + 1).bTiang = True) Or _
            (ArGedung(intvarx, intvary).bSpace = True And ArGedung(intvarx, intvary + 1).bGround = True)) Then
            ShowError.ErrorWindow = True
            ShowError.StrKata = "Tiang tidak bisa diletakkan !"
            ShowError.RedrawError = True
            PasangTiang = False
            Exit Sub
        End If
        
        'mulai letakkan isi tiang ke dalam array gedung untuk diproses
        Mouse_Button0 = False
        ArGedung(Screen.GetTileX, Screen.GetTileY).bytSprX = 4
        ArGedung(Screen.GetTileX, Screen.GetTileY).bytSprY = 2
        ArGedung(Screen.GetTileX, Screen.GetTileY).bSpace = False
        ArGedung(Screen.GetTileX, Screen.GetTileY).bTiang = True
        'jalankan proses pengecoran
        Call FillProgToTiang(Screen.GetTileX, Screen.GetTileY)
        Call DrawMoneyOnScreen(Screen.GetTileX, Screen.GetTileY, "-10", MOVE_RIGHT)
        
        Screen.Redraw = True
    ElseIf Mouse_Button1 = True And PasangTiang Then
        Mouse_Button1 = False
        PasangTiang = False
    
    ElseIf Mouse_Button0 = True And PASANGBATU Then
        'drop pasangbatu
        Mouse_Button0 = False
        PASANGBATU = False
        
    ElseIf Mouse_Button1 = True And PASANGBATU And LolosDinding(Screen.GetTileX, Screen.GetTileY) Then
        Mouse_Button1 = False
        PASANGBATU = False
        'pertama - tama tentukan pekerjaan untuk worker terpilih
        If Not BangunTembok Then
            PAUSE_GAME = True
            ShowError.ErrorWindow = True
            ShowError.StrKata = "Composition material not enough !"
            ShowError.RedrawError = True
            PASANGBATU = False
            Exit Sub
        End If
        
        If Worker(SelNo).Status <> WORK And Abs(Worker(SelNo).YTileOnArray - Screen.GetTileY) < 3 Then
            Worker(SelNo).Job = PASANG_BATU
            Worker(SelNo).JobX = Screen.GetTileX
            Worker(SelNo).JobY = Screen.GetTileY
        End If
        
    ElseIf Mouse_Button0 = True And PasangTangga Then
        Mouse_Button0 = False
        PasangTangga = False
    
    ElseIf Mouse_Button1 = True And PasangTangga Then
        Mouse_Button1 = False
        PasangTangga = False
        
        If Workshop.bytTangga < 1 Then
            PAUSE_GAME = True
            ShowError.ErrorWindow = True
            ShowError.StrKata = "Ladder not enough !"
            ShowError.RedrawError = True
            Exit Sub
        End If
        
        'pertama - tama tentukan pekerjaan untuk worker terpilih
        Worker(SelNo).Job = PASANG_TANGGA
        Worker(SelNo).JobX = Screen.GetTileX
        Worker(SelNo).JobY = Screen.GetTileY
        Workshop.bytTangga = Workshop.bytTangga - 1
    
    ElseIf Mouse_Button0 = True And Penghancur Then
        Mouse_Button0 = False
        Penghancur = False
        
    'cek apakah daerah bukan merupakan space dan bisa dihancurkan
    ElseIf Mouse_Button1 = True And Penghancur Then
        Mouse_Button0 = False
        Penghancur = False
        
        intvarx = Screen.GetTileX
        intvary = Screen.GetTileY
        If Not ArGedung(intvarx, intvary).bSpace Then
            'tentukan job untuk worker yang terpilih
            Worker(SelNo).Job = HANCUR
            Worker(SelNo).JobX = Screen.GetTileX
            Worker(SelNo).JobY = Screen.GetTileY
        End If
        
    ElseIf Mouse_Button0 = True And PasangCat Then
        Mouse_Button0 = False
        PasangCat = False
        
    'cek apakah daerah bukan merupakan space dan bisa dihancurkan
    ElseIf Mouse_Button1 = True And PasangCat Then
        Mouse_Button0 = False
        PasangCat = False
        
        If Workshop.bytCat < 1 Then
            PAUSE_GAME = True
            ShowError.ErrorWindow = True
            ShowError.StrKata = "Paint not enough !"
            ShowError.RedrawError = True
            Exit Sub
        End If
        
        intvarx = Screen.GetTileX
        intvary = Screen.GetTileY
        If (ArGedung(intvarx, intvary).bytSprX = 4 And ArGedung(intvarx, intvary).bytSprY = 3) Or _
        ArGedung(intvarx, intvary).bytSprX = 5 And ArGedung(intvarx, intvary).bytSprY = 3 Or _
        ArGedung(intvarx, intvary).bytSprX = 0 And ArGedung(intvarx, intvary).bytSprY = 4 Then
            Exit Sub
        End If
        
        If Not ArGedung(intvarx, intvary).bSpace Then
            'tentukan job untuk worker yang terpilih
            Worker(SelNo).Job = PASANG_CAT
            Worker(SelNo).JobX = Screen.GetTileX
            Worker(SelNo).JobY = Screen.GetTileY
            Workshop.bytCat = Workshop.bytCat - 1
        End If
    
    ElseIf Mouse_Button0 = True And PasangJendela Then
        Mouse_Button0 = False
        PasangJendela = False
        
    'cek apakah daerah bukan merupakan space dan bisa dihancurkan
    ElseIf Mouse_Button1 = True And PasangJendela Then
        Mouse_Button0 = False
        PasangJendela = False
        
        If Workshop.bytJendela < 1 Then
            PAUSE_GAME = True
            ShowError.ErrorWindow = True
            ShowError.StrKata = "Jumlah jendela tidak mencukupi !"
            ShowError.RedrawError = True
            Exit Sub
        End If
        
        intvarx = Screen.GetTileX
        intvary = Screen.GetTileY
        'tentukan job untuk worker yang terpilih
        Worker(SelNo).Job = PASANG_JENDELA
        Worker(SelNo).JobX = Screen.GetTileX
        Worker(SelNo).JobY = Screen.GetTileY
        Workshop.bytJendela = Workshop.bytJendela - 1
    End If
    
    'bagian minimap
    'agar minimap dapat digunakan untuk melakukan perubahan layar secara langsung
    'tanpa melalui penggulungan
    If ((CursorX > 693 And CursorX < 798) And (CursorY > 0 And CursorY < 105)) And Mouse_Button0 = True Then
        Mouse_Button0 = False
        Dim TmpVar As Integer
        Dim TmpVar2 As Integer
        TmpVar = Int((CursorX - 693) * SkalaMap)
        If TmpVar - 28 <= 0 Then
            Gedung.XGedung = 1
        ElseIf TmpVar + 28 > MAXGEDUNGX Then
            Gedung.XGedung = 93
        Else
            Gedung.XGedung = TmpVar - 28
        End If
        
        TmpVar2 = Int(CursorY * SkalaMap)
        If TmpVar2 - 14 <= 0 Then
            Gedung.YGedung = 1
        ElseIf TmpVar2 + 14 > MAXGEDUNGY Then
            Gedung.YGedung = 122
        Else
            Gedung.YGedung = (TmpVar2 - 14)
        End If
        
        Screen.Redraw = True
        BackBuffer.DrawText 100, 300, (CursorX - 693) & "," & CursorY, False
    End If
End Sub
Sub ShowWindowTips()
    If RedrawWindowTips Then
    
    RedrawWindowTips = False
    
    TempDX.SetForeColor RGB(255, 255, 255)
    
    TempDX.BltColorFill BoxRect(205, 205, 605, 405), QBColor(0)
    TempDX.DrawBox 200, 200, 600, 400
    TempDX.BltColorFill BoxRect(201, 201, 599, 399), RGB(0, 150, 100)
    TempDX.BltColorFill BoxRect(202, 202, 598, 220), RGB(0, 100, 200)
    TempDX.BltColorFill BoxRect(205, 225, 595, 365), RGB(200, 200, 0)
    
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = True
    End With
    TempDX.SetFont stFont
    TempDX.SetFontTransparency True
    TempDX.SetForeColor RGB(0, 0, 0)
    TempDX.DrawText 350, 202, "Tips and Help ", False
    'gambar OK
    TempDX.DrawRoundedBox 510, 370, 570, 390, 4, 4
    TempDX.DrawRoundedBox 430, 370, 460, 390, 4, 4
    TempDX.DrawRoundedBox 470, 370, 500, 390, 4, 4

    TempDX.DrawText 529, 372, "OK", False
    TempDX.DrawText 475, 372, ">>", False
    TempDX.DrawText 435, 372, "<<", False
    
    With stFont
        .Name = "Tahoma"
        .Size = 9
        .Bold = False
    End With
    TempDX.SetFont stFont
    TempDX.SetForeColor RGB(255, 255, 255)
    TempDX.DrawText 208, 228, TipsGame(NoTips), False
    
    End If
    
    BackBuffer.BltFast 200, 200, TempDX, BoxRect(200, 200, 605, 405), DDBLTFAST_WAIT
    
    'Cek CursorX dan CursorY mouse
    
    If Mouse_Button0 Then
    Mouse_Button0 = False
    Select Case CursorX
    Case 510 To 570     'tombol OK
    Select Case CursorY
    Case 370 To 390
        'tombol OK ditekan
        PAUSE_GAME = False
        WindowTips = False
    End Select
    
    Case 470 To 500
    Select Case CursorY
    Case 370 To 390
        'tombol >> ditekan
        If NoTips + 1 > 10 Then NoTips = 0 Else NoTips = NoTips + 1
        RedrawWindowTips = True
    End Select
    
    
    Case 430 To 460
    Select Case CursorY
    Case 370 To 390
        'tombol << ditekan
        If NoTips - 1 < 0 Then NoTips = 10 Else NoTips = NoTips - 1
        RedrawWindowTips = True
    End Select
    
    End Select
    End If
End Sub

Public Sub SkalaMiniMap()
    'Rutin ini akan menghitung skala pembagian antara
    'ukuran asli array dengan ukuran minimap
    If MAXGEDUNGX > MAXGEDUNGY Then
        'skala diambil dari maxgedungx
        SkalaMap = MAXGEDUNGX / 105
    Else
        'skala diambil dari maxgedungy
        SkalaMap = MAXGEDUNGY / 105
    End If
    'maka besar dari kotak wakil adalah sebesar
    'maxtilex\skalamap dan maxtiley\skalamap
    WidthMap = MAXTILEX / SkalaMap
    HeightMap = MAXTILEY / SkalaMap
    
    'StartX = MAXGEDUNGX / SkalaMap
    'StartY = MAXGEDUNGY / SkalaMap
End Sub


Public Sub UpdateMiniMap()
    'gambar sebelah pojok kiri layar pilihan
    'perhitungan dimulai dengan menggunakan pembagian panjang dan lebar mini map
    'dengan skala sesuai dengan maxgedungx dan maxgedungy
    'besar dari minimap adalah 100 x 100 pixel sehingga rumus skala adalah
    'maxgedungx / 100 dan maxgedungy / 100, dengan prinsip nilai maxgedungx atau maxgedungy harus paling besar diambil
    'dan dalam satu layar adalah sebesar 28 baris x 57 kolom petak
    
    'cari nilai gedung.xgedung dan gedung.ygedung sebagai patokan penggambaran kotak
    JarakX = Gedung.XGedung / SkalaMap
    JarakY = Gedung.YGedung / SkalaMap
    
    'BackBuffer.BltColorFill BoxRect(693 + CByte(JarakX), CByte(JarakY) + 1, 693 + (CByte(JarakX + WidthMap) - 1), CByte(JarakY + HeightMap) - 1), RGB(255, 128, 0)
    BackBuffer.DrawBox 693 + CByte(JarakX), CByte(JarakY), 693 + CByte(JarakX + WidthMap), CByte(JarakY + HeightMap)
    
    'update terhadap worker
    For intX = 1 To intLemm
        BackBuffer.BltFast 693 + CByte(Worker(intX).XTileOnArray / SkalaMap), CByte(Worker(intX).YTileOnArray / SkalaMap), Sprite, BoxRect(24, 0, 25, 1), DDBLTFAST_WAIT
    Next intX
    
End Sub
Sub InitWorker()
    intLemm = 12
    Dim Y As Byte
    For Y = 1 To intLemm
        Worker(Y).NoLemm = Y
        Worker(Y).Nama = RndName()
        Worker(Y).XTileOnArray = 10
        Worker(Y).YTileOnArray = 146
        Worker(Y).Stability = 90
        Worker(Y).Tolerance = Int(Rnd * 7) + 3
        Worker(Y).Perubah = 2
        Worker(Y).Active = True
        Worker(Y).WalkSpeed = 2
        Worker(Y).WorkSpeed = 25
        Worker(Y).Job = STAND
        Worker(Y).Status = STAND
        Worker(Y).XPoint = Worker(1).XTileOnArray
        Worker(Y).YPoint = Worker(1).YTileOnArray
        Worker(Y).WayMove = MOVE_RIGHT
    Next Y
End Sub
Public Sub UpdateStatusBar()
    'melakukan update terhadap status game
    With StatusGame
        With stFont
            .Bold = False
            .Size = 8
            .Name = "Comic Sans MS"
        End With
        'set to backbuffer
        BackBuffer.setDrawWidth 1
        BackBuffer.SetFont stFont
        BackBuffer.SetForeColor RGB(0, 0, 0)
        BackBuffer.SetForeColor RGB(255, 255, 255)
        DrawHiasBox 0, 0, 13, 10
    End With
End Sub

Sub UpdateCraneCockpit()
    Call Crane.UpdateCockpit
    If Crane.Visible Then BackBuffer.BltFast Crane.XCock, Crane.YCock, Sprite, BoxRect(14, 20, 28, 39), DDBLTFAST_WAIT
End Sub

Public Sub DrawOrangeBox(X1 As Integer, Y1 As Integer, X2 As Integer, Y2 As Integer, LayarTmp As DirectDrawSurface7)
    'untuk menggambar boks hias dengan parameter
    'X1,Y1      - Pojok Kiri atas
    'X2,Y2      - Pojok Kanan Bawah
    
    'pojok kiri atas
    LayarTmp.BltFast (X1 * 7), (Y1 * 7), BoxStatus, BoxRect(24, 0, 31, 7), DDBLTFAST_WAIT
    'garis horisontal atas dan bawah
    For X = X1 + 1 To X2 - 1
        LayarTmp.BltFast X * 7, Y1 * 7, BoxStatus, BoxRect(32, 0, 39, 7), DDBLTFAST_WAIT
        LayarTmp.BltFast X * 7, Y2 * 7, BoxStatus, BoxRect(32, 16, 39, 23), DDBLTFAST_WAIT
    Next X
    'pojok kanan atas
    LayarTmp.BltFast X2 * 7, Y1 * 7, BoxStatus, BoxRect(16, 0, 23, 7), DDBLTFAST_WAIT
    For X = Y1 + 1 To Y2 - 1
        LayarTmp.BltFast X1 * 7, (X * 7), BoxStatus, BoxRect(0, 8, 7, 15), DDBLTFAST_WAIT
        LayarTmp.BltFast X2 * 7, (X * 7), BoxStatus, BoxRect(16, 8, 23, 15), DDBLTFAST_WAIT
    Next X
    'pojok kiri bawah
    LayarTmp.BltFast X1 * 7, Y2 * 7, BoxStatus, BoxRect(0, 16, 7, 23), DDBLTFAST_WAIT
    'pojok kanan bawah
    LayarTmp.BltFast X2 * 7, Y2 * 7, BoxStatus, BoxRect(16, 16, 23, 23), DDBLTFAST_WAIT
    'mulai fill semua daerah kosong dengan warna orange
    LayarTmp.BltColorFill BoxRect(X1 * 7 + 2, Y1 * 7 + 2, (X2 * 7) + 5, (Y2 * 7) + 5), RGB(150, 0, 0)
End Sub

Sub UpdateWorker()
    'this routine greatly update all workers on game
    Dim Ty As Integer
    For Ty = 1 To intLemm
        With Worker(Ty)
        'jika worker nampak maka update posisi dan frame worker
        If .Active And .Visible Then
            If Not PAUSE_GAME Then .UpdateFrame
            .UpdatePosition
        End If
        Select Case .Status
        Case REST
            If .Visible Then
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, PANIC), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
        Case IDLE
            If .Visible Then
            Select Case .WayMove
            Case MOVE_LEFT
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, STAND_RIGHT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Case MOVE_RIGHT
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, STAND_LEFT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Case MOVE_UP
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, STAND_LEFT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Case MOVE_DOWN
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, STAND_RIGHT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End Select
            End If
            
        Case Walk
            .UpdatePosition
            If .Visible Then
            Select Case .WayMove
            Case MOVE_RIGHT
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, WALK_LEFT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            Case MOVE_LEFT
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, WALK_RIGHT), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End Select
            End If
            
        Case CLIMB
            If .Visible Then
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, UP_DOWN), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            
        Case WORK
            If .Visible Then
                BackBuffer.BltFast .XPosition + .XHeadSmooth, .YPosition + .YHeadSmooth, LemmDX, ArSpr(.Frame, BUILDER), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
            End If
            
            If .WorkNow >= .WorkSpeed And Not PAUSE_GAME Then
                'lakukan update terhadap progress pekerja
                .Progress = .Progress + 1
                If .Progress >= 100 Then
                    'maka kerjaan selesai
                    .Progress = 0
                    Select Case .Job
                    Case PASANG_BATU
                        intvarx = .JobX
                        intvary = .JobY
                        ArGedung(intvarx, intvary).bytSprX = 0
                        ArGedung(intvarx, intvary).bytSprY = 0
                        ArGedung(intvarx, intvary).bSpace = False
                        ArGedung(intvarx, intvary).bWall = True
                        Screen.Redraw = True
                        RedrawMiniMap = True
                        .Status = IDLE
                        .Job = STAND
                    Case PASANG_TANGGA
                        intvarx = .JobX
                        intvary = .JobY
                        'cek apakah posisi tangga berada di daerah objek lain
                        'apakah merupakan tembok ?
                        If ArGedung(intvarx, intvary).bWall And ArGedung(intvarx, intvary).bCat Then
                            ArGedung(intvarx, intvary).bytSprX = 0
                            ArGedung(intvarx, intvary).bytSprY = 5
                        ElseIf ArGedung(intvarx, intvary).bWall And Not ArGedung(intvarx, intvary).bCat Then
                            ArGedung(intvarx, intvary).bytSprX = 1
                            ArGedung(intvarx, intvary).bytSprY = 4
                        ElseIf ArGedung(intvarx, intvary).bCor Then
                            ArGedung(intvarx, intvary).bytSprX = 1
                            ArGedung(intvarx, intvary).bytSprY = 5
                        End If
                        ArGedung(intvarx, intvary).bSpace = False
                        ArGedung(intvarx, intvary).bLadder = True
                        Screen.Redraw = True
                        RedrawMiniMap = True
                        .Status = IDLE
                        .Job = STAND
                    Case PASANG_JENDELA
                        intvarx = .JobX
                        intvary = .JobY
                        ArGedung(intvarx, intvary).bytSprX = 4
                        ArGedung(intvarx, intvary).bytSprY = 3
                        ArGedung(intvarx, intvary).bSpace = False
                        ArGedung(intvarx, intvary).bWall = True
                        Screen.Redraw = True
                        RedrawMiniMap = True
                        .Status = IDLE
                        .Job = STAND
                    Case PASANG_CAT
                        intvarx = .JobX
                        intvary = .JobY
                        
                        If ArGedung(intvarx, intvary).bLadder Then
                            ArGedung(intvarx, intvary).bytSprX = 0
                            ArGedung(intvarx, intvary).bytSprY = 5
                        Else
                            ArGedung(intvarx, intvary).bytSprX = 1
                            ArGedung(intvarx, intvary).bytSprY = 0
                        End If
                        
                        ArGedung(intvarx, intvary).bSpace = False
                        ArGedung(intvarx, intvary).bCat = True
                        Screen.Redraw = True
                        RedrawMiniMap = True
                        .Status = IDLE
                        .Job = STAND
                    End Select
                    
                End If
                .WorkNow = 0
            ElseIf Not PAUSE_GAME And .WorkNow < .WorkSpeed Then
                .WorkNow = .WorkNow + 1
            End If
        End Select
        End With
        'BLOK PEMERIKSA PERJALANAN WORKER
        If Not PAUSE_GAME Then
        
        With Worker(Ty)
            If Not .IsArrived Then
                'worker belum berjalan, maka jalankan
                If Not .IsMoving Then
                    .IsMoving = True
                    .MoveWorker
                ElseIf .IsMoving Then
                    .ContinueMoving
                End If
            Else
                'periksa status pekerjaan pekerja sesuai dengan job diberikan
                Select Case .Job
                Case STAND
                    'belum ada pekerjaan maka berhentikan pekerja
                    If .Stability > 0 Then
                        .Status = IDLE
                    ElseIf .Stability = 0 Then
                        .Status = REST
                    End If
                Case PASANG_BATU, PASANG_TANGGA, PASANG_JENDELA, PASANG_CAT
                    .Status = WORK
                End Select
            End If
        End With
        End If
    Next Ty
End Sub


