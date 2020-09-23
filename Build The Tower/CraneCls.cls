VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CraneCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BaseX As Byte
Public BaseY As Byte
Public HeightCrane As Byte
Public LengthCrane As Byte
Public LebarTiang As Byte

'bagian cockpit crane
Public XCock As Integer
Public YCock As Integer
Public XCockpitOnArray As Byte
Public YCockpitOnArray As Byte

'bagian tiang vertikal crane
Public XTiang As Integer
Public YTiang As Integer
Public XTiangOnArray As Byte
Public YTiangOnArray As Byte

'bagian kepala crane
Public XHead As Byte
Public YHead As Byte
Public YKait As Byte
Public LengthChain As Integer
Public CraneSpeed As Integer
Public XPositionHead As Integer
Public YPositionHead As Integer
Public XHeadSmooth As Integer
Public WayMove As Byte

Public MoveDuration As Byte
Public CraneAction As Byte

Public LastPlaceCraneX As Byte
Public LastPlaceCraneY As Byte

Public CranePilih As Boolean
Public Occupied As Boolean
Public PointerDraw As Boolean
Public PointerLong As Byte

Public PickStuff As Byte

Public IsCraneMoving As Boolean
Public XPoint As Byte
Public YPoint As Byte

Public StartMoving As Boolean

Public Property Get byXHead() As Byte
    byXHead = XHead
End Property

Public Property Get byYHead() As Byte
    byYHead = YHead
End Property


Public Sub GlueTiang()
    'rutin ini untuk melakukan pemasangan tiang
    'ke dalam Array Gedung dan merubah nilai boolean Cor menjadi true
    Dim bytRep As Byte
    For bytRep = XTiangOnArray To (XTiangOnArray + LebarTiang) - 1
        ArGedung(bytRep, YTiangOnArray).bCor = True
        ArGedung(bytRep, YTiangOnArray).bGround = True
        ArGedung(bytRep, YTiangOnArray).bWalkable = True
        ArGedung(bytRep, YTiangOnArray).bTiang = True
        ArGedung(bytRep, YTiangOnArray).bSpace = True
        ArGedung(bytRep, YTiangOnArray).bytSprX = 0
        ArGedung(bytRep, YTiangOnArray).bytSprY = 4
    Next bytRep
End Sub

Public Sub LengthenArm()
    'rutin ini untuk memanjangkan lengan crane dengan
    'menggunakan tombol kanan sebagai tombol trigger
    LengthCrane = LengthCrane + 1
    Call DrawMoneyOnScreen(BaseX + LengthCrane, BaseY - HeightCrane - 1, "-10", MOVE_DOWN)
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprY = 1
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bCrane = True
    RedrawMiniMap = True
    Call GameMod.RefreshMiniMap
    
End Sub

Public Sub LowerBase()
    If HeightCrane <= 3 Then Exit Sub
    'rutin ini untuk menaikkan base crane
    'lakukan perubahan pada array gedung
    'ArGedung(BaseX, BaseY - HeightCrane).bytSprX = 2
    'ArGedung(BaseX, BaseY - HeightCrane).bytSprY = 0
    'ArGedung(BaseX, BaseY - HeightCrane).bCrane = False
    
    Call DrawMoneyOnScreen(BaseX - 2, BaseY - HeightCrane, "+5", MOVE_DOWN)
    
    'lakukan perubahan pada crane
    Dim bytPanjang As Byte
    For bytPanjang = 1 To LengthCrane
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane).bytSprX = 2
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane).bytSprY = 1
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane).bCrane = True
    
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bytSprX = 2
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bytSprY = 0
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bCrane = False
    Next bytPanjang
    
    'bagian pemberat
    'Pemberat
    ArGedung(BaseX - 1, BaseY - HeightCrane).bytSprX = 2
    ArGedung(BaseX - 1, BaseY - HeightCrane).bytSprY = 1
    ArGedung(BaseX - 1, BaseY - HeightCrane).bCrane = True
    ArGedung(BaseX - 2, BaseY - HeightCrane).bytSprX = 3
    ArGedung(BaseX - 2, BaseY - HeightCrane).bytSprY = 1
    ArGedung(BaseX - 2, BaseY - HeightCrane).bCrane = True
    
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bCrane = False
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bCrane = False
    
    'bagian cokcpit
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX, BaseY - HeightCrane - 1).bCrane = False
    
    YCockpitOnArray = YCockpitOnArray + 1
    ArGedung(BaseX, YCockpitOnArray).bytSprX = 0
    ArGedung(BaseX, YCockpitOnArray).bytSprY = 1
    ArGedung(BaseX, YCockpitOnArray).bCrane = True
    
    'bagian kepala
    YHead = YHead + 1
    
    HeightCrane = HeightCrane - 1
    RedrawMiniMap = True
    Call GameMod.RefreshMiniMap
End Sub

Public Sub ShortenArm()
    'rutin ini untuk memanjangkan lengan crane dengan
    'menggunakan tombol kanan sebagai tombol trigger
    If (LengthCrane <= 2) Or (XHead >= BaseX + LengthCrane) Then Exit Sub
    
    Call DrawMoneyOnScreen(BaseX + LengthCrane, BaseY - HeightCrane - 1, "+5", MOVE_DOWN)
    
    'bersihkan dengan gambar langit
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bCrane = False
    
    LengthCrane = LengthCrane - 1
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bytSprY = 1
    ArGedung(BaseX + LengthCrane, BaseY - HeightCrane - 1).bCrane = True
    'jika kepala crane berada di ujung , maka samakan dengan panjang tangan crane
    'If XHead > LengthCrane + BaseX Then XHead = LengthCrane + BaseX
    'kurangi nilai lenghtcrane
    RedrawMiniMap = True
    Call GameMod.RefreshMiniMap
    
End Sub

Public Sub RaiseBase()
    If HeightCrane >= (MAXGEDUNGY - 7) Then Exit Sub
    'rutin ini untuk menaikkan base crane
    'lakukan perubahan pada array gedung
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprX = 5
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX, BaseY - HeightCrane - 1).bCrane = True
    Call DrawMoneyOnScreen(BaseX - 2, BaseY - HeightCrane - 1, "-10", MOVE_DOWN)
    
    'lakukan perubahan pada crane
    Dim bytPanjang As Byte
    For bytPanjang = 1 To LengthCrane
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bytSprX = 2
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bytSprY = 0
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 1).bCrane = False
    
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 2).bytSprX = 2
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 2).bytSprY = 1
        ArGedung(BaseX + bytPanjang, BaseY - HeightCrane - 2).bCrane = True
    Next bytPanjang
    
    'bagian pemberat
    'Pemberat
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bCrane = False
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprY = 0
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bCrane = False
    
    ArGedung(BaseX - 1, BaseY - HeightCrane - 2).bytSprX = 2
    ArGedung(BaseX - 1, BaseY - HeightCrane - 2).bytSprY = 1
    ArGedung(BaseX - 1, BaseY - HeightCrane - 2).bCrane = True
    ArGedung(BaseX - 2, BaseY - HeightCrane - 2).bytSprX = 3
    ArGedung(BaseX - 2, BaseY - HeightCrane - 2).bytSprY = 1
    ArGedung(BaseX - 2, BaseY - HeightCrane - 2).bCrane = True
    
    'bagian cokcpit
    'ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprX = 5
    'ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprY = 0
    'ArGedung(BaseX, BaseY - HeightCrane - 1).bCrane = False
    
    ArGedung(BaseX, BaseY - HeightCrane - 2).bytSprX = 0
    ArGedung(BaseX, BaseY - HeightCrane - 2).bytSprY = 1
    ArGedung(BaseX, BaseY - HeightCrane - 2).bCrane = True
    
    YCockpitOnArray = YCockpitOnArray - 1
    'bagian kepala
    YHead = YHead - 1
    
    HeightCrane = HeightCrane + 1
    'LengthChain = LengthChain - 1
    'update layar minimap
    RedrawMiniMap = True
    Call GameMod.RefreshMiniMap
    
    'Call GameMod.UpdateMiniMap
End Sub

Public Sub ShowTiang()
    'rutin ini akan menampilkan gambar tiang sesuai dengan lebar tiang
    
End Sub

Public Function TiangLetak() As Boolean
    'fungsi ini untuk mengetahui apakah tiang
    'bisa diletakkan pada posisi sesuai dengan lebar tiang tersebut
    'Dengan kondisi sbb :
    '1. Di bawah Ytiangonarray telah ada boolean Cor untuk awal dan akhir tiang
    '2. Di samping kiri atau kanan XTiangOnArray sudah ada boolean Cor
    '3. Posisi diletakkan tiang harus bebas dari objek apapun
    '4. Posisi Tiang adalah berbentuk L (ada peletakkan(Y+1) di salah satu sisi dan penempelan di salah satu X-1 or X+1)
    TiangLetak = False
    Dim PosisiAwal As Byte
    Dim PosisiAkhir As Byte
    PosisiAwal = XTiangOnArray
    PosisiAkhir = (XTiangOnArray + LebarTiang) - 1
    If (ArGedung(PosisiAwal, YTiangOnArray + 1).bCor And ArGedung(PosisiAkhir, YTiangOnArray + 1).bCor) _
        Or (ArGedung(PosisiAwal - 1, YTiangOnArray).bCor And ArGedung(PosisiAkhir + 1, YTiangOnArray).bCor) _
        And (ArGedung(PosisiAwal, YTiangOnArray).bSpace And ArGedung(PosisiAkhir, YTiangOnArray).bSpace) _
        Or (ArGedung(PosisiAwal - 1, YTiangOnArray).bCor And ArGedung(PosisiAkhir, YTiangOnArray + 1).bCor) _
        Or (ArGedung(PosisiAkhir + 1, YTiangOnArray).bCor And ArGedung(PosisiAwal, YTiangOnArray + 1).bCor) Then
        
        TiangLetak = True
        
    End If
End Function

Public Sub UpdateCockpit()
    'melakukan update terhadap posisi XPosition dan YPosition
    XCock = ((XCockpitOnArray - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll
    YCock = ((YCockpitOnArray - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll
End Sub

Public Sub UpdateHead()
    'melakukan update terhadap posisi XPosition dan YPosition
    XPositionHead = (((XHead - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll) + XHeadSmooth
    YPositionHead = ((YHead - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll
End Sub

Public Sub DrawHead()
    'rutin ini untuk menggambar kepala crane
    'If (XPositionHead < 800 And XPositionHead > 0) And (YPositionHead > 0 And YPositionHead < 500) Then
    If VisibleHead(byXHead, byYHead) Then
        Dim YPos As Byte
        Call UpdateHead
        BackBuffer.BltFast XPositionHead, YPositionHead, Sprite, BoxRect(56, 20, 69, 39), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        
        'lakukan update terhadap tiang yang akan digunakan
        If Occupied Then
            Call UpdateTiang
            'gambar lebar tiang
            If VisibleTiang Then
                Dim XLebar As Byte
                For XLebar = 1 To LebarTiang
                    BackBuffer.BltFast XTiang + ((XLebar - 1) * TILEWIDTH), YTiang, Sprite, BoxRect(0, 80, 14, 99), DDBLTFAST_WAIT
                Next XLebar
            End If
        End If
        
        'gambar rantai
        For YPos = 1 To LengthChain
            BackBuffer.BltFast XPositionHead, (YPositionHead + (YPos * TILEHEIGHT)), Sprite, BoxRect(70, 20, 84, 39), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
        Next YPos
        'gambar cungkil
        If Not Occupied Then BackBuffer.BltFast XPositionHead, (YPositionHead + ((LengthChain + 1) * TILEHEIGHT)), Sprite, BoxRect(0, 40, 13, 59), DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    End If
    
    
End Sub

Public Sub MoveCrane(WayMove As Byte)
    'rutin untuk menggerakan head crane ke posisi tertentu
    'parameter posisix  - daerah x baru, posisiy - daerah y baru
    'perubahan posisi dilakukan dari daerah lastplacecranex dan y ke posisix,posisiy
    Select Case WayMove
    Case MOVE_LEFT
        MoveDuration = TILEWIDTH \ CraneSpeed
        XHeadSmooth = XHeadSmooth - CraneSpeed
    Case MOVE_RIGHT
        MoveDuration = TILEWIDTH \ CraneSpeed
        XHeadSmooth = XHeadSmooth + CraneSpeed
    End Select
End Sub
Public Property Get PickCockpit() As Boolean
    If Crane.XCock <= CursorX And Crane.XCock + TILEWIDTH >= CursorX And Crane.YCock <= CursorY And Crane.YCock + TILEHEIGHT >= CursorY Then
        PickCockpit = True
    Else
        PickCockpit = False
    End If
End Property

Public Sub PlaceCrane()
    'rutin untuk meletakkan gambar crane ke layar dengan parameter
    'BaseX      - Nilai X di array
    'BaseY      - Nilai Y di array
    
    Dim byY As Byte
    
    ArGedung(BaseX, BaseY).bytSprX = 4
    ArGedung(BaseX, BaseY).bytSprY = 0
    ArGedung(BaseX, BaseY).bCrane = True
    ArGedung(BaseX, BaseY).bSpace = False
    For byY = BaseY - HeightCrane To BaseY - 1
        ArGedung(BaseX, byY).bytSprX = 5
        ArGedung(BaseX, byY).bytSprY = 0
        ArGedung(BaseX, byY).bCrane = True
        ArGedung(BaseX, byY).bSpace = False
        ArGedung(BaseX, byY).bLadder = True
    Next byY
    
    'cockpit Crane
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprX = 0
    ArGedung(BaseX, BaseY - HeightCrane - 1).bytSprY = 1
    'kepala crane cockpit tidak menggunakan penempatan di dalam array untuk mencegah
    'terjadinya salah perhitungan oleh nilai koma dari pembagian
    'maka akan dianggap sama dengan objek worker
    ArGedung(BaseX, BaseY - HeightCrane - 1).bCrane = True
    ArGedung(BaseX, BaseY - HeightCrane - 1).bSpace = False
    ArGedung(BaseX, BaseY - HeightCrane - 1).bCockpit = True
    
    'tangan crane
    For byY = BaseX + 1 To BaseX + LengthCrane
        ArGedung(byY, BaseY - HeightCrane - 1).bytSprX = 2
        ArGedung(byY, BaseY - HeightCrane - 1).bytSprY = 1
        ArGedung(byY, BaseY - HeightCrane - 1).bCrane = True
        ArGedung(byY, BaseY - HeightCrane - 1).bSpace = False
    Next byY
    
    'bagian rantai crane
    'bagian kepala crane
    'ArGedung(XHead, BaseY - LengthChain - 1).bytSprX = 0
    'ArGedung(XHead, BaseY - LengthChain - 1).bytSprY = 2
    'ArGedung(XHead, BaseY - LengthChain - 1).bCrane = True
    
    'Pemberat
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprX = 2
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bytSprY = 1
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bCrane = True
    ArGedung(BaseX - 1, BaseY - HeightCrane - 1).bSpace = False
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprX = 3
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bytSprY = 1
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bCrane = True
    ArGedung(BaseX - 2, BaseY - HeightCrane - 1).bSpace = False
End Sub
Public Sub UpdateTiang()
    'melakukan update terhadap posisi XPosition dan YPosition
    XTiang = ((XTiangOnArray - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll + XHeadSmooth
    YTiang = ((YTiangOnArray - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll
End Sub

Public Property Get Visible() As Boolean
    If XCockpitOnArray >= Gedung.XGedung And XCockpitOnArray <= Gedung.XGedung + MAXTILEX And YCockpitOnArray >= Gedung.YGedung And YCockpitOnArray <= Gedung.YGedung + 31 Then
        Visible = True
    Else
        Visible = False
    End If
End Property

Public Function VisibleHead(XHead As Byte, YHead As Byte) As Boolean
    VisibleHead = False
    If ((XHead + MAXTILEX > Gedung.XGedung) And (XHead > Gedung.XGedung)) And ((YHead < Gedung.YGedung + MAXTILEY) And (YHead > Gedung.YGedung)) Then
        VisibleHead = True
    End If
End Function

Public Property Get VisibleTiang() As Boolean
    If (XTiangOnArray + LebarTiang) >= Gedung.XGedung And (XTiangOnArray + LebarTiang) <= Gedung.XGedung + MAXTILEX And YTiangOnArray >= Gedung.YGedung And YTiangOnArray <= Gedung.YGedung + 31 Then
        VisibleTiang = True
    Else
        VisibleTiang = False
    End If
End Property

Public Property Get XCockpit() As Byte
    XCockpit = BaseX
End Property

Public Property Get YCockpit() As Byte
    YCockpit = BaseY - (HeightCrane + 1)
End Property

Public Function MoveRight(XHead As Byte) As Boolean
    MoveRight = False
    If XHead < (LengthCrane + BaseX) Then
        MoveRight = True
    End If
End Function

Public Function MoveLeft(XHead As Byte) As Boolean
    MoveLeft = False
    If XHead > (BaseX + 1) Then
        MoveLeft = True
    End If
End Function

Public Sub ContinueMovingX()
    'rutin untuk menjalankan crane ke posisi selanjutnya
    If MoveDuration >= 1 Then
        MoveDuration = MoveDuration - 1
        Select Case WayMove
        Case MOVE_RIGHT
            XHeadSmooth = XHeadSmooth + CraneSpeed
        Case MOVE_LEFT
            XHeadSmooth = XHeadSmooth - CraneSpeed
        End Select
    Else
        IsCraneMoving = False
        XHeadSmooth = 0
        Select Case WayMove
        Case MOVE_RIGHT
            XHead = XHead + 1
        Case MOVE_LEFT
            XHead = XHead - 1
        End Select
    End If
End Sub

Public Function IsArrivedX() As Boolean
    'fungsi untuk melacak peletakan objek crane
    IsArrivedX = False
    If (XHead = XPoint) Then
        IsArrivedX = True
    End If
End Function

Public Function IsArrivedY() As Boolean
    'fungsi untuk melacak peletakan objek crane
    IsArrivedY = False
    If (YHead = YPoint) Then
        IsArrivedY = True
    End If
End Function


