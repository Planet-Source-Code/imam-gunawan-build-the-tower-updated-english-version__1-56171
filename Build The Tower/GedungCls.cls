VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GedungCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public XGedung As Integer
Public YGedung As Integer

Sub DrawHBeam()
    ArGedung(15, 146).bytSprX = 0
    ArGedung(15, 146).bytSprY = 3
    ArGedung(16, 146).bytSprX = 1
    ArGedung(16, 146).bytSprY = 3
    ArGedung(17, 146).bytSprX = 1
    ArGedung(17, 146).bytSprY = 3
    ArGedung(18, 146).bytSprX = 1
    ArGedung(18, 146).bytSprY = 3
    ArGedung(19, 146).bytSprX = 1
    ArGedung(19, 146).bytSprY = 3
    ArGedung(20, 146).bytSprX = 1
    ArGedung(20, 146).bytSprY = 3
    ArGedung(21, 146).bytSprX = 2
    ArGedung(21, 146).bytSprY = 3
End Sub

Public Function GedungValue(intX As Byte, intY As Byte) As RECT
    GedungValue.Left = intX * TILEWIDTH
    GedungValue.Top = intY * TILEHEIGHT
    GedungValue.Right = (intX * TILEWIDTH) + TILEWIDTH
    GedungValue.Bottom = (intY * TILEHEIGHT) + TILEHEIGHT
End Function

Sub InitBoundaries()
    'rutin untuk melakukan seal terhadap ruang bangun gedung
    'sesuai dengan objektif
    Dim bytX As Byte
    With Objektif
        For bytX = 1 To (.bytStart - 1)
            ArGedung(bytX, 147).bGround = False
        Next bytX
        For bytX = (.bytEnd + 1) To 150
            ArGedung(bytX, 147).bGround = False
        Next bytX
    End With
End Sub

Public Sub InitGedung()
    'untuk inisialisasi variabel gedung awal baru
    'isikan semua data ke dalam Array Gedung
    XGedung = 30
    YGedung = 122
    
    Screen.XScroll = 0
    Screen.YScroll = 0
    
    Dim intX As Integer
    Dim intY As Integer
    For intX = 0 To MAXGEDUNGX
        For intY = 0 To MAXGEDUNGY
            If intY = MAXGEDUNGY - 2 Or intY = MAXGEDUNGY - 3 Then
                ArGedung(intX, intY).bytSprX = 3
                ArGedung(intX, intY).bytSprY = 0
                ArGedung(intX, intY).bGround = True
                ArGedung(intX, intY).bWalkable = True
            Else
                ArGedung(intX, intY).bytSprX = 2
                ArGedung(intX, intY).bytSprY = 0
                ArGedung(intX, intY).bSpace = True
            End If
        Next intY
    Next intX
    
    'bikin tampilan H-Beam
    Call DrawHBeam
    
    'bikin workshop
    Call DrawWorkShop
    
    'Call MakeAllArrayFalse
    
    Call InitBoundaries
    
    'lakukan init terhadap pembangunan crane
    Crane.BaseX = 11
    Crane.BaseY = 146
    'tinggi crane
    Crane.HeightCrane = 5
    'panjang tangan crane
    Crane.LengthCrane = 30
    
    'kepala head crane
    Crane.XHead = Crane.BaseX + 2
    Crane.YHead = Crane.BaseY - Crane.HeightCrane - 1
    Crane.LengthChain = 2
    Crane.CraneSpeed = 2
    Crane.XPositionHead = ((Crane.XHead - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll
    Crane.YPositionHead = ((Crane.YHead - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll
    Crane.WayMove = MOVE_RIGHT
    
    Crane.XCockpitOnArray = Crane.BaseX
    Crane.YCockpitOnArray = Crane.BaseY - (Crane.HeightCrane + 1)
    Crane.LengthChain = 2
    Crane.PlaceCrane
End Sub

Sub MakeAllArrayFalse()
    'rutin untuk menjadikan semua isi array bernilai False
    'karena nilai walkable harus bernilai True untuk membuat worker berjalan
    Dim bytX As Byte
    Dim bytY As Byte
    For bytX = 1 To MAXGEDUNGX
        For bytY = 1 To MAXGEDUNGY
            ArGedung(bytX, bytY).bWalkable = False
        Next bytY
    Next bytX
End Sub

Public Sub DrawWorkShop()
    Dim bytX As Byte
    Dim bytY As Byte
    For bytX = 1 To 4
    For bytY = 0 To 3
        ArGedung(146 + bytX, 146 - bytY).bytSprX = 1
        ArGedung(146 + bytX, 146 - bytY).bytSprY = 2
        ArGedung(146 + bytX, 146 - bytY).bWall = True
    Next bytY
    Next bytX
    ArGedung(147, 142).bytSprX = 2
    ArGedung(147, 142).bytSprY = 2
    ArGedung(148, 142).bytSprX = 3
    ArGedung(148, 142).bytSprY = 2
End Sub
Sub SwapArr2Ged()
    'rutin untuk melakukan swap array ke gedung
    Dim bytX As Byte
    Dim bytY As Byte
    For bytX = 1 To MAXGEDUNGX
        For bytY = 1 To MAXGEDUNGY
            'sebelumnya pindahkan semua data bytSprx dan bytSpry ke temporary array
            TmpSpr(bytX, bytY).bytSprX = ArGedung(bytX, bytY).bytSprX
            TmpSpr(bytX, bytY).bytSprY = ArGedung(bytX, bytY).bytSprY
        Next bytY
    Next bytX
End Sub

Public Sub ViewGedung(Tampak As Byte)
    Dim bytX As Byte
    Dim bytY As Byte
    'rutin untuk merubah view bangunan dari sudut
    Select Case Tampak
    Case TAMPAK_LUAR
        For bytX = 1 To MAXGEDUNGX
            For bytY = 1 To MAXGEDUNGY
                'sebelumnya pindahkan semua data bytSprx dan bytSpry ke temporary array
                TmpSpr(bytX, bytY).bytSprX = ArGedung(bytX, bytY).bytSprX
                TmpSpr(bytX, bytY).bytSprY = ArGedung(bytX, bytY).bytSprY
                
                If ArGedung(bytX, bytY).bLadder And ArGedung(bytX, bytY).bCat Then
                    ArGedung(bytX, bytY).bytSprX = 1
                    ArGedung(bytX, bytY).bytSprY = 0
                ElseIf ArGedung(bytX, bytY).bLadder And Not ArGedung(bytX, bytY).bCat Then
                    ArGedung(bytX, bytY).bytSprX = 0
                    ArGedung(bytX, bytY).bytSprY = 0
                End If
            
                'lakukan perubahan pada saat itu
                If ArGedung(bytX, bytY).bCor Or _
                (ArGedung(bytX, bytY).bCat And ArGedung(bytX, bytY).bWall) Then
                    ArGedung(bytX, bytY).bytSprX = 1
                    ArGedung(bytX, bytY).bytSprY = 0
                End If
                
            Next bytY
        Next bytX
        
    Case TAMPAK_DALAM
        For bytX = 1 To MAXGEDUNGX
            For bytY = 1 To MAXGEDUNGY
                'pindahkan semua data bytSprx dan bytSpry dari temporary array
                ArGedung(bytX, bytY).bytSprX = TmpSpr(bytX, bytY).bytSprX
                ArGedung(bytX, bytY).bytSprY = TmpSpr(bytX, bytY).bytSprY
            Next bytY
        Next bytX
    End Select
End Sub




Sub Walk()
    Dim X As Byte
    For X = 50 To 81
        ArGedung(X, 143).bWalkable = True
    Next X
    For X = 50 To 58
        ArGedung(X, 139).bWalkable = True
    Next X
    For X = 95 To 101
        ArGedung(X, 143).bWalkable = True
    Next X
    For X = 115 To 119
        ArGedung(X, 143).bWalkable = True
    Next X
    For X = 0 To 150
        ArGedung(X, 147).bWalkable = True
        ArGedung(X, 148).bWalkable = True
    Next X
    
End Sub


