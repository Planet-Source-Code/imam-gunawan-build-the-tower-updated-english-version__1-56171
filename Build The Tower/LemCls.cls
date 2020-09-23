VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LemCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Kelas Worker
'Jati diri worker
Public Active As Boolean            'kondisi (non)aktif worker
Public NoLemm As Integer            'No worker ID
Public Nama As String               'Nama Worker
Public Honor As Integer             'Honor worker
Public Stability As Byte            'stabilitas worker
Public Tolerance As Byte            'toleransi terhadap pekerjaan
Public Progress As Byte             'progress kemajuan pekerjaan

Public Frame As Byte                'posisi frame worker
Public Way As Byte                  '
Public Job As Byte
Public JobX As Byte
Public JobY As Byte
Public Status As Byte               'status worker (walk,work,idle,rest,etc.)
Public XPosition As Integer         'posisi xposisi worker di layar
Public YPosition As Integer         'posisi yposisi worker di layar
Public Perubah As Integer
Public XTileOnArray As Byte         'posisi worker di array
Public YTileOnArray As Byte         'posisi worker di array
Public WorkSpeed As Byte            'nilai kecepatan perubahan antar frame dalam bekerja
Public WorkNow As Byte              'nilai kecepatan pekerja saat sekarang
Public CurrentSpeed As Byte         'kecepatan perubahan frame worker

Public WalkSpeed As Byte            'kecepatan jalan worker
Public IsMoving As Boolean          'nilai boolean menyatakan kondisi worker
Public WayMove As Byte              'Arah jalan worker
Public SearchMove As Byte           'Arah yang dicari worker
Public MoveDuration As Byte
Public XHeadSmooth As Integer       'nilai penambah untuk posisi x
Public YHeadSmooth As Integer       'nilai penambah untuk posisi y

Public XPoint As Byte               'nilai X yang diclick
Public YPoint As Byte               'nilai Y yang diclick
Public Sub ContinueMoving()
'rutin untuk menjalankan crane ke posisi selanjutnya
If MoveDuration >= 1 Then
    MoveDuration = MoveDuration - 1
    Select Case WayMove
    Case MOVE_RIGHT
        XHeadSmooth = XHeadSmooth + WalkSpeed
    Case MOVE_LEFT
        XHeadSmooth = XHeadSmooth - WalkSpeed
    Case MOVE_UP
        YHeadSmooth = YHeadSmooth - WalkSpeed
    Case MOVE_DOWN
        YHeadSmooth = YHeadSmooth + WalkSpeed
    End Select
Else
    IsMoving = False
    XHeadSmooth = 0
    YHeadSmooth = 0
    Select Case WayMove
    Case MOVE_RIGHT
        XTileOnArray = XTileOnArray + 1
    Case MOVE_LEFT
        XTileOnArray = XTileOnArray - 1
    Case MOVE_UP
        YTileOnArray = YTileOnArray - 1
    Case MOVE_DOWN
        YTileOnArray = YTileOnArray + 1
    End Select
    
    If SearchMove <> MOVE_NONE Then
        Select Case SearchMove
        Case MOVE_UP
            If ArGedung(XTileOnArray, YTileOnArray - 1).bLadder Or ArGedung(XTileOnArray, YTileOnArray - 2).bGround Then
                Status = CLIMB
                SearchMove = WayMove
                WayMove = MOVE_UP
            End If
        Case MOVE_DOWN
            If ArGedung(XTileOnArray, YTileOnArray + 1).bLadder Then
                Status = CLIMB
                SearchMove = WayMove
                WayMove = MOVE_DOWN
            End If
        Case MOVE_LEFT
            If ArGedung(XTileOnArray - 1, YTileOnArray + 1).bGround Then
                Status = WALK
                SearchMove = WayMove
                WayMove = MOVE_LEFT
            End If
        Case MOVE_RIGHT
            If ArGedung(XTileOnArray + 1, YTileOnArray + 1).bGround Then
                Status = WALK
                SearchMove = WayMove
                WayMove = MOVE_RIGHT
            End If
        End Select
    End If
    
    'cek apakah daerah yang dipijak masih bisa dipijak
    Select Case WayMove
    Case MOVE_LEFT
        If Not MoveLeft(XTileOnArray, YTileOnArray) Then
            'SearchMove = MOVE_NONE
            Status = IDLE
            XPoint = XTileOnArray
            'YPoint = YTileOnArray
        End If
    Case MOVE_RIGHT
        If Not MoveRight(XTileOnArray, YTileOnArray) Then
            'SearchMove = MOVE_NONE
            Status = IDLE
            XPoint = XTileOnArray
            'YPoint = YTileOnArray
        End If
    Case MOVE_UP
        If Not MoveUp(XTileOnArray, YTileOnArray) Then
            'SearchMove = MOVE_NONE
            Status = IDLE
            XPoint = XTileOnArray
            'YPoint = YTileOnArray
        End If
    Case MOVE_DOWN
        If Not MoveDown(XTileOnArray, YTileOnArray) Then
            'SearchMove = MOVE_NONE
            Status = IDLE
            XPoint = XTileOnArray
            'YPoint = YTileOnArray
        End If
    End Select
End If
End Sub
Property Get IsArrived() As Boolean
    'fungsi untuk melacak peletakan objek crane
    IsArrived = False
    If (XTileOnArray = XPoint) Then 'And (YTileOnArray = YPoint) Then
        IsArrived = True
        Status = IDLE
    End If
End Property

Public Function MoveLeft(XTile As Byte, YTile As Byte) As Boolean
    'bergerak ke kiri mempunyai 3 kondisi agar pekerja bisa berjalan
    '1. Daerah sebelah kiri mempunyai nilai walkable
    '2. Masih lebih besar dari nilai 1
    '3. Daerah sebelah kiri bisa dipijak atau mempunyai tanah
    MoveLeft = False
    If (XTile > 1) Then
    If (ArGedung(XTile - 1, YTile + 1).bWalkable = True) Then
        MoveLeft = True
    End If
    End If
End Function

Public Function MoveRight(XTile As Byte, YTile As Byte) As Boolean
    'bergerak ke kanan mempunyai 3 kondisi agar pekerja bisa berjalan
    '1. Daerah sebelah kanan mempunyai nilai walkable
    '2. Masih lebih kecil dari nilai maxgedungx
    '3. Daerah sebelah kanan bisa dipijak atau mempunyai tanah
    MoveRight = False
    If (XTile + 1 < MAXGEDUNGX) Then
    If (ArGedung(XTile + 1, YTile + 1).bWalkable = True) Then
        MoveRight = True
    End If
    End If
End Function

Public Function MoveUp(XTile As Byte, YTile As Byte) As Boolean
    MoveUp = False
    If ArGedung(XTile, YTile - 1).bLadder Then
        MoveUp = True
    Else
        If SearchMove = MOVE_LEFT Then
            If ArGedung(XTile - 1, YTile).bGround Then MoveUp = True
        ElseIf SearchMove = MOVE_RIGHT Then
            If ArGedung(XTile + 1, YTile).bGround Then MoveUp = True
        End If
    End If
End Function

Public Function MoveDown(XTile As Byte, YTile As Byte) As Boolean
    MoveDown = False
    If ArGedung(XTile, YTile + 1).bLadder = True Then
        MoveDown = True
    End If
End Function

Public Sub MoveWorker()
        Select Case WayMove
        Case MOVE_LEFT
            'If MoveLeft(XTileOnArray, YTileOnArray) Then
                MoveDuration = (TILEWIDTH \ WalkSpeed) - 1
                XHeadSmooth = XHeadSmooth - WalkSpeed
            'End If
        Case MOVE_RIGHT
            'If MoveRight(XTileOnArray, YTileOnArray) Then
                MoveDuration = (TILEWIDTH \ WalkSpeed) - 1
                XHeadSmooth = XHeadSmooth + WalkSpeed
            'End If
        Case MOVE_UP
            'If MoveUp(XTileOnArray, YTileOnArray) Then
                MoveDuration = (TILEHEIGHT \ WalkSpeed) - 1
                YHeadSmooth = YHeadSmooth - WalkSpeed
            'End If
        Case MOVE_DOWN
            'If MoveDown(XTileOnArray, YTileOnArray) Then
                MoveDuration = (TILEHEIGHT \ WalkSpeed) - 1
                YHeadSmooth = YHeadSmooth + WalkSpeed
            'End If
        End Select
End Sub

Public Sub UpdatePosition()
    'melakukan update terhadap posisi XPosition dan YPosition
    XPosition = ((XTileOnArray - Gedung.XGedung) * TILEWIDTH) + Screen.XScroll
    YPosition = ((YTileOnArray - Gedung.YGedung) * TILEHEIGHT) + Screen.YScroll
End Sub

Public Sub UpdateFrame()
    Dim MoveSpeed As Byte
    Select Case Status
    Case WALK
        MoveSpeed = WalkSpeed
    Case WORK
        MoveSpeed = WorkSpeed
    Case Else
        MoveSpeed = WalkSpeed
    End Select
    
    If CurrentSpeed >= MoveSpeed Then
        CurrentSpeed = 0
        If Frame < 7 Then
            Frame = Frame + 1
        Else
            Frame = 0
        End If
    Else
        CurrentSpeed = CurrentSpeed + 1
    End If
End Sub

Public Property Get Visible() As Boolean
    If XTileOnArray >= Gedung.XGedung And XTileOnArray <= Gedung.XGedung + 57 And YTileOnArray >= Gedung.YGedung And YTileOnArray <= Gedung.YGedung + 28 Then
        Visible = True
    Else
        Visible = False
    End If
End Property

Public Property Get XTile() As Byte
    XTile = XPosition \ TILEWIDTH
End Property

Public Property Get YTile() As Byte
    YTile = YPosition \ TILEHEIGHT
End Property

