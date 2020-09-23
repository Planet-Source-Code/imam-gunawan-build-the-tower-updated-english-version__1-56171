VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScreenCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public XScroll As Integer
Public YScroll As Integer
Public Redraw As Boolean
Public Sub CheckScroll()
    'digunakan untuk melacak scrolling terhadap pergerakan mouse
    If ScrollScreen Then
        Select Case ScrollWay
        Case SCROLL_LEFT
            Call ScrollLeft
        Case SCROLL_RIGHT
            Call ScrollRight
        Case SCROLL_UP
            Call ScrollUp
        Case SCROLL_DOWN
            Call ScrollDown
        End Select
    End If
End Sub



Public Sub DrawGedung(GedungX As Byte, GedungY As Byte, XScroll As Integer, YScroll As Integer)
    Dim bytX As Byte
    Dim bytY As Byte
    
    Dim DecX As Byte
    Dim DecY As Byte
    
    If Redraw Then
    
    DecX = GedungX - 1
    DecY = GedungY - 1
    
    Dim Y As Byte
    
    'menggambar portion bitmap gedung ke layar
    'Layar.BltColorFill BoxRect(0, 0, 800, 500), QBColor(0)
    
    For bytX = GedungX - 1 To (GedungX + MAXTILEX)
        For bytY = GedungY - 1 To (GedungY + MAXTILEY)
            Layar.BltFast ((bytX - DecX) * TILEWIDTH), ((bytY - DecY) * TILEHEIGHT), Sprite, Gedung.GedungValue(ArGedung(bytX, bytY).bytSprX, ArGedung(bytX, bytY).bytSprY), DDBLTFAST_WAIT
        Next bytY
    Next bytX
    
    Redraw = False
    End If
    
    BackBuffer.BltFast 0, 0, Layar, BoxRect(TILEWIDTH - XScroll, TILEHEIGHT - YScroll, (800 + TILEWIDTH) - XScroll, (560 - TILEHEIGHT) - YScroll), DDBLTFAST_WAIT
    
End Sub

Public Property Get GetTileX() As Byte
    'fungsi untuk mengetahui letak tilex
    GetTileX = (Int(CursorX / TILEWIDTH) + Gedung.XGedung) - IIf(Gedung.XGedung = 1, 1, 0)
End Property

Public Property Get GetTileY() As Byte
    'fungsi untuk mengetahui letak tilex
    GetTileY = Int(CursorY / TILEHEIGHT) + Gedung.YGedung - IIf(Gedung.YGedung = MAXGEDUNGY - 28, 0, 1)
End Property

Public Sub ScrollLeft()
    If Gedung.XGedung > 1 Then
        If XScroll > 0 Then
            If XScroll < TILEWIDTH Then XScroll = XScroll + SCROLLSPEED_LEFT_RIGHT
            If XScroll >= TILEWIDTH Then
                Gedung.XGedung = Gedung.XGedung - 1
                XScroll = SCROLLSPEED_LEFT_RIGHT
            End If
        Else    'jika pertama kali digulung
            XScroll = XScroll + SCROLLSPEED_LEFT_RIGHT
        End If
    End If
End Sub

Public Sub ScrollUp()
    If Gedung.YGedung > 1 Then
        If YScroll > 0 Then
            If YScroll < TILEHEIGHT Then YScroll = YScroll + SCROLLSPEED_UP_DOWN
            If YScroll >= TILEHEIGHT Then
                Gedung.YGedung = Gedung.YGedung - 1
                YScroll = SCROLLSPEED_UP_DOWN
            End If
        Else    'jika pertama kali digulung
            YScroll = YScroll + SCROLLSPEED_UP_DOWN
        End If
    End If
End Sub

Public Sub ScrollDown()
    If Gedung.YGedung < (MAXGEDUNGY - MAXTILEY) Then
        If YScroll < 0 Then
            If Abs(YScroll) < TILEHEIGHT Then YScroll = YScroll - SCROLLSPEED_UP_DOWN
            If Abs(YScroll) >= TILEHEIGHT Then
                Gedung.YGedung = Gedung.YGedung + 1
                YScroll = SCROLLSPEED_UP_DOWN * -1
            End If
        Else    'jika pertama kali digulung
            YScroll = YScroll + (SCROLLSPEED_UP_DOWN * -1)
        End If
    End If
End Sub

Public Sub ScrollRight()
    If Gedung.XGedung < (MAXGEDUNGX - MAXTILEX) Then
        If XScroll < 0 Then
            If Abs(XScroll) < TILEWIDTH Then XScroll = XScroll - SCROLLSPEED_LEFT_RIGHT
            If Abs(XScroll) >= TILEWIDTH Then
                Gedung.XGedung = Gedung.XGedung + 1
                XScroll = SCROLLSPEED_LEFT_RIGHT * -1
            End If
        Else    'jika pertama kali digulung
            XScroll = XScroll + SCROLLSPEED_LEFT_RIGHT * -1
        End If
    End If
End Sub



