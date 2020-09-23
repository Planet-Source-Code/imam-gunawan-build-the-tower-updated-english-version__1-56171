VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CuacaCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Suhu As Byte             'suhu cuaca
Public Status As Byte           'status cuaca
Public YMove As Byte
Public Panjang As Byte
Public JamMulai As Byte
Public MenitMulai As Byte
Public JamSelesai As Byte
Public MenitSelesai As Byte
Public JamSekarang As Byte
Public MenitSekarang As Byte
Public HariIni As Date          'hari ini dalam pekerjaan
Public Delayment As Byte        'bekerja dalam satu delay setiap max_fps
Public EachFPS As Integer       'nilai perubah untuk max_fps

Sub Initialize()
    Cuaca.JamMulai = (JAM_MULAI)
    Cuaca.MenitMulai = (MENIT_MULAI)
    Cuaca.JamSelesai = (JAM_SELESAI)
    Cuaca.MenitSelesai = (MENIT_SELESAI)
    Cuaca.JamSekarang = (JAM_MULAI)
    Cuaca.MenitSekarang = (MENIT_MULAI)
    Cuaca.HariIni = #8/14/2000#
End Sub

