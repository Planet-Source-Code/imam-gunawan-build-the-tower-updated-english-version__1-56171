VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyboardCls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX As New DirectX7
Private Di As DirectInput
Private diKeyB As DirectInputDevice
Private diKeyState As DIKEYBOARDSTATE

Public aHwnd As Long

Public Function CloseDI() As Boolean

    'unload all the keyboard device
    diKeyB.Unacquire
    Set diKeyB = Nothing
    Set Di = Nothing
    Set DX = Nothing
    
End Function

Public Sub ClearDI()
    diKeyB.Unacquire
End Sub

Public Function InitDI(Optional HaveMouse As Boolean = False) As Boolean
    On Error GoTo Keluar
    
    'create the input device
    Set Di = DX.DirectInputCreate()
    'create the keyboard device
    Set diKeyB = Di.CreateDevice("GUID_SysKeyboard")
    'set the format keyboard
    diKeyB.SetCommonDataFormat DIFORMAT_KEYBOARD
    
    'set the cooperative for which handle the device
    diKeyB.SetCooperativeLevel aHwnd, DISCL_BACKGROUND Or DISCL_EXCLUSIVE
    InitDI = True
    
    Exit Function
    
Keluar:
    InitDI = False
End Function


Public Function KeyPressed(lKey As Long) As Boolean
    KeyPressed = (diKeyState.Key(lKey) <> 0)
End Function

Public Function KeyStatus(lKey As Long) As Boolean
    If (diKeyState.Key(lKey) And &H80) <> 0 Then
        KeyStatus = True
    Else
        KeyStatus = False
    End If
End Function

Public Function SuccessAcquire() As Boolean
    diKeyB.Acquire
    diKeyB.GetDeviceStateKeyboard diKeyState
    If (Err.Number = 0) Then
        SuccessAcquire = True
    Else
        SuccessAcquire = False
    End If
End Function


