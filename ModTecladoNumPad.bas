Attribute VB_Name = "ModTecladoNumPad"
Public Declare Function PeekMessage Lib "user32" Alias _
  "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, _
  ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, _
  ByVal wRemoveMsg As Long) As Long

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public Type MSG
   hwnd As Long
   message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type

Public Const PM_NOREMOVE = &H0
Public Const WM_KEYDOWN = &H100 '256
Public Const WM_KEYUP = &H101 '257
Public Const VK_RETURN = &HD

Public Function IsKeyPad(FRM As Form) As Boolean
    
    Dim MyMsg As MSG, RetVal As Long
    IsKeyPad = False
    
    'pass:
    '  MSG structure to receive message information
    '  my window handle
    '  low and high filter of 0, 0 to trap all messages
    '  PM_NOREMOVE to leave the keystroke in the message queue
    '  use PM_REMOVE (1) to remove it
    RetVal = PeekMessage(MyMsg, FRM.hwnd, 0, 0, PM_NOREMOVE)
    
    ' now, per Q77550, you should look for a MSG.wParam of VK_RETURN
    ' if this was the keystroke, then test bit 24 of the lparam - if ON,
    ' then keypad was used, otherwise, keyboard was used
    If RetVal <> 0 Then 'si es cero es que no hay mensaje!
        
       'MyMsg.message es el valor del evento en caso _
            del teclado (?) WM_KEYUP o WM_KEYDOWN
       'me da siempre 258!!!!????
       'MyMsg.wParam = VK_RETURN es la letra apretada
          If MyMsg.lParam And &H1000000 Then
             IsKeyPad = True
          Else
            IsKeyPad = False
          End If
        'End If
    Else
        'No hay mensaje disponible.
    End If
End Function
