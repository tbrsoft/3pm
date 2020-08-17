Attribute VB_Name = "libreria"
Public Lista_reproduccion(1000) As String
Public guardar, locaerr, tamanio, indice, contador, pasa, k, X, click As Integer
Public drive, path, pattern As String
Public disc, directorio, archivos As String
Public despla As Boolean
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long
#If Win32 Then
    Declare Sub MensajeBip Lib "user32" (ByVal N As Long)
#Else
    Declare Sub MensajeBip Lib "User" (ByVal N As Integer)
#End If

' Declaración de una rutina de Windows. Esta instrucción es para el módulo.
'Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
'Declare Function waveOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long
'Declare Function waveOutSetPlaybackRate Lib "winmm.dll" (ByVal hWaveOut As Long, ByVal dwRate As Long) As Long
'Declare Function waveOutOpen Lib "winmm.dll" (lphWaveOut As Long, ByVal uDeviceID As Long, lpFormat As WAVEFORMAT, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
'public sub Function WaitMessage Lib "user32" () As Long
Declare Function WaitMessage Lib "user32" () As Long
'Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Color As Long) As Long





Global s(22), es(22)
Global im As Integer

Public Sub errores()
locaerr = 1
 
   Select Case Errs
         Case 481
           MsgBox prompt:="La imagen no está Disponible. ", _
            Buttons:=vbExclamation
         
         Case 72
            MsgBox prompt:="Archivo no está Disponible. ", _
            Buttons:=vbExclamation
         Case 381
            MsgBox prompt:=" final de Archivo...", _
            Buttons:=vbExclamation
           Exit Sub
          Case 68
            MsgBox prompt:="La unidad no está preparada. Inserte un disco en la unidad.", _
            Buttons:=vbExclamation
          Case 76
            MsgBox prompt:="No se encuentra el Directorio de las Imagenes.", _
            Buttons:=vbExclamation
          Exit Sub
          Case 424
             MsgBox prompt:="No se encuentra el Directorio de las Imagenes.", _
            Buttons:=vbExclamation
          Exit Sub
     End Select
   

End Sub
Public Sub Centrarimagen(Exterior As PictureBox, Interior As Image)

Interior.Move (Exterior.Width - Interior.Width) / 2, (Exterior.Height - Interior.Height) / 2

End Sub
Public Sub CentrarControlEnControl(Exterior As Control, Interior As Control)

Interior.Move (Exterior.Width - Interior.Width) / 2, (Exterior.Height - Interior.Height) / 2

End Sub

Public Sub CentrarControl(Exterior As Form, Interior As Control)

Interior.Move (Exterior.ScaleWidth - Interior.Width) / 2, (Exterior.ScaleHeight - Interior.Height) / 2

End Sub


Public Sub Centrar(Objeto As Object)

Objeto.Top = (Screen.Height - Objeto.Height) / 2
Objeto.Left = (Screen.Width - Objeto.Width) / 2

End Sub
'' ***********************************

 Public Sub Centrar2(Objeto As Object, hform As Form)
Objeto.Top = (hform.ScaleHeight - (Objeto.Height - 3200)) / 2
Objeto.Left = (hform.ScaleWidth - Objeto.Width) / 2

 End Sub
Public Sub fondocolor(hform As Form, r As Integer, g As Integer, b As Integer)
Screen.MousePointer = 11

With hform
.AutoRedraw = True
.DrawStyle = 6
.DrawMode = 13
If .ScaleHeight / 255 > 1 Then
    .DrawWidth = .ScaleHeight / 255
Else
    .DrawWidth = 2  'Previene malos resultados en el aspecto
End If
.ScaleMode = 3
End With

Y = 0

If r And Not g And Not b Then

  For i = 0 To hform.ScaleHeight
    hform.Line (0, Y)-(hform.Width, Y + (hform.ScaleHeight / 255)), RGB(i, g, b), BF
    Y = Y + (hform.ScaleHeight / 255)
  Next i

ElseIf g And Not b And Not r Then

  For i = 0 To hform.ScaleHeight
    hform.Line (0, Y)-(hform.Width, Y + (hform.ScaleHeight / 255)), RGB(r, i, b), BF
    Y = Y + (hform.ScaleHeight / 255)
  Next i

ElseIf b And Not g And Not r Then

  For i = 0 To hform.ScaleHeight
    hform.Line (0, Y)-(hform.Width, Y + (hform.ScaleHeight / 255)), RGB(r, g, i), BF
    Y = Y + (hform.ScaleHeight / 255)
 Next i

End If

Screen.MousePointer = 0

End Sub

Public Sub pantalla(tform As Form)
With tform
.Top = 0
.Left = 0
.Height = Screen.Height
.Width = Screen.Width
End With

End Sub

'--------------------------------------------------------------------------------
Public Sub limpiar(kform As Form, r As Integer, g As Integer, b As Integer)
Screen.MousePointer = 11

With kform
.AutoRedraw = True
.DrawStyle = 6
.DrawMode = 13
If .ScaleHeight / 255 > 1.5 Then
    .DrawWidth = .ScaleHeight / 255
Else
    .DrawWidth = 2  'Previene malos resultados en el aspecto
End If
.ScaleMode = 3
End With

Y = 0

If b And Not g And Not r Then

  For i = 0 To kform.ScaleHeight
    
    kform.Line (0, Y)-(kform.Width, Y + (kform.ScaleHeight / 255)), RGB(r, g, i), BF
    Y = Y + (kform.ScaleHeight / 255)
 Next i

End If

Screen.MousePointer = 0

End Sub
'Public Sub main()
'inicio.Show 0
'
'End Sub
Public Sub fondo(frm As Form, elcolor As Integer)
frm.DrawStyle = vbInsideSolid
frm.DrawWidth = 2
frm.ScaleMode = vbPixels
frm.ScaleHeight = 256
For i = 0 To 255
  Select Case elcolor
  Case 0
    frm.Line (0, i)-(frm.Width, i + 1), RGB(255 - i, 0, 0), B
  Case 1
    frm.Line (0, i)-(frm.Width, i + 1), RGB(0, 255 - i, 0), B
  Case 2
  frm.Line (0, i)-(frm.Width, i + 1), RGB(0, 0, 255 - i), B
 Case Else
  MsgBox "error interno al seleccionar colores"
  Exit For
  End Select
Next i

End Sub

Public Sub fondopicture(rpic As PictureBox, r As Integer, g As Integer, b As Integer)
Screen.MousePointer = 11

With rpic
.AutoRedraw = True
.DrawStyle = 6
.DrawMode = 13
If .ScaleHeight / 255 > 1 Then
    .DrawWidth = .ScaleHeight / 255
Else
    .DrawWidth = 2  'Previene malos resultados en el aspecto
End If
.ScaleMode = 3
End With

Y = 0

If r And Not g And Not b Then

  For i = 0 To rpic.ScaleHeight
    rpic.Line (0, Y)-(rpic.Width, Y + (rpic.ScaleHeight / 255)), RGB(i, g, b), BF
    Y = Y + (rpic.ScaleHeight / 255)
  Next i

ElseIf g And Not b And Not r Then

  For i = 0 To rpic.ScaleHeight
    rpic.Line (0, Y)-(rpic.Width, Y + (rpic.ScaleHeight / 255)), RGB(r, i, b), BF
    Y = Y + (rpic.ScaleHeight / 255)
  Next i

ElseIf b And Not g And Not r Then

  For i = 0 To rpic.ScaleHeight
    rpic.Line (0, Y)-(rpic.Width, Y + (rpic.ScaleHeight / 255)), RGB(r, g, i), BF
    Y = Y + (rpic.ScaleHeight / 255)
 Next i

End If

Screen.MousePointer = 0

End Sub
Public Sub fondo2(frm As Form, r As Integer, v As Integer, a As Integer)
'no funciona
frm.DrawStyle = vbInsideSolid
frm.DrawWidth = 2
frm.ScaleMode = vbPixels
frm.ScaleHeight = 256
For i = 0 To 255
    frm.Line (0, i)-(frm.Width, i + 1), RGB(r - i * r, v - i * v, a - i * a), B
Next i

End Sub
