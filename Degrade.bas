Attribute VB_Name = "mdDegrade"
Public Sub Degrade(QC As Object, Optional Zoom As Long = 0, _
    Optional Horizontal As Boolean = False)
    
Dim H As Long
Dim W As Long

Select Case Zoom
    Case 0
        H = QC.Height / 15
        W = QC.Width / 15
    Case 1
        H = QC.Height
        W = QC.Width
    Case 2
        H = QC.Height * 15
        W = QC.Width * 15
End Select

Dim Clr() As Long
Dim Cl As Long
SaberColor QC.BackColor, Clr()
Dim ya As Long

If Horizontal = True Then
    For ya = 0 To W
        Cl = ((ya / (W)) * 90)
        QC.Line (ya, 0)-(ya, H), RGB(Clr(0) + Cl, Clr(1) + Cl, Clr(2) + Cl)
    Next ya
Else
    For ya = 0 To H
        Cl = ((ya / (H)) * 90)
        QC.Line (0, ya)-(W, ya), RGB(Clr(0) + Cl, Clr(1) + Cl, Clr(2) + Cl)
    Next ya
End If

Dim BCl As Long
BCl = RGB(180, 180, 180)
End Sub

'/////////////////////////////////////////////////////////////
'/////////////////////////////////////////////////////////////
Private Sub SaberColor(Color As Long, Colores() As Long)
'Colores(0) = Rojo
'Colores(1) = Verde
'Colores(2) = Azul

ReDim Colores(2)

Dim Ra As Long
Dim Ga As Long
Dim Ba As Long


Dim XX As Long
XX = Color

'obtener el azul
Ba = XX \ (256 ^ 2)
'sacar el azul del numero largo
XX = XX - (Ba * (256 ^ 2))
'obtener el Green
Ga = XX \ 256
'sacar el green del numero
Ra = XX - (Ga * 256)
'lo que queda es el red

Colores(0) = Ra
Colores(1) = Ga
Colores(2) = Ba
End Sub


