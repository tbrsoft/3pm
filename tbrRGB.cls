VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "tbrRGB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'el color long de ole_color es en realidad
'(ROJO * (256 ^ 2)) + (Green * 256) + Blue

'R G y B son de cero a 255

Public Function AcercarColores(ColorINI As Long, ColorFin As Long, Var As Byte) As Long
    'devuelve el color mas cerca a color fin
    'VAR es de 1 a 255 e indica la variacion de cada componente de un color
    'NO PUEDE ser negativo, el signo lo define segun la diferencia de cada color
    
    Dim R As Long, G As Long, B As Long
    GetRGB ColorINI, R, G, B
    
    Dim R2 As Long, G2 As Long, B2 As Long
    GetRGB ColorFin, R2, G2, B2
    
    Dim tmpVAR As Integer
    '***********ROJO*********
    If R < R2 Then
        R = R + Var
        If R > 255 Then R = 255
        If R > R2 Then R = R2
    Else
        R = R - Var
        If R < 0 Then R = 0
        If R < R2 Then R = R2
    End If
    '***********GREEN*********
    If G < G2 Then
        G = G + Var
        If G > 255 Then G = 255
        If G > G2 Then G = R2
    Else
        G = G - Var
        If G < 0 Then G = 0
        If G < G2 Then G = G2
    End If
    '***********BLUE*********
    If B < B2 Then
        B = B + Var
        If B > 255 Then B = 255
        If B > B2 Then B = B2
    Else
        B = B - Var
        If B < 0 Then B = 0
        If B < B2 Then B = B2
    End If
       
    AcercarColores = RGB(R, G, B)
    
End Function

Public Function ChangeColor(ColorINI As Long, _
                            VarRojo As Integer, VarGreen As Integer, VarBlue As Integer, _
                            TopRojo As Integer, TopGreen As Integer, TopBlue As Integer) _
                            As Long

    'devuelve el color final
    'los parametros vienen integer para poder tener valores negativos
    'no necesita mucho calculo para llegar a otro color buscado gradualmente ya _
    que el tope lo para
    
    Dim R As Integer, G As Integer, B As Integer
    GetRGB ColorINI, CByte(R), CByte(G), CByte(B)
    
    R = R + VarRojo
    If R < 0 Then R = 0
    If (VarRojo < 0 And R < TopRojo) Or (VarRojo > 0 And R > TopRojo) Then R = TopRojo
    If R > 255 Then R = 255
    
    G = G + VarGreen
    If G < 0 Then G = 0
    If (VarGreen < 0 And G < TopGreen) Or (VarGreen > 0 And G > TopGreen) Then G = TopGreen
    If G > 255 Then G = 255

    B = B + VarBlue
    If B < 0 Then B = 0
    If (VarBlue < 0 And B < TopBlue) Or (VarBlue > 0 And B > TopBlue) Then B = TopBlue
    If B > 255 Then B = 255
    
    ChangeColor = RGB(R, G, B)
    
End Function

Public Function GetRGB(lColor As Long, _
                       ByRef bROJO As Long, ByRef bGREEN As Long, ByRef bBLUE As Long) _
                       As Byte
    
    Dim R As Long, G As Long, B As Long
    Dim Resto As Long
    
    Resto = lColor
    'EXTRAER EL ROJO
    If lColor >= 65536 Then
        B = Resto \ 65536
    Else
        B = 0
    End If
    
    'DE LO QUE QUEDA SACAR EL GREEN
    Resto = Resto - (B * 65536)
    If Resto >= 256 Then
        G = Resto \ 256
    Else
        G = 0
    End If
    Resto = Resto - (G * 256)
    'LO QUE QUEDA ES EL AZUL
    R = Resto
    
    bROJO = R
    bGREEN = G
    bBLUE = B
    
End Function


