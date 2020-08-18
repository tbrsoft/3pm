Attribute VB_Name = "Globales"
'para el teclado
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function SetKeyboardState Lib "user32" (lppbKeyState As Byte) As Long


Public FSO As New Scripting.FileSystemObject
Public AP As String
Public CREDITOS As Long ' fichas cargadas (o temas habilitados para cargar)
Public TEMA_REPRODUCIENDO As String 'tema actual. Para poder mostrar el texto
'si no hay nada el valor es "sin reproduccion actual"
Public TEMA_SIGUIENTE As String 'tema actual. Para poder mostrar el texto
'si no hay nada el valor es "no hay proximo tema"
Public TEMAS_EN_LISTA 'numero de temas a reproducir despues del actual
Public ESTOY_REPRODUCIENDO As Boolean 'saber si hay tema en curso
Public TIEMPO_RESTANTE_TEMA_ACTUAL As Long 'tiempo en segundos restante
Public ESTOY As Integer 'indica en que pantalla estoy
'estoy vale 0=seleccion de disco, 1= dentro de un disco viendo los temas
Public MATRIZ_DISCOS() As String 'path,nombrecarpeta
Public MATRIZ_TEMAS() As String 'path,nombreTema. se usa solo para cargar lstTemas,
'este los ordena alfabeticamente
'despues se toma ubicacionActual+lstTemas+".mp3"
Public MATRIZ_TOTAL() As String '(Carpdisco,PathTema/duracion)
Public MATRIZ_LISTA() As String 'lista de temas a reproducir. No incluye el TEMA_REPRODUCIENDO
Public TOTAL_DISCOS As Long ' total de discos
Public UbicDiscoActual As String 'path del disco actual
'sirve para no usar la MATRIZ_TEMAS y poder ordenar los temas de cada disco
Public WAIT_EMPIEZA As Integer 'esperar 5 segundos por comienzo de tema

Public Function txtInLista(lista As String, Orden As Integer, Separador As String) As String
    'devuelve "OUT LISTA" si se solicita un orden no existente
    'separador es la "," o "-"
    Dim lAct As String, lOrden As Integer
    Dim palabra(40) As String
    Dim c As Integer
    c = 1: lOrden = 0
    Do While c <= Len(lista)
        lAct = Mid(lista, c, 1)
        If lAct = Separador Then
            lOrden = lOrden + 1
        Else
            palabra(lOrden) = palabra(lOrden) + lAct
        End If
        c = c + 1
    Loop
    'si oreden solicitado>ultimo oreden de la lista...
    If Orden > lOrden Then txtInLista = "OUT LISTA": Exit Function
    txtInLista = palabra(Orden)
End Function

Public Sub CargarProximosTemas()

    'cargar lblProximoTema
    Dim strProximos As String, TotTemas As Integer
    For c = 1 To UBound(MATRIZ_LISTA)
        'el indice 0 no existe ni existira por eso el C+1
        strProximos = strProximos + txtInLista(MATRIZ_LISTA(c), 1, ",")
        strProximos = strProximos + vbCrLf
    Next
    frmINDEX.lblProximoTema = strProximos
    TotTemas = UBound(MATRIZ_LISTA)
    frmINDEX.lblTemasEnLista = "En lista: " + Trim(Str(TotTemas))

End Sub

Public Sub OnOffCAPS(vKey As KeyCodeConstants, PRENDER As Boolean)
    Dim keys(255) As Byte
    ' leer el estado actual del teclado
    GetKeyboardState keys(0)
    ' invertir el bit 0 de la tecla virtual en la que estamos interesados
    ' keys(vKey) = keys(vKey) Xor 1
    If PRENDER Then
        keys(vKey) = 1
    Else
        keys(vKey) = 0
    End If
    ' forzar el nuevo estado del teclado
    SetKeyboardState keys(0)
End Sub
