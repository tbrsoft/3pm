Attribute VB_Name = "MP3Andres"
Option Explicit
Option Compare Text

Private CambiandoPos As Boolean
Private dPos As Double

Public sLista As String
Public sEstadoActual() As String
Public sUnidad As String

Public m_Tocando As Boolean
Public m_TocandoLista As Boolean
Public m_queFichero As Long

' Clase para manejar el fichero a tocar
Public m_csplay As cPlayWMP

Public m_CD As cComDlg

Public Sub EjecutarTema(Tema As String)
    OnOffCAPS vbKeyCapital, True
    ' Tocar el fichero
    Set m_csplay = New cPlayWMP
    'm_csplay.Volumen = -4000
    m_csplay.FileName = Tema
    m_csplay.Volumen = frmINDEX.SLvolumen
    m_csplay.Tocar Tema
    ' El valor de cada paso del HScrollPos
    ESTOY_REPRODUCIENDO = True
    TEMA_REPRODUCIENDO = Tema
    frmINDEX.lblTemaSonando = FSO.GetBaseName(Tema) + " / " + FSO.GetBaseName(FSO.GetParentFolderName(Tema))
    frmINDEX.Timer1.Interval = 1000 'reloj que cuenta el tiempo restante
End Sub

Public Sub EMPEZAR_SIGUIENTE()
    With frmINDEX
        .Timer1.Interval = 0
        .lblTiempoRestante = "Restante 0:00"
        'si hay algun elemento en la lista ejecutarlo
        If UBound(MATRIZ_LISTA) > 0 Then
            Dim TemaDeMatriz As String
            TemaDeMatriz = txtInLista(MATRIZ_LISTA(1), 0, ",")
            'reacomodar la matriz para quitar el primer elemento
            Dim c As Long
            For c = 1 To UBound(MATRIZ_LISTA)
                If c < UBound(MATRIZ_LISTA) Then
                    'cuando sea cualquiera menos el ultimo
                    MATRIZ_LISTA(c) = MATRIZ_LISTA(c + 1)
                Else
                    'cuando sea el ultimo
                    'redefinir la matriz con un indice menos
                    ReDim Preserve MATRIZ_LISTA(c - 1)
                End If
            
            Next
            EjecutarTema TemaDeMatriz
            CargarProximosTemas
        Else
            'si no hay temas mostrar la leyenda que lo indica
            OnOffCAPS vbKeyCapital, False
            .lblTiempoRestante = "Restante 0:00"
            .lblTemaSonando = "Sin reproduccion actual"
            .lblProximoTema = "No hay próximo tema"
            TEMA_REPRODUCIENDO = "Sin reproduccion actual"
            ESTOY_REPRODUCIENDO = False
        
        End If
    End With

End Sub
