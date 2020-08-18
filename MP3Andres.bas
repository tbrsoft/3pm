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
    ' Tocar el fichero
    Set m_csplay = New cPlayWMP
    'm_csplay.Volumen = -4000
    m_csplay.FileName = Tema
    m_csplay.Volumen = frmINDEX.SLvolumen
    m_csplay.Tocar Tema
    ' El valor de cada paso del HScrollPos
    ESTOY_REPRODUCIENDO = True
    TEMA_REPRODUCIENDO = Tema
    frmINDEX.lblTemaSonando = FSO.GetBaseName(Tema) + " / " + FSO.GetParentFolderName(Tema)
    frmINDEX.Timer1.Interval = 1000 'reloj que cuenta el tiempo restante
End Sub
