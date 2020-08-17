VERSION 5.00
Begin VB.Form ftPlayAM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproductor de audio con Windows Media Player (IE5)"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "ftPlay.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   7425
   Begin VB.TextBox txtLista 
      Height          =   315
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   16
      Text            =   "txtLista"
      Top             =   5820
      Width           =   6315
   End
   Begin VB.HScrollBar HScrollVol 
      Height          =   195
      LargeChange     =   100
      Left            =   2640
      Max             =   0
      Min             =   -10000
      SmallChange     =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   960
      Width           =   1155
   End
   Begin VB.HScrollBar HScrollPos 
      Height          =   195
      LargeChange     =   5
      Left            =   900
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   1635
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Guardar lista..."
      Height          =   405
      Left            =   2250
      TabIndex        =   14
      Top             =   5310
      Width           =   1515
   End
   Begin VB.TextBox txtUnidad 
      Height          =   315
      Left            =   900
      TabIndex        =   13
      Text            =   "C:"
      Top             =   5310
      Width           =   585
   End
   Begin VB.CommandButton cmdFade 
      Caption         =   "&Fade"
      Height          =   345
      Left            =   3960
      TabIndex        =   7
      Top             =   660
      Width           =   1215
   End
   Begin VB.CommandButton cmdFic 
      Height          =   345
      Index           =   0
      Left            =   6090
      Picture         =   "ftPlay.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   " Tocar la lista "
      Top             =   5310
      Width           =   345
   End
   Begin VB.CommandButton cmdFic 
      Height          =   345
      Index           =   1
      Left            =   6480
      Picture         =   "ftPlay.frx":0590
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   " Hacer Pausa/Reanudar la lista "
      Top             =   5310
      Width           =   345
   End
   Begin VB.CommandButton cmdFic 
      Height          =   345
      Index           =   2
      Left            =   6870
      Picture         =   "ftPlay.frx":06DE
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   " Parar la lista "
      Top             =   5310
      Width           =   345
   End
   Begin VB.ListBox lstLista 
      DragIcon        =   "ftPlay.frx":082C
      Height          =   3660
      Left            =   210
      OLEDropMode     =   1  'Manual
      Style           =   1  'Checkbox
      TabIndex        =   11
      Top             =   1500
      Width           =   6975
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   60
      Top             =   600
   End
   Begin VB.CommandButton cmdFic1 
      Height          =   345
      Index           =   2
      Left            =   6360
      Picture         =   "ftPlay.frx":0C6E
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   " Parar de tocar el fichero "
      Top             =   660
      Width           =   345
   End
   Begin VB.CommandButton cmdFic1 
      Height          =   345
      Index           =   1
      Left            =   5970
      Picture         =   "ftPlay.frx":0DBC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   " Hacer Pausa/Reanudar "
      Top             =   660
      Width           =   345
   End
   Begin VB.CommandButton cmdFic1 
      Height          =   345
      Index           =   0
      Left            =   5580
      Picture         =   "ftPlay.frx":0F0A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   " Tocar el fichero desde el principio "
      Top             =   660
      Width           =   345
   End
   Begin VB.CommandButton cmdExaminar 
      Caption         =   "..."
      Height          =   345
      Index           =   0
      Left            =   6750
      TabIndex        =   2
      ToolTipText     =   " Seleccionar el fichero o la lista a usar "
      Top             =   210
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   900
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Text            =   "C:\Sonidos\Emilia-Big big World.wav"
      Top             =   210
      Width           =   5775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000E&
      Index           =   1
      X1              =   120
      X2              =   7260
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   7260
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Label Label1 
      Caption         =   "&Lista:"
      Height          =   255
      Index           =   3
      Left            =   210
      TabIndex        =   15
      Top             =   5850
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "&Unidad:"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   12
      Top             =   5340
      Width           =   645
   End
   Begin VB.Label lblVolumen 
      BackColor       =   &H00000000&
      Caption         =   "lblVolumen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   5
      Top             =   690
      Width           =   1155
   End
   Begin VB.Label lblTiempo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "lblTiempo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   900
      TabIndex        =   3
      Top             =   690
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   " F&ichero:"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "ftPlayAM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' Prueba de csPlayWMP                                               (12/Dic/99)
'
' ©Guillermo 'guille' Som, 1999
'------------------------------------------------------------------------------
Option Explicit
Option Compare Text

Private CambiandoPos As Boolean
Private dPos As Double

Private sLista As String
Private sEstadoActual() As String
Private sUnidad As String

Private m_Tocando As Boolean
Private m_TocandoLista As Boolean
Private m_queFichero As Long

' Clase para manejar el fichero a tocar
Private m_csPlay As cPlayWMP

Private m_CD As cComDlg

Private Sub cmdExaminar_Click(Index As Integer)
    On Local Error Resume Next
    
    With m_CD
        .hWnd = Me.hWnd
        .FileName = Text1
        .Filter = "Tipos admitidos (*.wav; *.mp3; *.m3u)|*.wav;*.mp3;*.m3u|Ficheros Wav (*.wav)|*.wav|Ficheros MP3 (*.mp3)|*.mp3|Lista de ficheros (*.m3u)|*.m3u"
        .CancelError = True
        .ShowOpen
        If Err = 0 Then
            Text1 = .FileName
            If InStr(.FileName, ".m3u") = 0 Then
                m_csPlay.FileName = .FileName
            Else
                sLista = .FileName
                ' abrir la lista
                AbrirLista
            End If
        End If
    End With
    
    Err = 0

End Sub

Private Sub cmdFade_Click()
    Static Invertir As Boolean
    
    If Invertir Then
        cmdFade.Caption = "&Fade"
        DoEvents
        m_csPlay.HacerFade 4, 1, 0
    Else
        cmdFade.Caption = "&Restaurar"
        DoEvents
        m_csPlay.HacerFade 3, 1, -5000
    End If
    HScrollVol.Value = m_csPlay.Volumen
    Invertir = Not Invertir
End Sub

Private Sub cmdFic_Click(Index As Integer)
    ' Comandos para tocar la lista
    Select Case Index
    Case 0  ' Tocar
        m_TocandoLista = True
        TocarLista
    Case 1  ' Pausa
        If m_csPlay.EstadoActual = ecsPaused Then
            m_csPlay.Tocar
        Else
            m_csPlay.Pausa
        End If
    Case 2  ' Parar
        m_TocandoLista = False
    End Select
End Sub

Private Sub cmdFic1_Click(Index As Integer)
    ' Tocar el fichero
    
    On Error Resume Next
    
    Select Case Index
    Case 0
        m_csPlay.Tocar Text1
        
        ' El valor de cada paso del HScrollPos
        dPos = m_csPlay.Duration / HScrollPos.Max
        HScrollPos.Value = 0
        HScrollVol.Value = m_csPlay.Volumen
        
        m_Tocando = True
    Case 1
        ' Pausa
        If m_csPlay.EstadoActual = ecsPaused Then
            m_csPlay.Tocar
        Else
            m_csPlay.Pausa
        End If
    Case 2
        ' Parar
        m_csPlay.Parar
        m_Tocando = False
    End Select
    '
    Err = 0
End Sub

Private Sub cmdGuardar_Click()
    ' Guardar el contenido de la lista                              (22/Ago/99)
    Dim sFic As String
    Dim nFic As Long
    Dim i As Long
    
    On Error Resume Next
    
    With m_CD
        .hWnd = Me.hWnd
        .DialogTitle = "Guardar lista"
        .FileName = txtLista
        .Filter = "Lista MP3 y TXT (*.m3u; *.txt)|*.m3u;*.txt|Lista MP3 (*.m3u)|*.m3u|Ficheros de texto (*.txt)|*.txt"
        .CancelError = True
        .ShowSave
        If Err Then
            Err = 0
            Exit Sub
        Else
            sLista = .FileName
            txtLista = sLista
        End If
    End With
    
    On Local Error GoTo ErrGuardar
    
    If lstLista.ListCount Then
        sFic = sLista
        nFic = FreeFile
        Open sFic For Output As nFic
        For i = 0 To lstLista.ListCount - 1
            Print #nFic, lstLista.List(i)
        Next
        Close nFic
    End If
    
    Exit Sub
ErrGuardar:
    Close
    MsgBox "Se ha producido el error:" & vbCrLf & _
           Err.Number & " " & Err.Description
    
    Err = 0
End Sub

Private Sub Form_Load()
    Const cSonidos As String = "C:\Sonidos\"
    Dim sFic As String
    Dim sTmp As String
    Dim nFic As Long
    Const cMsg As String = "Prueba de csPlayWMP"
    Dim i As Long
    
    HScrollPos.Max = 100
    HScrollVol.min = -10000&
    HScrollVol.Max = 0&
    
    ' Para mostrar el estado actual                                 (12/Dic/99)
    ' ecspEstado.mpStopped =0, ecspEstado.mpClosed = 8
    ReDim sEstadoActual(ecspEstado.mpStopped To ecspEstado.mpClosed)
    sEstadoActual(ecspEstado.mpClosed) = "Closed"
    sEstadoActual(ecspEstado.mpPaused) = "Paused"
    sEstadoActual(ecspEstado.mpPlaying) = "Playing"
    sEstadoActual(ecspEstado.mpScanForward) = "ScanForward"
    sEstadoActual(ecspEstado.mpScanReverse) = "ScanReverse"
    sEstadoActual(ecspEstado.mpSkipForward) = "SkipForward"
    sEstadoActual(ecspEstado.mpSkipReverse) = "SkipReverse"
    sEstadoActual(ecspEstado.mpStopped) = "Stopped"
    sEstadoActual(ecspEstado.mpWaiting) = "Waiting"
    
    '-------------------------------------------------------------- (22/Ago/99)
    ' Comprobar si los ficheros de prueba están en la unidad indicada
    txtUnidad = "E:"
    ComprobarUnidad
    
    Text1 = sUnidad & "\Sonidos\Jennifer Lopez-If you had my love mini.wav"
    
    Timer1.Enabled = False
    
    lblTiempo = ""
    lblVolumen = " vol: "
    
    ' Crear los objetos
    Set m_CD = New cComDlg
    Set m_csPlay = New cPlayWMP
    
    ' Asignar el volumen a 0 (normal), ya que en el IE5 se asigna a -600
    m_csPlay.Volumen = 0
    
    ' Llenar la lista de ficheros
    ' (también se admiten ficheros de texto con los ficheros a tocar)
    sFic = AppPath & "\Lista.m3u"
    sLista = sFic
    txtLista = sFic
    
    If Len(Dir$(sFic)) Then
        AbrirLista
    Else
        With lstLista
            .Clear
            sFic = Dir$(cSonidos & "*.wav")
            Do While Len(sFic)
                .AddItem cSonidos & sFic
                sFic = Dir$
            Loop
        End With
    End If
    
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_csPlay = Nothing
    
    Set m_CD = Nothing
    
    Set ftPlayAM = Nothing
End Sub

Private Sub HScrollPos_Change()
    On Error Resume Next
    
    If CambiandoPos = False Then
        CambiandoPos = True
        m_csPlay.CurrentPosition = HScrollPos.Value * dPos
        CambiandoPos = False
    End If
    
    Err = 0
End Sub

Private Sub HScrollVol_Change()
    'On Error Resume Next
    m_csPlay.Volumen = HScrollVol.Value
    'Err = 0
End Sub

Private Sub lstLista_DragDrop(Source As Control, X As Single, Y As Single)
    ListRowMove Source, DragIndex, ListRowCalc(Source, Y)
End Sub

Private Sub lstLista_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyDelete Then
        With lstLista
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    .RemoveItem i
                End If
            Next
        End With
    End If
End Sub

Private Sub lstLista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DragIndex = ListRowCalc(lstLista, Y)
    lstLista.Drag
End Sub

Private Sub lstLista_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Añadir los ficheros soltados
    Dim i As Long
    Dim j As Long
    
    ' Posicionarlos antes del que está seleccionado                 (22/Ago/99)
    j = lstLista.ListIndex
    If j < 0 Then j = 0
    
    With Data
        For i = 1 To .Files.Count
            lstLista.AddItem .Files(i), j
        Next
    End With
End Sub

Private Sub Text1_DragDrop(Source As Control, X As Single, Y As Single)
    If TypeName(Source) = "ListBox" Then
        Text1 = Source.List(Source.ListIndex)
    End If
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1 = Data.Files(1)
End Sub

Private Sub Timer1_Timer()
    
    On Local Error Resume Next
    
    With m_csPlay
        ' Si se está tocando el fichero...
        If m_Tocando Then
            ' mostrar el tiempo restante
            lblTiempo = .TiempoRestante & " (" & sEstadoActual(.EstadoActual) & ")"
            
            ' Mostrar el valor de la barra de desplazamiento,
            ' (si no se está cambiando en este momento)
            If CambiandoPos = False Then
                CambiandoPos = True
                HScrollPos.Value = .CurrentPosition / dPos
                CambiandoPos = False
            End If
            
            ' Si se ha terminado
            If .Terminado Then
                m_Tocando = False
            End If
        Else
            ' Si no se está tocando, mostrar el tiempo total,
            ' si hay un fichero cargado
            If .FicheroCargado Then
                lblTiempo = .TiempoTotal & " (" & sEstadoActual(.EstadoActual) & ")"
            Else
                lblTiempo = "---"
            End If
        End If
        ' El 0 es el valor máximo del volumen
        If .Volumen = 0 Then
            lblVolumen = " vol: " & .Volumen & " (max)"
        Else
            lblVolumen = " vol: " & .Volumen
        End If
    End With
    
    Err = 0
End Sub

Private Sub TocarLista()
    ' Tocar los ficheros de la lista
    Dim i As Long, j As Long
    Dim sFic As String
    
    j = lstLista.ListCount - 1
    i = 0
    Label1(0).BackColor = vbRed
    Do While i <= j
        If Not lstLista.Selected(i) Then
            sFic = lstLista.List(i)
            lstLista.Selected(i) = True
            
            ' Comprobar si es un link y hay que modificar la unidad (14/Dic/99)
            If InStr(sFic, ".lnk") Then
                sFic = m_csPlay.Lnk2Path(sFic)
                If Left$(sFic, 2) <> sUnidad Then
                    sFic = sUnidad & Mid$(sFic, 3)
                End If
            End If
            
            m_csPlay.FileName = sFic
            
            ' Sólo si es un fichero "aceptable"
            If m_csPlay.TiempoTotal > 0 Then
                m_Tocando = True
                '
                m_csPlay.Tocar sFic
                Text1 = sFic
                
                ' El valor de cada paso del HScrollPos
                dPos = m_csPlay.Duration / HScrollPos.Max
                HScrollPos.Value = 0
                
                HScrollVol.Value = m_csPlay.Volumen
                '
                ' Esperar a que termine de tocar
                With m_csPlay
                    Do While .Terminado = False
                        If m_TocandoLista = False Then
                            .Parar
                            ' Obligar a salir del bucle de tocar las canciones
                            i = j
                            Exit Do
                        End If
                        DoEvents
                    Loop
                End With
            End If
        End If
        i = i + 1
    Loop
    Label1(0).BackColor = vbButtonFace
End Sub

Private Sub ComprobarUnidad()
    ' Comprueba si la unidad indicada está disponible               (22/Ago/99)
    ' (se buscará el directorio \Sonidos y algún fichero WAV)
    Dim sFic As String
    Dim i As Long
    Dim j As Long
    Static YaEstoy As Boolean
    
    ' No permitir la reentrada
    If YaEstoy Then Exit Sub
    
    YaEstoy = True
        
    On Local Error Resume Next
    
    sUnidad = Trim$(txtUnidad)
    sFic = sUnidad & "\Sonidos\*.wav"
    
    i = Len(Dir$(sFic))
    If Err <> 0 Or i = 0 Then
        ' No está disponible, buscar otra unidad
        For j = Asc("C") To Asc("Z")
            Err = 0
            sUnidad = Chr$(j) & ":"
            sFic = sUnidad & "\Sonidos\*.wav"
            i = Len(Dir$(sFic))
            ' Si no se produce error es que hemos encontrado algo
            If Err = 0 And i <> 0 Then
                txtUnidad = sUnidad
                YaEstoy = False
                Exit Sub
            End If
        Next
        ' Si se llega aquí es que no hemos hallado nada...
        txtUnidad = Left$(CurDir$, 2)
        ' Asignar también sUnidad,                                  (12/Dic/99)
        ' sino se quedaría la última comprobada
        sUnidad = txtUnidad
    End If
    
    Err = 0
    
    YaEstoy = False
End Sub

Private Sub txtUnidad_KeyPress(KeyAscii As Integer)
    ' Si se pulsa INTRO, comprobar la unidad escrita                (22/Ago/99)
    If KeyAscii = 13 Then
        KeyAscii = 0
        ComprobarUnidad
        AsignarFicheros
    End If
End Sub

Private Sub AsignarFicheros()
    ' Asignar la nueva unidad a los ficheros                        (22/Ago/99)
    Dim j As Long
    Dim sFic As String
    
    sFic = Text1
    If Left$(sFic, 2) <> sUnidad Then
        Text1 = sUnidad & Mid$(sFic, 3)
    End If
    
    With lstLista
        For j = 0 To .ListCount - 1
            sFic = .List(j)
            If Left$(sFic, 2) <> sUnidad Then
                .List(j) = sUnidad & Mid$(sFic, 3)
            End If
        Next
    End With
End Sub

Private Sub AbrirLista()
    Dim nFic As Long
    Dim sFic As String
    Dim i As Long
    Dim sTmp As String
    
    sFic = sLista
    
    nFic = FreeFile
    If Len(Dir$(sFic)) Then
        lstLista.Clear
        
        On Local Error Resume Next
        
        Open sFic For Input As nFic
        Do While Not EOF(nFic)
            Line Input #nFic, sTmp
            Err = 0
            ' Si no existe ese fichero,
            i = Len(Dir$(sTmp))
            If Err <> 0 Or i = 0 Then
                ' quitar el nombre de la unidad y asignar la hallada
                i = InStr(sTmp, ":")
                If i Then
                    sTmp = Mid$(sTmp, i + 1)
                End If
                sTmp = sUnidad & sTmp
            End If
            lstLista.AddItem sTmp
        Loop
        Close nFic
        
        Err = 0
    End If
End Sub

Private Function AppPath() As String
    ' Quita la barra de directorios,                                (20/Dic/99)
    ' y lo devuelve con la primera en mayúsculas y el resto en minúsculas
    Dim sTmp As String
    
    sTmp = App.Path
    If Right$(sTmp, 1) = "\" Then
        sTmp = Left$(sTmp, Len(sTmp) - 1)
    End If
    AppPath = StrConv(sTmp, vbProperCase)
End Function
