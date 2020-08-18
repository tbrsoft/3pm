VERSION 5.00
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form fcsPlay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "csPlay (con WMP/IE5)"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "fcsPlay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MediaPlayerCtl.MediaPlayer ActiveMovie1 
      Height          =   1125
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   4005
      AudioStream     =   -1
      AutoSize        =   -1  'True
      AutoStart       =   0   'False
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   0
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   -600
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "fcsPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' fcsPlay                                                           (24/Jul/99)
' Formulario para tocar los ficheros usando el Windows Media Player del IE5
'
' ©Guillermo 'guille' Som, 1999
'------------------------------------------------------------------------------
Option Explicit

Private m_SegundosRestantes As Long

Private m_Terminado As Boolean

Private m_TiempoRestante As String
Private m_TiempoTotal As String

Private m_durMin As Double
Private m_durSec As Double

Private Resto As Long
Private durMinutos As Double
Private durSegundos As Double

Private m_EstadoActual As ecspEstado
Private m_FicheroCargado As Boolean

Private Sub ActiveMovie1_OpenComplete()
    ' Este evento ya no se produce en el control Windows Media Player
    ' Se recomienda usar ReadyStateChange.
    ' Desde ese evento llamo a este procedimiento                   (24/Jul/99)
    On Local Error Resume Next
    
    ' Aunque Microsoft dice que este evento está obsoleto, se usa...
    ' además también está en ReadyStateChange
    m_FicheroCargado = True
    m_Terminado = False
    
    ' Aquí se puede producir un error si no se puede tocar          (28/Jun/99)
    Err = 0
    ' Asignar la duración total en minutos y segundos
    Resto = ActiveMovie1.Duration
    
    If Err Then
        m_Terminado = True
        Resto = 0
    End If
    
    ' Los segundos restantes siempre serán los segundos por tocar
    m_SegundosRestantes = Resto
    
    m_durMin = Fix(Resto / 60)
    m_durSec = Resto - m_durMin * 60
    m_TiempoTotal = Format$(m_durMin, "00") & "." & Format$(m_durSec, "00")
    
    ' Estos valores son para el tiempo restante
    durMinutos = 0
    durSegundos = 0
    m_TiempoRestante = Format$(durMinutos, "00") & "." & Format$(durSegundos, "00")

    If Err Then
        m_TiempoRestante = "00:00 (ERROR)"
    End If
    
    Err = 0
End Sub

Private Sub ActiveMovie1_Timer()
    ' Este evento ya no está disponible en el Windows Media Player control
    ' Se llama desde ActiveMovie1_PositionChange
    On Local Error Resume Next
    
    ' En este evento se procesa la información a mostrar
    With ActiveMovie1
        Resto = .Duration - .CurrentPosition
        
        If Err Then
            Resto = 0
        End If
        
        ' Los segundos restantes siempre serán los segundos por tocar
        m_SegundosRestantes = Resto
        
        durMinutos = Fix(Resto / 60)
        durSegundos = Resto - durMinutos * 60
        m_TiempoRestante = Format$(durMinutos, "00") & "." & Format$(durSegundos, "00")
    End With
    
    Err = 0
End Sub

Private Sub ActiveMovie1_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
'
'Private Sub ActiveMovie1_StateChange(ByVal OldState As Long, ByVal NewState As Long)
'    ' Este evento no se produce en el Windows Media Player          (24/Jul/99)
'    ' En su lugar usar PlayStateChange
    '
    '
    ' Posibles valores de los parámetros xxxState
    ' mpStopped         0   Playback is stopped.
    ' mpPaused          1   Playback is paused.
    ' mpPlaying         2   Stream is playing.
    ' mpWaiting         3   Waiting for stream to begin.
    ' mpScanForward     4   Stream is scanning forward.
    ' mpScanReverse     5   Stream is scanning in reverse.
    ' mpSkipForward     6   Skipping to next.
    ' mpSkipReverse     7   Skipping to previous
    ' mpClosed          8   Stream is not open.
    '
    ' newState = 0-Stop, 1-Pausa, 2-Play
    ' Valores del ActiveMovie
    ' amvStopped    0   The player is stopped.
    ' amvPaused     1   The player is paused.
    ' amvRunning    2   The player is playing the multimedia file.
    '
    m_EstadoActual = NewState
    
    ' Nueva comprobación                                            (12/Dic/99)
    Select Case m_EstadoActual
    Case mpPaused, mpPlaying
        m_Terminado = False
    Case Else
        m_Terminado = True
        EMPEZAR_SIGUIENTE
    End Select
    
'    If m_EstadoActual = ecsStopped Then
'        m_Terminado = True
'    Else
'        m_Terminado = False
'    End If
End Sub

'Private Sub ActiveMovie1_PositionChange(ByVal oldPosition As Double, ByVal newPosition As Double)
'    ' Indica la posición anterior y la actual                       (24/Jul/99)
'    ' Los valores están dados en segundos...
'    '
'    ' Pero esto ocurre sólo si se manipula la posición con el control
'    '
'    If oldPosition <> newPosition Then
'        ' informar de que ha habido cambios
'        ' Esto ocurre cuando se Para la canción con .Parar
'        ' por tanto, quitarlo para que no se haga un lío...
'        ActiveMovie1_Timer
'    End If
'End Sub

Private Sub ActiveMovie1_ReadyStateChange(ReadyState As MediaPlayerCtl.ReadyStateConstants)
    ' Los valores posibles de ReadyState son:
    ' amvUninitialized  1   The FileName property has not been initialized.
    ' amvLoading        0   The ActiveMovie Control is asynchronously loading a file.
    ' amvInteractive    3   The control loaded a file,
    '                       and downloaded enough data to play the file,
    '                       but has not yet received all data.
    ' amvComplete       4   All data has been downloaded.
    '
    ' Los nombres y valores de la enumeración no han cambiado       (24/Jul/99)
    Select Case ReadyState
    Case amvInteractive
        m_FicheroCargado = True
    Case amvComplete
        ActiveMovie1_OpenComplete
    Case Else
        m_FicheroCargado = False
    End Select
End Sub


Private Sub Form_Initialize()
    ' Por si se cierra el formulario y se accede a alguna propiedad (28/Jun/99)
    m_Terminado = True
End Sub

Private Sub Form_Terminate()
    m_Terminado = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Dejar de tocar el fichero
    On Local Error Resume Next
    ActiveMovie1.Stop
    m_Terminado = True
    
    Err = 0
    'On Local Error GoTo 0
    
    Set fcsPlay = Nothing
End Sub

Friend Property Get FicheroCargado() As Boolean
    ' Propiedad de sólo lectura
    ' El valor se asignará al cargarse completamente el fichero
    FicheroCargado = m_FicheroCargado
End Property

Friend Property Get EstadoActual() As ecspEstado
    ' Devuelve el estado actual del fichero que se está tocando
    EstadoActual = m_EstadoActual
End Property

Friend Property Get TiempoTotal() As String
    On Local Error Resume Next
    
    ' Asignar la duración total en minutos y segundos
    Resto = ActiveMovie1.Duration
    m_durMin = Fix(Resto / 60)
    m_durSec = Resto - m_durMin * 60
    m_TiempoTotal = Format$(m_durMin, "00") & "." & Format$(m_durSec, "00")
    
    TiempoTotal = m_TiempoTotal
    
    Err = 0
    'On Local Error GoTo 0
End Property

Friend Property Get TiempoRestante() As String
    ' Ahora hay que "forzar" a crear la información:
    ActiveMovie1_Timer
    TiempoRestante = m_TiempoRestante
End Property

Friend Property Let Terminado(ByVal NewValue As Boolean)
    m_Terminado = NewValue
End Property

Friend Property Get Terminado() As Boolean
    Terminado = m_Terminado
End Property

Friend Property Get SegundosRestantes() As Long
    ' Devuelve los segundos que quedan por tocar
    On Local Error Resume Next
    
    Err = 0
    ' En este evento se procesa la información a mostrar
    With ActiveMovie1
        ' Sólo si está tocando
        If .PlayState = mpPlaying Then
            m_SegundosRestantes = .Duration - .CurrentPosition
            If Err Then
                m_SegundosRestantes = 0
            End If
        End If
    End With
    Err = 0
    SegundosRestantes = m_SegundosRestantes
End Property

