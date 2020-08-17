VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPlayWMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'------------------------------------------------------------------------------
' cPlayWMP                                                          (12/Dic/99)
' Clase para "tocar" ficheros usando el Windows Media Player (quartz.dll)
'
' ©Guillermo 'guille' Som, 1999
'------------------------------------------------------------------------------
Option Explicit
Option Compare Text

Public Enum ecspEstado
    ecsStopped
    ecsPaused
    ecsRunning
    ' Valores del Windows Media Player:                             (12/Dic/99)
    mpStopped = 0           ' Playback is stopped.
    mpPaused = 1            ' Playback is paused.
    mpPlaying = 2           ' Stream is playing.
    mpWaiting = 3           ' Waiting for stream to begin.
    mpScanForward = 4       ' Stream is scanning forward.
    mpScanReverse = 5       ' Stream is scanning in reverse.
    mpSkipForward = 6       ' Skipping to next.
    mpSkipReverse = 7       ' Skipping to previous
    mpClosed = 8            ' Stream is not open.
End Enum

Private m_FileName As String
Private m_HayFichero As Boolean

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal NewValue As String)
    ' Asignar el nombre del fichero a tocar
    '
    Dim i As Long
    
    On Local Error Resume Next
    
    Err = 0
    
    ' Comprobar si existe
    ' Es mejor hacer una asignación, por si el path da error        (31/May/99)
    i = Len(Dir$(NewValue))
    If Err Then
        i = 0
        Err = 0
    End If
    
    If i Then
        ' Asignar el nombre del fichero a la variable y al control
        ' Tener en cuenta si es acceso directo
        m_FileName = Me.Lnk2Path(NewValue)
        fcsPlay.ActiveMovie1.FileName = m_FileName
        m_HayFichero = True
    Else
        'RaiseEvent HayError("FileName.Let:: No se ha encontrado el fichero indicado")
        m_HayFichero = False
        m_FileName = ""
        fcsPlay.ActiveMovie1.FileName = ""
        fcsPlay.Terminado = True
    End If
    
    If Err Then
        'RaiseEvent HayError("FileName.Let:: Se ha producido un error al cargar el fichero" & vbCrLf & Err.Description)
        m_HayFichero = False
        m_FileName = ""
        fcsPlay.ActiveMovie1.FileName = ""
        fcsPlay.Terminado = True
    End If
    
    Err = 0
    'On Local Error GoTo 0
End Property

Private Sub Class_Initialize()
    ' Cargar el formulario
    Load fcsPlay
End Sub

Private Sub Class_Terminate()
    ' Descargar el formulario
    Unload fcsPlay
End Sub

Public Sub Tocar(Optional sFileName As String = "")
    ' Empezar a tocar el fichero
    '
    Dim t1 As Double, t2 As Double
    
    On Local Error Resume Next
    
    Err = 0
    
    ' Si se especifica un fichero como parámetro, usar ese fichero
    If Len(sFileName) Then
        ' En la propiedad Let FileName se tiene en cuenta si es un acceso directo
        Me.FileName = sFileName
    End If
    '
    If m_HayFichero Then
        With fcsPlay
            t2 = TimeSerial(Hour(Now), Minute(Now), Second(Now) + 2)
            ' Esperar a que se cargue el fichero
            Do While .FicheroCargado = False
                DoEvents
                ' Si pasan 2 segundos y no se ha terminado de cargar...
                t1 = TimeSerial(Hour(Now), Minute(Now), Second(Now))
                If t1 > t2 Then
                    Exit Do
                End If
            Loop
            If .FicheroCargado Then
                ' Empezar a tocarlo
                ' Antes era .ActiveMovie.Run                        (24/Jul/99)
                ' Ahora no es .Run sino .Play
                .ActiveMovie1.Play
                'RaiseEvent Tocando(ecsRunning)
            End If
            'RaiseEvent AlTocar(.FicheroCargado)
        End With
    Else
        'RaiseEvent HayError("Tocar:: No se puede tocar el fichero")
        'RaiseEvent AlTocar(False)
        fcsPlay.Terminado = True
    End If
    
    Err = 0
    'On Local Error GoTo 0
End Sub

Public Sub Parar()
    ' Parar de tocar el fichero
    
    On Local Error Resume Next
    
    Err = 0
    
    With fcsPlay.ActiveMovie1
        .Stop
        ' Posicionarlo al principio
        .CurrentPosition = 0
    End With
    
    'If Err Then
    '    'RaiseEvent HayError("Parar:: Error al parar el fichero" & vbCrLf & Err.Description)
    'End If
    
    'RaiseEvent Tocando(ecsStopped)
    
    Err = 0
    'On Local Error GoTo 0
End Sub

Public Sub Pausa()
    ' Hacer una pausa o reanudar
    
    On Local Error Resume Next
    
    Err = 0
    
    With fcsPlay
        If .EstadoActual = ecsPaused Then
            '.ActiveMovie1.Run
            .ActiveMovie1.Play
            'RaiseEvent Tocando(ecsRunning)
        ElseIf .EstadoActual = ecsRunning Then
            .ActiveMovie1.Pause
            'RaiseEvent Tocando(ecsPaused)
        End If
    End With
    
    ' Si se produce aquí un error es que no se puede tocar
    ' o no hay fichero que se pueda tocar
    If Err Then
        'RaiseEvent HayError("Pausa:: Error al hacer pausa/reanudar" & vbCrLf & Err.Description)
        fcsPlay.Terminado = True
    End If
    
    Err = 0
    'On Local Error GoTo 0
End Sub

Public Property Get TiempoTotal() As String
    TiempoTotal = fcsPlay.TiempoTotal
End Property

Public Property Get TiempoRestante() As String
    TiempoRestante = fcsPlay.TiempoRestante
End Property

Public Property Get Terminado() As Boolean
    Terminado = fcsPlay.Terminado
End Property

Public Property Get FicheroCargado() As Boolean
    FicheroCargado = fcsPlay.FicheroCargado
End Property

Public Property Get SegundosRestantes() As Long
    SegundosRestantes = fcsPlay.SegundosRestantes
End Property

Public Property Get EstadoActual() As ecspEstado
    ' Devuelve el estado actual del fichero que se está tocando
    EstadoActual = fcsPlay.EstadoActual
End Property

Public Property Get Volumen() As Long
    ' Devuelve el volumen actual
    On Local Error Resume Next
    
    Volumen = fcsPlay.ActiveMovie1.Volume
    
    Err = 0
End Property

Public Property Let Volumen(ByVal NewValue As Long)
    ' Establece el volumen
    Dim VolumenAnt As Long
    
    On Local Error Resume Next
    Err = 0
    
    With fcsPlay.ActiveMovie1
        VolumenAnt = .Volume
        .Volume = NewValue
        If Err Then
            .Volume = VolumenAnt
        End If
    End With
    
    Err = 0
    'On Local Error GoTo 0
End Property

Public Sub HacerFade(Optional ByVal nSegundos As Long = 3&, _
                    Optional ByVal lngPasos As Long = 1&, _
                    Optional ByVal MinVol As Long = -5000&)
    ' Hace fade en el volumen, llevandolo desde el actual a cero,
    ' también permite hacerlo de menor a mayor
    '   En el control ActiveMovie, el volumen va desde -10000 a 0
    '   Siendo -10000 silencio total y 0 el volumen máximo
    '
    ' Hacer fade llevando el volumen hasta cero         ( 1/Ago/98)
    ' o durantes los segundos indicados
    '   por defecto es 3 seg.
    ' el valor de lngPasos es para la cuenta hacia atrás
    '   por defecto es -1
    '
    ' Ejemplo para hacer un fade para bajar el volumen:
    '   HacerFade 2, 1, -5000
    ' Ejemplo para hacer un fade para aumentar el volumen:
    '   HacerFade 2, 2, 0
    '
    Dim i As Long
    Dim j As Long, k As Long
    Dim horaActual As Date
    Dim tmpVolActual As Long
    
    On Local Error Resume Next
    
    Err = 0
    
    ' Siempre será un número positivo
    ' ya que el nivel de volumen va desde -10000 a 0
    lngPasos = Abs(lngPasos)
    
    horaActual = Now
    '
    tmpVolActual = Me.Volumen
    '
    ' El bucle se hará siempre desde el volumen actual al mínimo indicado
    j = tmpVolActual
    k = MinVol
    If j > k Then lngPasos = -lngPasos
    For i = j To k Step lngPasos
        Me.Volumen = i
        DoEvents
        If Second(Now - horaActual) > nSegundos Then
            Exit For
        End If
    Next
    
    If Err Then
        Me.Volumen = 0
    End If
    
    Err = 0
End Sub

Public Function Lnk2Path(ByVal sFic As String) As String
    '--------------------------------------------------------------------------
    ' Convertir un fichero con extensión .LNK en el nombre completo (14/May/99)
    '--------------------------------------------------------------------------
    '
    ' Devuelve el path y nombre del fichero de un fichero .LNK
    ' La estructura de un fichero de extensión .LNK es:
    ' >>>Empezando por el final:<<<
    ' (Esto no es válido para los links de MS-DOS)
    '   4 Nulos
    '   Path de inicio completo
    '   1 Nulo
    '   1 byte con la longitud del path
    '
    '   Si el byte que precede a longitud no es un nulo:
    '       Nombre del fichero
    '       1 Nulo
    '       1 byte con la longitud del nombre
    '       (el nombre seguramente empezará con .\)
    '   Si el byte que precede a la longitud es un nulo:
    '       1 Nulo (este byte comprobado)
    '       Nombre completo
    '       1 Nulo
    '       Nombre del Path en formato UNC (\\Computadora\Recurso)
    '       1 Nulo
    '--------------------------------------------------------------------------
    '
    Dim nFic As Long
    Dim sTmp As String
    Dim i As Long
    
    On Error Resume Next
    
    ' Leer el contenido del fichero .LNK
    If InStr(sFic, ".lnk") = 0 Then
        ' Si no tiene la extensión LNK devolver el mismo fichero
        Lnk2Path = sFic
        Err = 0
        Exit Function
    End If
    
    nFic = FreeFile
    Open sFic For Binary As nFic
    i = LOF(nFic)
    ' Esto puede dar error de fin de fichero
    'sTmp = input$(LOF(nFic), nFic)
    sTmp = Space$(i)
    Get nFic, , sTmp
    Close nFic
    
    Dim j As Long
    Dim c As Long
    Dim s1 As String    ' s1 será el path
    Dim s2 As String    ' s2 será el nombre del fichero
    
    ' Recorrerlo desde el final
    ' pero saltarse los cuatro últimos
    For i = Len(sTmp) - 4 To 1 Step -1
        c = Asc(Mid$(sTmp, i, 1))
        ' Si es un Nulo...
        If c = 0 Then
            If Len(s1) = 0 Then
                ' Es el path
                j = Asc(Mid$(sTmp, i - 1, 1))
                s1 = Mid$(sTmp, i + 1, j)
                i = i - 1
                j = 1
                nFic = 0
                If Asc(Mid$(sTmp, i - 1, 1)) = 0 Then
                    i = i - 1
                    j = 0
                    nFic = i
                End If
            ElseIf Len(s2) = 0 Then
                ' Es el nombre del fichero
                If j = 0 And nFic > 0 Then
                    s2 = Mid$(sTmp, i + 1, nFic - i - 1)
                    nFic = i
                Else
                    j = Asc(Mid$(sTmp, i - 1, 1))
                    s2 = Mid$(sTmp, i + 1, j)
                    nFic = 0
                End If
            Else
                Exit For
            End If
        End If
    Next
    ' El path devuelto será:
    '   El path de inicio + el nombre del fichero
    '
    ' Añadirle la barra de directorios a s1 y quitarle la que tenga s2
    If Right$(s1, 1) <> "\" Then
        s1 = s1 & "\"
    End If
    ' Tomar sólo el nombre del fichero sin el path
    For i = Len(s2) To 1 Step -1
        If Mid$(s2, i, 1) = "\" Then
            s2 = Mid$(s2, i + 1)
            Exit For
        End If
    Next
    Lnk2Path = s1 & s2
    
    Err = 0
End Function

Public Property Get CurrentPosition() As Double
    CurrentPosition = fcsPlay.ActiveMovie1.CurrentPosition
End Property

Public Property Let CurrentPosition(ByVal NewValue As Double)
    On Error Resume Next
    fcsPlay.ActiveMovie1.CurrentPosition = NewValue
    If Err Then
        fcsPlay.ActiveMovie1.CurrentPosition = 0
    End If
    Err = 0
End Property

Public Property Get Duration() As Double
    ' Devuelve la duración total del fichero                        (15/Dic/99)
    Duration = fcsPlay.ActiveMovie1.Duration
End Property
