VERSION 5.00
Begin VB.UserControl VUMeter 
   BackColor       =   &H00000000&
   ClientHeight    =   5085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7740
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   ScaleHeight     =   5085
   ScaleWidth      =   7740
   Begin VB.PictureBox picWMP 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404000&
      DrawWidth       =   4
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   4935
      Left            =   810
      ScaleHeight     =   4875
      ScaleWidth      =   6075
      TabIndex        =   0
      Top             =   60
      Width           =   6135
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   5
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   1230
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   5
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   1260
      Width           =   735
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   4
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   990
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   4
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   1020
      Width           =   735
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   3
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   750
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   3
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   780
      Width           =   735
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   2
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   510
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   2
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   540
      Width           =   735
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   1
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   270
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   1
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   300
      Width           =   735
   End
   Begin VB.Shape P2 
      BackStyle       =   1  'Opaque
      Height          =   255
      Index           =   0
      Left            =   6960
      Shape           =   4  'Rounded Rectangle
      Top             =   30
      Width           =   735
   End
   Begin VB.Shape P 
      BackStyle       =   1  'Opaque
      Height          =   250
      Index           =   0
      Left            =   50
      Shape           =   4  'Rounded Rectangle
      Top             =   60
      Width           =   735
   End
End
Attribute VB_Name = "VUMeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mIsPlaying As Boolean
Private Const MAXPNAMELEN = 32  '  longitud máx. del nombre del producto (incluido NULL)
Private Type WAVEOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        dwFormats As Long
        wChannels As Integer
        dwSupport As Long
End Type

Private Type WaveFormatEx
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long 'wavehdr_tag
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer      'wMid
    ProductID As Integer       'wPid
    DriverVersion As Long       'MMVERSIONS vDriverVersion
    ProductName(1 To 32) As Byte 'szPname[MAXPNAMELEN]
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Private Const WAVE_INVALIDFORMAT = &H0&                 '/* invalid format */
Private Const WAVE_FORMAT_1M08 = &H1&                   '/* 11.025 kHz, Mono,   8-bit
Private Const WAVE_FORMAT_1S08 = &H2&                   '/* 11.025 kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_1M16 = &H4&                   '/* 11.025 kHz, Mono,   16-bit
Private Const WAVE_FORMAT_1S16 = &H8&                   '/* 11.025 kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_2M08 = &H10&                  '/* 22.05  kHz, Mono,   8-bit
Private Const WAVE_FORMAT_2S08 = &H20&                  '/* 22.05  kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_2M16 = &H40&                  '/* 22.05  kHz, Mono,   16-bit
Private Const WAVE_FORMAT_2S16 = &H80&                  '/* 22.05  kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_4M08 = &H100&                 '/* 44.1   kHz, Mono,   8-bit
Private Const WAVE_FORMAT_4S08 = &H200&                 '/* 44.1   kHz, Stereo, 8-bit
Private Const WAVE_FORMAT_4M16 = &H400&                 '/* 44.1   kHz, Mono,   16-bit
Private Const WAVE_FORMAT_4S16 = &H800&                 '/* 44.1   kHz, Stereo, 16-bit
Private Const WAVE_FORMAT_QUERY = &H1
Private Const WAVE_FORMAT_DIRECT = &H8
Private Const WAVE_FORMAT_DIRECT_QUERY = (WAVE_FORMAT_QUERY Or WAVE_FORMAT_DIRECT)
Private Const WAVE_FORMAT_PCM = 1   '  Necesario en archivos de recursos para #ifndef RC_INVOKED
Private Const WAVE_VALID = &H3         '  ;Interno




Private Const WHDR_DONE = &H1&              '/* done bit */
Private Const WHDR_PREPARED = &H2&          '/* set if this header has been prepared */
Private Const WHDR_BEGINLOOP = &H4&         '/* loop start block */
Private Const WHDR_ENDLOOP = &H8&           '/* loop end block */
Private Const WHDR_INQUEUE = &H10&          '/* reserved for driver */

Private Const WIM_OPEN = &H3BE
Private Const WIM_CLOSE = &H3BF
Private Const WIM_DATA = &H3C0
Private Const WAVE_MAPPER = -1&
'flags de WaveOutOpen
Private Const WAVE_MAPPED = &H4


Private Type WAVEFORMAT
        wFormatTag As Integer
        nChannels As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
End Type

Private Type MMTIME
        wType As Long
        u As Long
End Type

'funciones de WaveIn
Private Declare Function waveInAddBuffer Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInPrepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long
Private Declare Function waveInUnprepareHeader Lib "winmm" (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer As Long, ByVal WaveHdrStructSize As Long) As Long

Private Declare Function waveInGetNumDevs Lib "winmm" () As Long
Private Declare Function waveInGetDevCaps Lib "winmm" Alias "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal WaveInCapsPointer As Long, ByVal WaveInCapsStructSize As Long) As Long

Private Declare Function waveInOpen Lib "winmm" (WaveDeviceInputHandle As Long, ByVal WhichDevice As Long, ByVal WaveFormatExPointer As Long, ByVal CallBack As Long, ByVal CallBackInstance As Long, ByVal flags As Long) As Long
Private Declare Function waveInClose Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Private Declare Function waveInStart Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInReset Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long
Private Declare Function waveInStop Lib "winmm" (ByVal WaveDeviceInputHandle As Long) As Long

Dim VerdeOFF As Long
Dim AmarilloOFF As Long
Dim RojoOFF As Long
Dim VerdeON As Long
Dim AmarilloON As Long
Dim RojoON As Long

Dim DevHandle As Long
Dim LH As Long, RH As Long 'alto de las barra L y R
Dim InData(0 To 511) As Byte
Dim m_Dispositivos As Long
Dim m_inHabilitado As Boolean
Dim Devices() As String
Dim BarrasEnVUmeter As Long
Dim m_CarFantastic As Boolean 'efecto auto fantastico. Se usa cuando no reproduce para mostrar algo cheto
Dim m_Borde As Long 'efecto auto fantastico. Se usa cuando no reproduce para mostrar algo cheto
Dim m_EspacioEntreBarras As Long
Dim contVu As Long
Dim contVuOff As Long
Dim MaxLH As Long, MaxRH As Long
Dim ContTopVU As Long

Public Property Get AnchoBarra() As Long
    'Valor del ancho de las barras. Sirve para saber que zona esta libre
    AnchoBarra = P(c).Width
End Property

Private Sub UserControl_Initialize()
    'esto se inicia solo cuando se carga el control en ejecucuion
    'inicializar los dispositivos
    m_CarFantastic = False
    m_Borde = False
    m_EspacioEntreBarras = 25
    m_Dispositivos = Dispositivos
    BarrasEnVUmeter = 6
    VerdeOFF = &H808000
    AmarilloOFF = &H8080&
    RojoOFF = &H80&
    VerdeON = &HFF00&
    AmarilloON = &HFFFF&
    RojoON = &HFF&
End Sub

Property Get Dispositivos()
    'solo lectura
    Dim Caps As WaveInCaps, Which As Long
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(Caps), Len(Caps))
        ''If Caps.Formats And WAVE_FORMAT_1M08 Then
        'if Caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
        ''se cargan todos los dispositivos
            'Call DevicesBox.AddItem(StrConv(Caps.ProductName, vbUnicode), Which)
            ReDim Preserve Devices(Which)
            Devices(Which) = StrConv(Caps.ProductName, vbUnicode)
            'ccc.AddItem StrConv(Caps.ProductName, vbUnicode), Which
        'End If
    Next
    If Which = 0 Then
        'inhabilitar las barras
        m_inHabilitado = True
        'MsgBox "You have no audio output devices!", vbCritical, "Ack!"
        End
    Else
        m_inHabilitado = False
    End If
    
    Dispositivos = Which
End Property
Property Get NombreDispositivo(Indice As Long) As String
    'solo lecture
    If Indice > UBound(Devices) Then
        NombreDevispositivo = "No existe"
    Else
        NombreDevispositivo = Devices(Indice)
    End If
End Property

Property Let CarFantastic(new_CarFantastic As Boolean)
    m_CarFantastic = new_CarFantastic
    contVu = 0
    PropertyChanged "CarFantastic"
End Property

Property Get CarFantastic() As Boolean
    CarFantastic = m_CarFantastic
End Property

Property Let EspacioEntreBarras(new_EspacioEntreBarras As Long)
    m_EspacioEntreBarras = new_EspacioEntreBarras
    PropertyChanged "EspacioEntreBarras"
    Call UserControl_Resize
End Property

Property Get EspacioEntreBarras() As Long
    EspacioEntreBarras = m_EspacioEntreBarras
End Property


Property Let Borde(new_Borde As Long)
    If new_Borde > 1 Then new_Borde = 1
    If new_Borde < 0 Then new_Borde = 0
    m_Borde = new_Borde
    c = 0
    Do While c < BarrasEnVUmeter
        P(c).BorderStyle = Val(m_Borde)
        P2(c).BorderStyle = Val(m_Borde)
        c = c + 1
    Loop
    PropertyChanged "Borde"
End Property

Property Get Borde() As Long
    Borde = m_Borde
End Property


Property Get inHabilitado() As Boolean
    'solo lectura
    inHabilitado = m_inHabilitado
End Property

Public Sub DoStop()
    mIsPlaying = False
    Call waveInReset(DevHandle)
    Call waveInClose(DevHandle)
    DevHandle = 0

End Sub

Public Sub DoStart()
    mIsPlaying = True
    'DoStop
    Static WAVEFORMAT As WaveFormatEx
    With WAVEFORMAT
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2 'Two channels -- left and right
        .SamplesPerSec = 11025
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    
    Dim OpIN As Long

    
    OpIN = waveInOpen(DevHandle, 0, VarPtr(WAVEFORMAT), 0, 0, 0)
    'ver si ya esta abierto
    If OpIN = 4 Then
        'esta abierto. Cargar en DevHandle el valor de DevHandle devuelto en la primera apertura
        DevHandle = Val(ReadFile("c:\devhw.dat"))
        'maercar como secundario
    End If
    If OpIN = 0 Then
        'se abrio ok debo grabar un archivo temporal para que lo lea este OCX cuando
        'se abra nuevamente
        WriteFile "c:\DevHw.dat", CStr(DevHandle)
    End If
        
    If DevHandle = 0 Then
        m_inHabilitado = True
        UserControl.BackColor = vbRed
        'Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!")
        Exit Sub
    End If
    UserControl.BackColor = vbBlack
    'si es secundario no vuelve a abrir nada
    waveInStart (DevHandle)
    m_inHabilitado = False
    Call Visualize
    
End Sub

Property Let ChannelOut(New_ChannelOUT As Long)
    If New_ChannelOUT = 1 Then
        m_ChannelOUT = 1
    Else
        m_ChannelOUT = 2
    End If
End Property

Property Get ChannelOut() As Long
    ChannelOut = m_ChannelOUT
End Property
Private Sub DrawData()
    DoEvents
    'si esta reproduciendo mostrar vumetro, si no mostrar el auto fantastico
    If m_CarFantastic Then
        'BarrasEnVUmeter es el numero de barras
        PorcBarrasPintadasON = contVu / BarrasEnVUmeter * 100
        'verde
        If PorcBarrasPintadasON <= 20 And PorcBarrasPintadasON >= 0 Then ColorOn = &HFF00&: coloroff = &H8000&
        'amarillo
        If PorcBarrasPintadasON <= 75 And PorcBarrasPintadasON > 20 Then: ColorOn = &HFFFF&: coloroff = &H8080&
        'rojo
        If PorcBarrasPintadasON <= 100 And PorcBarrasPintadasON > 75 Then: ColorOn = &HFF&: coloroff = &H80&
        
        P(contVu).BackColor = ColorOn
        P2(contVu).BackColor = ColorOn
        
        Dim LargoGusano As Long
        LargoGusano = BarrasEnVUmeter / 3
        If contVu >= LargoGusano Then
            contVuOff = contVu - LargoGusano
        Else
            contVuOff = BarrasEnVUmeter - (LargoGusano - contVu)
        End If
        
        PorcBarrasPintadasOff = contVuOff / BarrasEnVUmeter * 100
        'verde
        If PorcBarrasPintadasOff <= 20 And PorcBarrasPintadasOff >= 0 Then ColorOn = &HFF00&: coloroff = &H8000&
        'amarillo
        If PorcBarrasPintadasOff <= 75 And PorcBarrasPintadasOff > 20 Then: ColorOn = &HFFFF&: coloroff = &H8080&
        'rojo
        If PorcBarrasPintadasOff <= 100 And PorcBarrasPintadasOff > 75 Then: ColorOn = &HFF&: coloroff = &H80&
        
        P(contVuOff).BackColor = coloroff
        P2(contVuOff).BackColor = coloroff
        
        contVu = contVu + 1
        If contVu = BarrasEnVUmeter Then contVu = 0
    Else
        'Plot the data...
        'los impares son un canal y los pares el otro
        LH = (InData(1) - 120) '120 es cuando no hay sonido
        RH = (InData(2) - 120) '120 es cuando no hay sonido
        Dim TopeVU As Long 'maximo al que supongo que llegará
        TopeVU = 150
        'BarrasEnVUmeter es el numero de barras
        contVu = 0: a = 0: b = 0
        ContTopVU = ContTopVU + 1
        'cada 18 vuelta se fija el tope de vuelta
        If ContTopVU > 18 Then
            MaxLH = 0
            MaxRH = 0
            ContTopVU = 0
        End If
        'MaxLH = BarrasEnVUmeter: MaxRH = BarrasEnVUmeter esta en resize para que no se cambien los valores
        Do While contVu < BarrasEnVUmeter
            PorcBarrasPintadas = contVu / BarrasEnVUmeter * 100
            'verde
            If PorcBarrasPintadas <= 20 And PorcBarrasPintadas >= 0 Then ColorOn = &HFF00&: coloroff = &H8000&
            'amarillo
            If PorcBarrasPintadas <= 75 And PorcBarrasPintadas > 20 Then: ColorOn = &HFFFF&: coloroff = &H8080&
            'rojo
            If PorcBarrasPintadas <= 100 And PorcBarrasPintadas > 75 Then: ColorOn = &HFF&: coloroff = &H80&

            If contVu = MaxLH Then GoTo SiguienteRh 'que no repinte la ubicacion del tope
            
            If LH > TopeVU / BarrasEnVUmeter * (contVu + 1) Then
                P(contVu).BackColor = ColorOn
            Else
                If a = 0 Then
                    If contVu >= MaxLH Then
                        MaxLH = contVu
                        P(contVu).BackColor = vbBlack
                    Else
                        MaxLH = MaxLH - 1
                        P(MaxLH).BackColor = vbBlack
                    End If
                    a = a + 1
                Else
                    P(contVu).BackColor = coloroff
                End If
            End If
SiguienteRh:
            If contVu = MaxRH Then GoTo SIGUIENTE
            If RH > TopeVU / BarrasEnVUmeter * (contVu + 1) Then
                P2(contVu).BackColor = ColorOn
            Else
                If b = 0 Then
                    If contVu >= MaxRH Then
                        MaxRH = contVu
                        P2(contVu).BackColor = vbBlack
                    Else
                        MaxRH = MaxRH - 1
                        P2(MaxRH).BackColor = vbBlack
                    End If
                    b = b + 1
                Else
                    P2(contVu).BackColor = coloroff
                End If
            End If
SIGUIENTE:
            MostrarWMP RH, LH, contVu, 2
            contVu = contVu + 1
        Loop
    End If
End Sub

Private Sub Visualize()
    Static Wave As WaveHdr
    Wave.lpData = VarPtr(InData(0))
    Wave.dwBufferLength = 512 'This is now 512 so there's still 256 samples per channel
    Wave.dwFlags = 0
    Do
        Call waveInPrepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        Call waveInAddBuffer(DevHandle, VarPtr(Wave), Len(Wave))
        Do
            'Nothing -- we're waiting for the audio driver to mark
            'this wave chunk as done.
            DoEvents
        Loop Until ((Wave.dwFlags And WHDR_DONE) = WHDR_DONE) Or DevHandle = 0
    
        Call waveInUnprepareHeader(DevHandle, VarPtr(Wave), Len(Wave))
        If DevHandle = 0 Then
            'The device has closed...
            Exit Do
        End If
        Call DrawData
        DoEvents
    Loop While DevHandle <> 0  'While the audio device is open
End Sub

Private Sub UserControl_Resize()
    On Local Error Resume Next
    'ver cuantas barras tiene que haber
    Dim BarrasToShow As Long
    BarrasToShow = UserControl.Height / (P(0).Height + m_EspacioEntreBarras)
    If BarrasToShow > BarrasEnVUmeter Then
        'agregar barras
        Do While BarrasEnVUmeter < BarrasToShow
            Load P(BarrasEnVUmeter)
            Load P2(BarrasEnVUmeter)
            P(BarrasEnVUmeter).Visible = True
            P2(BarrasEnVUmeter).Visible = True
            BarrasEnVUmeter = BarrasEnVUmeter + 1
        Loop
    Else
        'quitar barras solo si hasta las 12 originales
        Do While BarrasEnVUmeter > BarrasToShow And BarrasEnVUmeter > 6
            Unload P(BarrasEnVUmeter - 1)
            Unload P2(BarrasEnVUmeter - 1)
            BarrasEnVUmeter = BarrasEnVUmeter - 1
        Loop
    End If
    'reubicar todas las barras
    c = 0
    Do While c < BarrasEnVUmeter
        If c = 0 Then
            P(c).Top = UserControl.Height - P(c).Height - m_EspacioEntreBarras
        Else
            P(c).Top = P(c - 1).Top - P(c).Height - m_EspacioEntreBarras
            P(c).Left = P(0).Left
        End If
        c = c + 1
    Loop
    'darles a todos el mimo ancho
    c = 0
    Do While c < BarrasEnVUmeter
        'P(c).Width = (UserControl.Width - 150) / 2
        c = c + 1
    Loop
    'acomodar los P2
    c = 0
    Do While c < BarrasEnVUmeter
        P2(c).Width = P(c).Width
        P2(c).Height = P(c).Height
        P2(c).Left = UserControl.Width - P2(c).Width 'P(c).Left + P(c).Width + 20
        P2(c).Top = P(c).Top
        c = c + 1
    Loop
    
    
End Sub

Private Sub WriteFile(Arch As String, Texto As String)
    libre = FreeFile
    Open Arch For Output As libre
        Write #libre, Texto
    Close libre
End Sub

Private Function ReadFile(Arch As String) As String
    libre = FreeFile
    Open Arch For Input As libre
        Input #libre, Texto
    Close libre
    ReadFile = Texto
End Function

Private Sub UserControl_Terminate()
    DoStop
End Sub

Public Property Get IsPlaying() As Boolean
    IsPlaying = mIsPlaying
End Property

Public Sub MostrarWMP(R As Long, L As Long, Cont As Long, MODO As Long)
    'se pasan los valores de los canales R y L
    'de ahí tengo que inventar algo
    'cada 18 vueltas se fija el topo, aprovecho ese numero para hace alguna cosa mas
    If Cont = 0 Then
        'limpiar
        picWMP.Cls
    End If
    R = Abs(R)
    L = Abs(L)
    Select Case MODO
        Case 1 ' 2 pelotas amarilla y roja
            picWMP.FillColor = vbRed
            picWMP.Circle (picWMP.Width / 2 + Int(Rnd * 1000), picWMP.Height / 2), Abs(R) * 10, vbRed
            picWMP.FillColor = vbYellow
            picWMP.Circle (picWMP.Width / 2 - Int(Rnd * 1000), picWMP.Height / 2), Abs(L) * 10, vbYellow
        Case 2 ' 2 pelotas amarilla y roja
            picWMP.BackColor = 1000000 + (R + L)
            For a = 1 To 2
                picWMP.FillColor = 1000000 + (R + L)
                picWMP.Circle (Int(Rnd * picWMP.Width / 2), Int(Rnd * picWMP.Height)), R + L, 10000 + (R + L)
                picWMP.Circle (picWMP.Width / 2 + Int(Rnd * picWMP.Width / 2), Int(Rnd * picWMP.Height)), R + L, 100000 + (R + L)
            
            Next
            
            
        Case 3 ' 2 pelotas amarilla y roja
            picWMP.FillColor = vbBlue
            picWMP.Circle (0, 0), Abs(R) * 10, vbBlue
            picWMP.FillColor = vbWhite
            picWMP.Circle (picWMP.Width, picWMP.Height), Abs(L) * 10, vbWhite
            picWMP.FillColor = vbBlue
            picWMP.Circle (picWMP.Width / 2 + Int(Rnd * 1000), picWMP.Height / 2), Abs(R) * 10, vbBlue
            picWMP.FillColor = vbWhite
            picWMP.Circle (picWMP.Width / 2 - Int(Rnd * 1000), picWMP.Height / 2), Abs(L) * 10, vbWhite
        Case 4 ' 2 pelotas amarilla y roja
            picWMP.DrawWidth = 6
            picWMP.FillColor = vbBlue
            picWMP.Circle (0, 0), R * 30, vbBlue
            picWMP.FillColor = vbWhite
            picWMP.Circle (picWMP.Width, 0), L * 30, vbWhite
            picWMP.ForeColor = vbBlue
            picWMP.Line (R * 30, L * 30)-(L * 30, R * 30)
            picWMP.ForeColor = vbWhite
            picWMP.Line (picWMP.Width - R * 30, L * 30)-(picWMP.Width - L * 30, R * 30)
        Case 5 ' 2 pelotas amarilla y roja
            picWMP.DrawWidth = 3
            picWMP.FillColor = vbBlue
            picWMP.Circle (0, 0), R * 5, vbBlue
            picWMP.FillColor = vbWhite
            picWMP.Circle (picWMP.Width, 0), L * 5, vbWhite
            
            picWMP.FillColor = vbWhite
            picWMP.Circle (R * 30, L * 30), L
            
            picWMP.FillColor = vbBlue
            picWMP.Circle (L * 30, R * 30), L
            
            picWMP.FillColor = vbWhite
            picWMP.Circle (L * 30, L * 30), L
            
            picWMP.FillColor = vbBlue
            picWMP.Circle (R * 30, R * 30), L
            
            picWMP.FillColor = vbWhite
            picWMP.Circle (R * 20, L * 20), L
            
            picWMP.FillColor = vbBlue
            picWMP.Circle (L * 20, R * 20), L
            
        Case 6
            picWMP.DrawWidth = 3
            
            'pelotitas de x,y fijos
            picWMP.FillColor = 50 + R * 140
            picWMP.Circle (picWMP.Width / 4, picWMP.Height / 4), R * 5, vbBlack
            picWMP.FillColor = 5000000 + L * 140
            picWMP.Circle (picWMP.Width * 0.75, picWMP.Height * 0.75), L * 5, vbBlack
            
            'picWMP.FillColor = R * 140
            'picWMP.Circle (picWMP.Width / 5, picWMP.Height / 5), R, R * 140
            'picWMP.FillColor = L * 140
            'picWMP.Circle (picWMP.Width * 0.8, picWMP.Height * 0.8), L, L * 140
            
            'pelotitas de x,y variables
            
            'Randomize Timer
            a = Int(Rnd * picWMP.Width / 4 + R * 10)
            b = Int(Rnd * picWMP.Height / 4 + R * 10)
            picWMP.FillColor = a + b
            picWMP.Circle (a, b), R
            
            c = Int(Rnd * picWMP.Width / 4 + R * 10)
            d = Int(Rnd * picWMP.Height / 4 + R * 10)
            c = c + picWMP.Width * 0.5
            d = d + picWMP.Height * 0.5
            picWMP.FillColor = c + d
            picWMP.Circle (c, d), R
        Case 7
            For a = 1 To 5
                Dim AltoBarra As Long
                AltoBarra = (picWMP.Height / 2) / 15
                
                picWMP.DrawWidth = AltoBarra - (R)
                picWMP.ForeColor = Int(Rnd * R * 4)
                picWMP.Line (0, Int(Rnd * (a * AltoBarra)))-(picWMP.Width, Int(Rnd * (a * AltoBarra)))
                
                picWMP.DrawWidth = AltoBarra - (L)
                picWMP.ForeColor = Int(Rnd * L)
                picWMP.Line (0, picWMP.Height - Int(Rnd * (a * AltoBarra)))-(picWMP.Width, picWMP.Height - Int(Rnd * (a * AltoBarra)))
            Next
    End Select
End Sub
