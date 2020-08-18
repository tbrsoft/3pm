VERSION 5.00
Begin VB.UserControl MP3Play 
   BackColor       =   &H00FF00FF&
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1500
   PropertyPages   =   "MP3Play.ctx":0000
   ScaleHeight     =   1620
   ScaleWidth      =   1500
   ToolboxBitmap   =   "MP3Play.ctx":0011
   Begin VB.Timer Reloj 
      Index           =   3
      Left            =   570
      Top             =   510
   End
   Begin VB.Timer Reloj 
      Index           =   2
      Left            =   570
      Top             =   30
   End
   Begin VB.Timer Reloj 
      Index           =   1
      Left            =   120
      Top             =   510
   End
   Begin VB.Timer Reloj 
      Index           =   0
      Left            =   120
      Top             =   30
   End
End
Attribute VB_Name = "MP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Long, ByVal lpstrBuffer As String, ByVal uLength As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private Declare Function mciGetDeviceID Lib "winmm.dll" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long
Private Declare Function midiOutSetVolume Lib "winmm.dll" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long

'Property Variables:
Dim m_FileName(4) As String 'ahora son 4 (para kundera 6.8.200)
Dim m_Volumen(4) As Long 'ahora son 4 (para kundera 6.8.200)
Dim dwReturn As Long

Private Alias(4) As String
'lista de alias para usar en enganches o similares

Event Played(SecondsPlayed As Long, iAlias As Long)
Event BeginPlay(iAlias As Long)
Event EndPlay(iAlias As Long)

Private TMPs As String

Private Sub Reloj_Timer(Index As Integer)
    On Error GoTo ERmp3
    tERR.Anotar "002-0001"
    'primero ver si ermina el tema
    If IsPlaying(CLng(Index)) = False Then
        RELOJ(Index).Interval = 0
        RaiseEvent EndPlay(CLng(Index))
        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
    End If
    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
    tERR.Anotar "002-0005"
    RaiseEvent Played(PositionInSec(CLng(Index)), CLng(Index))
    Exit Sub
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayctl" + ".acpw"
    Resume Next
End Sub

Private Sub UserControl_Initialize()
    Alias(0) = "MP3Play0"
    Alias(1) = "Mp3Play1"
    Alias(2) = "Mp3Play2"
    Alias(3) = "Mp3Play3"
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1620
    UserControl.Width = 1500
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
End Property

Public Property Get Volumen(iAlias As Long) As Long
    On Error GoTo ERmp3
    tERR.Anotar "002-0011", iAlias
    m_Volumen(iAlias) = m_Volumen(iAlias) / 10
    tERR.Anotar "002-0012"
    Volumen = m_Volumen(iAlias)
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpv"
End Property

Public Property Let Volumen(iAlias As Long, ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi máquina anda del 0 al 1000
    tERR.Anotar "002-0013"
    m_Volumen(iAlias) = New_Volumen * 10 ' * 30 - 3000
    TMPs = "SetAudio " + Alias(iAlias) + " Volume To " + CStr(m_Volumen(iAlias))
    tERR.Anotar "002-0014", iAlias, TMPs, IAA, IAANext
    Ret = mciSendString(TMPs, 0, 0, 0)
    tERR.Anotar "002-0015", Ret
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        tERR.AppendLog "NoVolumenEn:" + CStr(Ret), "MpPalyCtl" + ".acpw"
    End If
    tERR.Anotar "002-0018"
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPalyCtl" + ".acpx"
    Resume Next
End Property

Public Property Get FileName(iAlias As Long) As String
    FileName = m_FileName(iAlias)
End Property

Public Property Let FileName(iAlias As Long, ByVal New_FileName As String)
    m_FileName(iAlias) = New_FileName
End Property

Public Function isPlayingAny() As Boolean
    Dim TmpB As Boolean
    TmpB = False
    Dim A44 As Long
    For A44 = 0 To 3
        If IsPlaying(A44) Then
            TmpB = True
            Exit For
        End If
    Next A44
    isPlayingAny = TmpB
End Function

Public Function IsPlaying(iAlias As Long) As Boolean
    
    On Error GoTo ERmp3
    tERR.Anotar "002-0027"
    If m_FileName(iAlias) = "" Then
        tERR.Anotar "002-0028", iAlias
        IsPlaying = False
    Else
        tERR.Anotar "002-0029"
        Static s As String * 30
        tERR.Anotar "002-0030", HabilitarVUMetro, NoVumVID
        Ret = mciSendString("status " + Alias(iAlias) + " mode", s, Len(s), 0)
        tERR.Anotar "002-0031", Ret, iAlias
        If Ret = 263 Then '263 es cuando no ha abierto nada
            IsPlaying = False
            Exit Function
        End If
        If Ret <> 0 Then
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            tERR.AppendLog "ERR IsPlaying=Status:" + CStr(Ret)
            'WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        'EN ESTE CASO ES NULO O ALGO ASI
        'YA QUE MCI NO TIENE LA CAPACIDAD DE STATUS!!!
        If Ret = 274 Then
            tERR.Anotar "002-0034b", iAlias
            IsPlaying = True
        Else
            tERR.Anotar "002-0034", s
            IsPlaying = (Mid(s, 1, 7) = "playing")
        End If
        
        IsPlaying = (Mid(s, 1, 7) = "playing")
    End If
    Exit Function
ERmp3:
    tERR.Anotar "002-0034b"
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpx"
    Resume Next
End Function

Public Function DoOpen(iAlias As Long)
    'DoStop    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    On Error GoTo ERmp3
    tERR.Anotar "002-0035"
    Dim Ret As String * 128
    tERR.Anotar "002-0036", iAlias, Alias(iAlias)
    Ret = mciSendString("Close " + Alias(iAlias), 0, 0, 0)
    tERR.Anotar "002-0037"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "NoCierraMCI.RET:" + CStr(Ret), m_FileName(iAlias)
    End If
    tERR.Anotar "002-0040"
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    tERR.Anotar "002-0041"
    If FSO.FileExists(m_FileName(iAlias)) = False Then        '
        tERR.Anotar "002-0042", iAlias
        tERR.AppendLog "MpPlayCtl.DoOpen.NoExistFile.acqb", m_FileName(iAlias)
        tERR.Anotar "002-0043"
        Exit Function
    End If
    tERR.Anotar "002-0044"
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0045", m_FileName(iAlias)
    FileNameSHORT = Left$(TMP, lenShort)
    tERR.Anotar "002-0046"
    glo_hWnd = hwnd
    tERR.Anotar "002-0047"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " style child"
    tERR.Anotar "002-0048"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0049"
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            tERR.AppendLog "WINDOWS NO REPRODUCE MP3!!!. INSTALE EL REPRODUCTOR CORESPONDIENTE " + _
                "A SU VERSION DE WINDOWS"
            MsgBox "No se ha podido abrir el fichero debido a un problema existente en Windows. " + vbCrLf + _
                "Revise que el reproductor multimedia de Windows este instalado y funcione correctamente." + _
                "Notifique a tbrSoft de esto para más detalles"
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        tERR.AppendLog "DoOpen.NoAbre." + CStr(dwReturn), "acqe"
    End If
    tERR.Anotar "002-0054"
    'uso todo en milisegundos
    dwReturn = mciSendString("set " + Alias(iAlias) + _
        " time format milliseconds", 0, 0, 0)
    tERR.Anotar "002-0055"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "FormatMilisec.acqf", "MpPlayCtl"
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acdy"
    Resume Next
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, _
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, iAlias As Long)
    
    On Error GoTo ERmp3
    'DoStop
    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    tERR.Anotar "002-0058"
    Dim Ret As String * 128
    tERR.Anotar "002-0059", iAlias
    Ret = mciSendString("Close " + Alias(iAlias), 0, 0, 0)
    tERR.Anotar "002-0060", Ret
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "acqga", "MpPlayCtl"
    End If
    tERR.Anotar "002-0063"
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    tERR.Anotar "002-0065"
    If Dir(m_FileName(iAlias)) = "" Then
        tERR.Anotar "002-0065b"
        tERR.AppendLog "NoEx.acqh", m_FileName(iAlias)
        Exit Function
    End If
    tERR.Anotar "002-0068"
    lenShort = GetShortPathName(m_FileName(iAlias), TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0069", m_FileName(iAlias)
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = Left$(TMP, lenShort)
    tERR.Anotar "002-0070", FileNameSHORT
    glo_hWnd = hwnd
    tERR.Anotar "002-0071", HabilitarVUMetro, NoVumVID, HWind
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias " + _
        Alias(iAlias) + " style " + _
        Style + " parent " + CStr(HWind)
        
    tERR.Anotar "002-0072", cmdToDo '                         xxxxxx
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0073", dwReturn, Salida2, Style, HWind
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlay.DoOpenVid.acqh." + CStr(dwReturn), m_FileName(iAlias)
    End If
    tERR.Anotar "002-0076", X1, X2, Y1, Y2
    cmdToDo = "put " + Alias(iAlias) + " window at " + CStr(X1) + " " + CStr(Y1) + _
        " " + CStr(X2) + " " + CStr(Y2) + " "
    tERR.Anotar "002-0077", Salida2, Style, HWind
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    'lo uso para ver info de los videos que andan bien tambien
    'tERR.AppendLog "ADD77:" + CStr(dwReturn), m_FileName
    If dwReturn <> 0 Then
        '********************
        'si es 346 es que no tiene ventana de presentacion (base+90=MCIERR_NO_WINDOW)
        '¿¿¿¿¿????????
        'probe con style popup y overlapped y no sirven ni solucionan
        If dwReturn = 346 Then
            tERR.AppendLog "MCIERR_NO_WINDOW=346. No hay ventana de presentacion!!!" + CStr(dwReturn), m_FileName(iAlias)
            'pasa con videos con codecs nuevos que al cargarse si funciona en
            'WMP pero no en 3PM!!!!!!!!!!!!!!!!!!!
        '********************
        Else
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlay.DoOpenVid.WindowAt.acqi." + CStr(dwReturn), m_FileName(iAlias)
        End If
    End If
    
    'Dim s As String
    's = Space(30)
    'cmdToDo = "capability MP3Play can stretch"
    'dwReturn = mciSendString(cmdToDo, s, Len(s), 0)
    'If dwReturn <> 0 Then MsgBox "Error n°: " + Str(dwReturn)
   '
   ' s = Trim(Left(s, 4))
   ' If UCase(s) = "TRUE" Then
   '     cmdToDo = "window MP3Play stretch"
   '     dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
   '     If dwReturn <> 0 Then MsgBox "Error n°: " + Str(dwReturn)
   ' End If
        
    'uso todo en milisegundos
    tERR.Anotar "002-0081"
    mciSendString "set " + Alias(iAlias) + " time format milliseconds", 0, 0, 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqj"
    Resume Next
End Function

Public Function DoPlay(iAlias As Long, Optional FullScreen As Boolean = False)
    On Error GoTo ERmp3
    tERR.Anotar "002-0082", CStr(FullScreen), iAlias
    If FullScreen Then
        dwReturn = mciSendString("play " + Alias(iAlias) + " fullscreen", 0, 0, 0)
    Else
        dwReturn = mciSendString("play " + Alias(iAlias), 0, 0, 0)
    End If
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPlay.Play." + m_FileName(iAlias), ".acqk"
    End If
    tERR.Anotar "002-0086"
    RELOJ(iAlias).Interval = 1000
    tERR.Anotar "002-0087"
    RaiseEvent BeginPlay(iAlias)
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acql"
    Resume Next
End Function

Public Function DoPause(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0088"
    dwReturn = mciSendString("pause " + Alias(iAlias), 0, 0, 0)
    tERR.Anotar "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPause." + m_FileName(iAlias), ".acqm"
    End If
    tERR.Anotar "002-0092"
    RELOJ(iAlias).Interval = 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqn"
    Resume Next
End Function

Public Function DoStop(iAlias As Long) As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0093"
    dwReturn = mciSendString("stop " + Alias(iAlias), 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoStop." + m_FileName(iAlias), ".acqo"
    End If
    tERR.Anotar "002-0097"
    RELOJ(iAlias).Interval = 0
    tERR.Anotar "002-0098"
    RaiseEvent EndPlay(iAlias)
    
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqp"
    Resume Next
End Function

Public Function DoClose(iAlias As Long) As String
    'iAlias es el que hay que cerrar o:
        '99 cierra todos
        
    On Error GoTo ERmp3
    tERR.Anotar "002-0099", iAlias
    If iAlias = 99 Then 'cierra todos
        Dim F11 As Long
        For F11 = 0 To 3
            dwReturn = mciSendString("close " + Alias(F11), 0, 0, 0)
            If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
                LogErrorMCI dwReturn
                tERR.AppendLog "MpPlayCtl.DoClose." + m_FileName(F11), ".acqr"
            End If
            RELOJ(F11).Interval = 0
        Next F11
    Else 'o solo el elegido
        dwReturn = mciSendString("close " + Alias(iAlias), 0, 0, 0)
        If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
            LogErrorMCI dwReturn
            tERR.AppendLog "MpPlayCtl.DoClose." + m_FileName(iAlias), ".acqr"
        End If
        tERR.Anotar "002-0103"
        RELOJ(iAlias).Interval = 0
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlatCtl" + ".acqq"
    Resume Next
End Function

Public Function PercentPlay(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0104"
    PercentPlay = PositionInSec(iAlias) / LengthInSec(iAlias) * 100
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqs"
    Resume Next
End Function

Public Function PositionInSec(iAlias As Long) As Long
    On Local Error GoTo ErrFunc
    Static s As String * 30
    
    dwReturn = mciSendString("status " + Alias(iAlias) + " position", s, Len(s), 0)
    tERR.Anotar "acqc", dwReturn, iAlias
    If dwReturn = 263 Then 'esta cerrado!!
        PositionInSec = -1
        Exit Function
    End If
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        tERR.AppendLog "NPS:" + CStr(dwReturn), "acqa"
    End If
    'esta funcion anda joya!!!
    PositionInSec = CLng(SoloNumeros(s)) / 1000
    Exit Function
ErrFunc:
    On Error GoTo ErrFunc2
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acqd"
    'el error puede ser que el primer caracter de S no sea valido
    PositionInSec = CLng(SoloNumeros(s)) / 1000
    Exit Function
ErrFunc2:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acqe"
    Resume Next
End Function

Public Function Position(iAlias As Long) As String
    On Local Error GoTo ErrFunc
    Static s As String * 30
    tERR.Anotar "002-0113", iAlias
    dwReturn = mciSendString("status " + Alias(iAlias) + " position", s, Len(s), 0)
    tERR.Anotar "002-0114"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.Position." + m_FileName(iAlias), ".acqt"
    End If
    tERR.Anotar "002-0117"
    SEC = Int(Mid$(s, 1, Len(s)) / 1000)
    tERR.Anotar "002-0118"
    If SEC < 60 Then Position = "0:" & Format(SEC, "00")
    tERR.Anotar "002-0119"
    If SEC > 59 Then
        tERR.Anotar "002-0120"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0121"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0122"
        Position = Format(MINS, "00") & ":" & Format(SEC, "00")
    End If
    
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqt"
    'el error puede ser que el primer caracter de S no sea valido
    Position = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqu"
    Resume Next
End Function

Public Function FaltaInSec(iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0129"
    Static s As String * 30
    tERR.Anotar "002-0130"
    FaltaInSec = LengthInSec(iAlias) - PositionInSec(iAlias) 'llamo a la funcion para que se manejen los errores desde ahi
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqu"
    Resume Next
End Function

Public Function Falta(iAlias As Long) As String
    On Error GoTo ERmp3
    Dim MINS As Long, SEC As Long
    tERR.Anotar "002-0131"
    SEC = FaltaInSec(iAlias)
    tERR.Anotar "002-0132", SEC
    If SEC < 60 Then Falta = "0:" & Format(SEC, "00")
    tERR.Anotar "002-0133", SEC
    If SEC > 59 Then
        tERR.Anotar "002-0134"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0135"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0136"
        Falta = Format(MINS, "00") & ":" & Format(SEC, "00")
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqv"
    Resume Next
End Function

Public Function LengthInSec(iAlias As Long) As Long
    tERR.Anotar "002-0137"
    
    On Local Error GoTo ErrFunc
    Static s As String * 30
    tERR.Anotar "002-0138b", iAlias, m_FileName(iAlias)
    dwReturn = mciSendString("status " + Alias(iAlias) + " length", s, Len(s), 0)
    If dwReturn <> 0 And dwReturn <> 263 Then
        tERR.AppendLog "138b400", "acqw"
        LogErrorMCI dwReturn
    End If
    tERR.Anotar "002-0141"
    'esta funcion anda joya!!!!
    LengthInSec = CLng(SoloNumeros(s)) / 1000
    'esto era una porqueria!!!!
    'LengthInSec = Int(Trim(Mid$(s, 1, Len(s)) / 1000))
    Exit Function
ErrFunc:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.Prueba2" + ".acqx"
    Resume Next
End Function

Public Function Length(iAlias As Long) As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0148"
    SEC = LengthInSec(iAlias) 'pateo posibles errores a LengthInSec
    tERR.Anotar "002-0149"
    If SEC < 60 Then Length = "0:" & Format(SEC, "00")
    tERR.Anotar "002-0150"
    If SEC > 59 Then
        tERR.Anotar "002-0151"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0152"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0153"
        Length = Format(MINS, "00") & ":" & Format(SEC, "00")
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqx"
    Resume Next
End Function

Public Function SeekTo(Second, iAlias As Long)
    On Error GoTo ERmp3
    tERR.Anotar "002-0154", iAlias
    If IsPlaying(iAlias) = True Then
        tERR.Anotar "002-0155"
        dwReturn = mciSendString("play " + Alias(iAlias) + " from " & Second, 0, 0, 0)
        tERR.Anotar "002-0156"
        If dwReturn <> 0 And dwReturn <> 282 Then '282 es que pide un lugar de tiempo que no existe!
            tERR.Anotar "002-0157"
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.Anotar "002-0158"
            tERR.AppendLog "MpPlayCtl.SeekTo.Open." + m_FileName(iAlias) + ".acqy"
        End If
    Else
        tERR.Anotar "002-0159"
        dwReturn = mciSendString("seek " + Alias(iAlias) + " to " & Second, 0, 0, 0)
        If dwReturn <> 0 And dwReturn <> 263 And dwReturn <> 282 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlayCtl.SeekTo.Close." + m_FileName(iAlias) + ".acqz"
        End If
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acra"
    Resume Next
End Function

Function Record() 'no se como hara con los alias, estimo que graba todo ¿¿??
    On Error GoTo ERmp3
    tERR.Anotar "002-0162"
    dwReturn = mciSendString("Close MP3rec", 0, 0, 0)
    tERR.Anotar "002-0163"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        tERR.Anotar "002-0164"
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.Anotar "002-0165"
        tERR.AppendLog "MpPlayCtl.acrb"
    End If
    tERR.Anotar "002-0166"
    Dim cmdToDo As String * 255
    tERR.Anotar "002-0167"
    'abrir nuevo
    tERR.Anotar "002-0168"
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    tERR.Anotar "002-0169"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrc"
        Exit Function
    End If
    'iniciar grabacion
    tERR.Anotar "002-0174"
    cmdToDo = "record MP3rec"
    tERR.Anotar "002-0175"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrd"
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acre"
    Resume Next
End Function

Function StopRecord()
    On Error GoTo ERmp3
    Dim cmdToDo As String * 255
    tERR.Anotar "002-0178"
    'parar nuevo
    cmdToDo = "stop MP3rec"
    tERR.Anotar "002-0179"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0180"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrf"
    End If
    'grabar grabacion
    tERR.Anotar "002-0182"
    cmdToDo = "save MP3rec c:\3pm.wav"
    tERR.Anotar "002-0183"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0184"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrg"
    End If
    
    'cerrra grabacion
    tERR.Anotar "002-0185"
    cmdToDo = "Close MP3rec "
    tERR.Anotar "002-0186"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0187"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.acrh"
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acrh"
    Resume Next
End Function

Public Sub LogErrorMCI(CodeErrMCI)
    On Error GoTo ERmp3
    
    Dim Buffer As String, Largo As Integer
    Buffer = Space$(512)
    
    Largo = mciGetErrorString(CodeErrMCI, Buffer, Len(Buffer))
    
    Dim ErrTEXT As String
    
    ErrTEXT = Left(Buffer, Len(Buffer))
    'en este writelog pongo la fecha y hora
    tERR.AppendLog "MciErr:" + Trim(CStr(CodeErrMCI)), ErrTEXT
    Exit Sub
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlay.acri"
    Resume Next
End Sub

'Public Sub WriteLog(TXT As String, PonerFecha As Boolean)
'
'    TXT = "Linea: " + LineaError + vbCrLf + TXT
'    If FSO.FileExists(AP + "log.txt") = False Then
'        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
'        TE.Close
'    End If
'
'    'ver si no es demasiado grande
'    If FileLen(AP + "log.txt") > 100000 Then 'hasta 100 KB aguanto
'        'pasarlo a otro archivo y volver a vrearlo
'        If FSO.FileExists(AP + "OLDlog.txt") Then FSO.DeleteFile AP + "OLDlog.txt", True
'        FSO.MoveFile AP + "log.txt", AP + "OLDlog.txt"
'        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
'        TE.Close
'    End If
'    'finalmente escribir
'
'
'    Set TE = FSO.OpenTextFile(AP + "log.txt", ForAppending, False)
'    If PonerFecha Then
'        TE.WriteLine vbCrLf + Trim(Str(Date)) + " / " + Trim(Str(time)) + vbCrLf + TXT
'    Else
'        TE.WriteLine vbCrLf + TXT
'    End If
'    TE.Close
'End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    On Local Error GoTo ErrFunc
    tERR.Anotar "002-0192"
    QuickLargoDeTema = "N/S"
    '------------cerrar si estaba abierto--------------
    tERR.Anotar "002-0193"
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    tERR.Anotar "002-0194"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no hay abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "acrj." + CStr(dwReturn)
    End If
    '------------abrir--------------
    Dim cmdToDo As String * 255
    
    tERR.Anotar "002-0195"
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    tERR.Anotar "002-0196"
    If Dir(TemaQuick) = "" Then       '
        tERR.Anotar "002-0197"
        tERR.AppendLog "acrk." + CStr(dwReturn)
        Exit Function
    End If
    tERR.Anotar "002-0198"
    lenShort = GetShortPathName(TemaQuick, TMP, 255)
    tERR.Anotar "002-0199"
    FileNameSHORT = Left$(TMP, lenShort)
    tERR.Anotar "002-0200"
    glo_hWnd = hwnd
    tERR.Anotar "002-0201"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3quick"
    tERR.Anotar "002-0202"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0203"
    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
        LogErrorMCI dwReturn
        tERR.AppendLog "acrk." + CStr(dwReturn)
        'En este caso queda tildado y no puede volver a mostrar hasta que se
        'cierre el MCI original (el que reproduce)
        Exit Function
    End If
    tERR.Anotar "002-0204"
    If dwReturn <> 0 And dwReturn <> 264 Then
        LogErrorMCI dwReturn
        tERR.AppendLog "acrl." + CStr(dwReturn)
        Exit Function
    End If
    tERR.Anotar "002-0205"
    '------------poner en milisegundos--------------
    dwReturn = mciSendString("set MP3quick time format milliseconds", 0, 0, 0)
    tERR.Anotar "002-0206"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        tERR.AppendLog "acrl." + CStr(dwReturn) + "." + TemaQuick
    End If
    '------------ver el largo--------------
    tERR.Anotar "002-0207"
    Static s As String * 30
    tERR.Anotar "002-0208"
    dwReturn = mciSendString("status MP3quick length", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        tERR.AppendLog "acrm." + CStr(dwReturn)
    End If
    tERR.Anotar "002-0209"
    SEC = CLng(SoloNumeros(s)) / 1000
    'sec = Int(Mid$(s, 1, Len(s)) / 1000)
    tERR.Anotar "002-0210"
    If SEC < 60 Then QuickLargoDeTema = "00:" & Format(SEC, "00")
    tERR.Anotar "002-0211"
    If SEC > 59 Then
        tERR.Anotar "002-0212"
        MINS = Int(SEC / 60)
        tERR.Anotar "002-0213"
        SEC = SEC - (MINS * 60)
        tERR.Anotar "002-0214"
        QuickLargoDeTema = Format(MINS, "00") & ":" & Format(SEC, "00")
        tERR.Anotar "002-0215"
    End If
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    tERR.AppendLog "acrn." + TemaQuick
    'el error puede ser que el primer caracter de S no sea valido
    LargoQuick = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    tERR.AppendLog "acro." + TemaQuick
    'Dim Ret As String * 128
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "acrp." + TemaQuick
    End If
    
End Function

Private Function SoloNumeros(TXT As String) As String
    Dim Largo As Long
    Largo = Len(TXT)
    Dim TmpNumber As String
    TmpNumber = ""
    Dim Letra As String
    For A = 1 To Largo
        Letra = Mid(TXT, A, 1)
        If IsNumeric(Letra) Then
            TmpNumber = TmpNumber + Letra
        End If
    Next
    If TmpNumber = "" Then TmpNumber = "0"
    SoloNumeros = TmpNumber
End Function




