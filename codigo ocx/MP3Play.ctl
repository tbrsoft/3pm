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
      Left            =   570
      Top             =   570
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
'Default Property Values:
Const m_def_FileName = ""
'Property Variables:
Dim m_FileName As String
Dim m_Volumen As Long
Dim dwReturn As Long

Event Played(SecondsPlayed As Long)
Event BeginPlay()
Event EndPlay()

Private Sub Reloj_Timer()
    On Error GoTo ERmp3
    tERR.Anotar "002-0001"
    'primero ver si ermina el tema
    If IsPlaying = False Then
        RELOJ.Interval = 0
        RaiseEvent EndPlay
        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
    End If
    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
    tERR.Anotar "002-0005"
    RaiseEvent Played(PositionInSec)
    Exit Sub
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayctl" + ".acpw"
    Resume Next
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 1620
    UserControl.Width = 1500
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Volumen() As Long
    On Error GoTo ERmp3
    tERR.Anotar "002-0011"
    m_Volumen = m_Volumen / 10
    tERR.Anotar "002-0012"
    Volumen = m_Volumen
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpv"
End Property

Public Property Let Volumen(ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi m�quina anda del 0 al 1000
    tERR.Anotar "002-0013"
    m_Volumen = New_Volumen * 10 ' * 30 - 3000
    tERR.Anotar "002-0014"
    Ret = mciSendString("SetAudio MP3Play Volume To " + CStr(m_Volumen), 0, 0, 0)
    tERR.Anotar "002-0015"
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        tERR.AppendLog "NoVolumenEn:" + CStr(Ret), "MpPalyCtl" + ".acpw"
    End If
    tERR.Anotar "002-0018"
    PropertyChanged "Volumen"
    Exit Property
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPalyCtl" + ".acpx"
    Resume Next
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_FileName = m_def_FileName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo ERmp3
    tERR.Anotar "002-0025"
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    tERR.Anotar "002-0026"
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Exit Sub
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acpx"
End Sub

Public Function IsPlaying() As Boolean
    On Error GoTo ERmp3
    tERR.Anotar "002-0027"
    If m_FileName = "" Then
        tERR.Anotar "002-0028"
        IsPlaying = False
    Else
        tERR.Anotar "002-0029"
        Static s As String * 30
        tERR.Anotar "002-0030", HabilitarVUMetro, NoVumVID
        Ret = mciSendString("status MP3Play mode", s, Len(s), 0)
        tERR.Anotar "002-0031", Ret
        If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            tERR.AppendLog "ERR IsPlaying=Status:" + CStr(Ret)
            'WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        'EN ESTE CASO ES NULO O ALGO ASI
        'YA QUE MCI NO TIENE LA CAPACIDAD DE STATUS!!!
        If Ret = 274 Then
            tERR.Anotar "002-0034b"
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

Public Function DoOpen()
    'DoStop    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    On Error GoTo ERmp3
    tERR.Anotar "002-0035"
    Dim Ret As String * 128
    tERR.Anotar "002-0036"
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    tERR.Anotar "002-0037"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "NoCierraMCI.RET:" + CStr(Ret), m_FileName
    End If
    tERR.Anotar "002-0040"
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    tERR.Anotar "002-0041"
    If FSO.FileExists(m_FileName) = False Then        '
        tERR.Anotar "002-0042"
        tERR.AppendLog "MpPlayCtl.DoOpen.NoExistFile.acqb", m_FileName
        tERR.Anotar "002-0043"
        Exit Function
    End If
    tERR.Anotar "002-0044"
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0045", m_FileName
    FileNameSHORT = Left$(TMP, lenShort)
    tERR.Anotar "002-0046"
    glo_hWnd = hwnd
    tERR.Anotar "002-0047"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style child"
    tERR.Anotar "002-0048"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0049"
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la m�quina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            tERR.AppendLog "WINDOWS NO REPRODUCE MP3!!!. INSTALE EL REPRODUCTOR CORESPONDIENTE " + _
                "A SU VERSION DE WINDOWS"
            MsgBox "No se ha podido abrir el fichero debido a un problema existente en Windows. " + vbCrLf + _
                "Revise que el reproductor multimedia de Windows este instalado y funcione correctamente." + _
                "Notifique a tbrSoft de esto para m�s detalles"
                
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        tERR.AppendLog "DoOpen.NoAbre." + CStr(dwReturn), "acqe"
    End If
    tERR.Anotar "002-0054"
    'uso todo en milisegundos
    dwReturn = mciSendString("set MP3Play time format milliseconds", 0, 0, 0)
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
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    
    On Error GoTo ERmp3
    'DoStop
    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    tERR.Anotar "002-0058"
    Dim Ret As String * 128
    tERR.Anotar "002-0059"
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    tERR.Anotar "002-0060", Ret
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        tERR.AppendLog "DoOpenVid.acqg.NoCierra", "MpPlayCtl"
    End If
    tERR.Anotar "002-0063"
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    tERR.Anotar "002-0065"
    If Dir(m_FileName) = "" Then
        tERR.Anotar "002-0065b"
        tERR.AppendLog "NoExist.DoOpenVid.acqh", m_FileName
        Exit Function
    End If
    tERR.Anotar "002-0068"
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    tERR.Anotar "002-0069", m_FileName
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = Left$(TMP, lenShort)
    tERR.Anotar "002-0070", FileNameSHORT
    glo_hWnd = hwnd
    tERR.Anotar "002-0071", HabilitarVUMetro, NoVumVID, HWind
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style " + _
        Style + " parent " + CStr(HWind)
    tERR.Anotar "002-0072", cmdToDo '                         xxxxxx
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    tERR.Anotar "002-0073", dwReturn, Salida2, Style, HWind
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlay.DoOpenVid.acqh." + CStr(dwReturn), m_FileName
    End If
    tERR.Anotar "002-0076", X1, X2, Y1, Y2
    cmdToDo = "put MP3Play window at " + CStr(X1) + " " + CStr(Y1) + _
        " " + CStr(X2) + " " + CStr(Y2) + " "
    tERR.Anotar "002-0077", Salida2, Style, HWind
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    'lo uso para ver info de los videos que andan bien tambien
    'tERR.AppendLog "ADD77:" + CStr(dwReturn), m_FileName
    If dwReturn <> 0 Then
        '********************
        'si es 346 es que no tiene ventana de presentacion (base+90=MCIERR_NO_WINDOW)
        '�����????????
        'probe con style popup y overlapped y no sirven ni solucionan
        If dwReturn = 346 Then
            tERR.AppendLog "MCIERR_NO_WINDOW=346. No hay ventana de presentacion!!!" + CStr(dwReturn), m_FileName
            'pasa con videos con codecs nuevos que al cargarse si funciona en
            'WMP pero no en 3PM!!!!!!!!!!!!!!!!!!!
        '********************
        Else
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlay.DoOpenVid.WindowAt.acqi." + CStr(dwReturn), m_FileName
        End If
    End If
    
    'Dim s As String
    's = Space(30)
    'cmdToDo = "capability MP3Play can stretch"
    'dwReturn = mciSendString(cmdToDo, s, Len(s), 0)
    'If dwReturn <> 0 Then MsgBox "Error n�: " + Str(dwReturn)
   '
   ' s = Trim(Left(s, 4))
   ' If UCase(s) = "TRUE" Then
   '     cmdToDo = "window MP3Play stretch"
   '     dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
   '     If dwReturn <> 0 Then MsgBox "Error n�: " + Str(dwReturn)
   ' End If
        
    'uso todo en milisegundos
    tERR.Anotar "002-0081"
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqj"
    Resume Next
End Function

Public Function DoPlay(Optional FullScreen As Boolean = False)
    On Error GoTo ERmp3
    tERR.Anotar "002-0082", CStr(FullScreen)
    If FullScreen Then
        dwReturn = mciSendString("play MP3Play fullscreen", 0, 0, 0)
    Else
        dwReturn = mciSendString("play MP3Play", 0, 0, 0)
    End If
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPlay.Play." + m_FileName, ".acqk"
    End If
    tERR.Anotar "002-0086"
    RELOJ.Interval = 1000
    tERR.Anotar "002-0087"
    RaiseEvent BeginPlay
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acql"
    Resume Next
End Function

Public Function DoPause()
    On Error GoTo ERmp3
    tERR.Anotar "002-0088"
    dwReturn = mciSendString("pause MP3Play", 0, 0, 0)
    tERR.Anotar "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoPause." + m_FileName, ".acqm"
    End If
    tERR.Anotar "002-0092"
    RELOJ.Interval = 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqn"
    Resume Next
End Function

Public Function DoStop() As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0093"
    dwReturn = mciSendString("stop MP3Play", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.DoStop." + m_FileName, ".acqo"
    End If
    tERR.Anotar "002-0097"
    RELOJ.Interval = 0
    tERR.Anotar "002-0098"
    RaiseEvent EndPlay
    
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqp"
    Resume Next
End Function

Public Function DoClose() As String
    On Error GoTo ERmp3
    tERR.Anotar "002-0099"
    dwReturn = mciSendString("close MP3Play", 0, 0, 0)
    If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
        LogErrorMCI dwReturn
        tERR.AppendLog "MpPlayCtl.DoClose." + m_FileName, ".acqr"
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    tERR.Anotar "002-0103"
    RELOJ.Interval = 0
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlatCtl" + ".acqq"
    Resume Next
End Function

Public Function PercentPlay()
    On Error GoTo ERmp3
    tERR.Anotar "002-0104"
    PercentPlay = PositionInSec / LengthInSec * 100
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqs"
    Resume Next
End Function

Public Function PositionInSec()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    tERR.Anotar "acqc", dwReturn
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.NoPositionInSec.RET:" + CStr(dwReturn), "acqa"
    End If
    'esta funcion anda joya!!!
    PositionInSec = CLng(SoloNumeros(s)) / 1000
    Exit Function
ErrFunc:
    On Error GoTo ErrFunc2
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acqd"
    'el error puede ser que el primer caracter de S no sea valido
    PositionInSec = Int(Mid$(s, 2, Len(s)) / 1000)
    Exit Function
ErrFunc2:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl.acqe"
    Resume Next
End Function

Public Function Position()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    tERR.Anotar "002-0113"
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    tERR.Anotar "002-0114"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        tERR.AppendLog "MpPlayCtl.Position." + m_FileName, ".acqt"
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

Public Function FaltaInSec()
    On Error GoTo ERmp3
    tERR.Anotar "002-0129"
    Static s As String * 30
    tERR.Anotar "002-0130"
    FaltaInSec = LengthInSec - PositionInSec 'llamo a la funcion para que se manejen los errores desde ahi
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acqu"
    Resume Next
End Function

Public Function Falta() As String
    On Error GoTo ERmp3
    Dim MINS As Long, SEC As Long
    tERR.Anotar "002-0131"
    SEC = FaltaInSec
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

Public Function LengthInSec()
    tERR.Anotar "002-0137"
    
    On Local Error GoTo ErrFunc
    Static s As String * 30
    tERR.Anotar "002-0138b", m_FileName
    dwReturn = mciSendString("status MP3Play length", s, Len(s), 0)
    If dwReturn <> 0 Then
        tERR.Anotar "MpPlayCtl.LengthInSec.Len." + m_FileName, ".acqw"
        
        tERR.AppendLog "MpPlayCtl.LengthInSec.Len." + m_FileName, ".acqw"
        
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        
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

Public Function Length()
    On Error GoTo ERmp3
    tERR.Anotar "002-0148"
    SEC = LengthInSec 'pateo posibles errores a LengthInSec
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

Public Function SeekTo(Second)
    On Error GoTo ERmp3
    tERR.Anotar "002-0154"
    If IsPlaying = True Then
        tERR.Anotar "002-0155"
        dwReturn = mciSendString("play MP3Play from " & Second, 0, 0, 0)
        tERR.Anotar "002-0156"
        If dwReturn <> 0 Then
            tERR.Anotar "002-0157"
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.Anotar "002-0158"
            tERR.AppendLog "MpPlayCtl.SeekTo.Open." + m_FileName + ".acqy"
        End If
    Else
        tERR.Anotar "002-0159"
        dwReturn = mciSendString("seek MP3Play to " & Second, 0, 0, 0)
        If dwReturn <> 0 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            tERR.AppendLog "MpPlayCtl.SeekTo.Close." + m_FileName + ".acqz"
        End If
    End If
    Exit Function
ERmp3:
    tERR.AppendLog tERR.ErrToTXT(Err), "MpPlayCtl" + ".acra"
    Resume Next
End Function

Function Record()
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
    cmdToDo = "save MP3rec 3pm.wav"
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
        tERR.AppendLog "acrl." + CStr(dwReturn) + "." + m_FileName
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
    tERR.AppendLog "acrn." + m_FileName
    'el error puede ser que el primer caracter de S no sea valido
    LargoQuick = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    tERR.AppendLog "acro." + m_FileName
    'Dim Ret As String * 128
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        tERR.AppendLog "acrp." + m_FileName
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
    SoloNumeros = TmpNumber
End Function




