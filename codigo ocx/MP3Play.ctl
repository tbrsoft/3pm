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
    CaminoError "002-0001"
    'primero ver si ermina el tema
    If IsPlaying = False Then
        CaminoError "002-0002"
        RELOJ.Interval = 0
        CaminoError "002-0003"
        RaiseEvent EndPlay
        CaminoError "002-0004"
        
        Exit Sub 'ESTO NO ESTABA!!!!!!!, seguia mandando el evento!!!!!!!!!!
    End If
    'y SOLO si no termino largar el evento. Antes estaba alreves!!!!!!!!!
    CaminoError "002-0005"
    RaiseEvent Played(PositionInSec)
    Exit Sub
ERmp3:
    WriteLog "Reloj - timer del MP3PLAY" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Sub

Private Sub UserControl_Resize()
    On Error GoTo ERmp3
    CaminoError "002-0006"
    UserControl.Height = 1620
    CaminoError "002-0007"
    UserControl.Width = 1500
Exit Sub
ERmp3:
    WriteLog "UserContro, Resize-Mp3Play" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    On Error GoTo ERmp3
    CaminoError "002-0008"
    Enabled = UserControl.Enabled
    Exit Property
ERmp3:
    WriteLog "Enabled" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    On Error GoTo ERmp3
    CaminoError "002-0009"
    UserControl.Enabled() = New_Enabled
    CaminoError "002-0010"
    PropertyChanged "Enabled"
    Exit Property
ERmp3:
    WriteLog "Enabled" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

Public Property Get Volumen() As Long
    On Error GoTo ERmp3
    CaminoError "002-0011"
    m_Volumen = m_Volumen / 10
    CaminoError "002-0012"
    Volumen = m_Volumen
    Exit Property
ERmp3:
    WriteLog "Enabled" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

Public Property Let Volumen(ByVal New_Volumen As Long)
    On Error GoTo ERmp3
    'en mi máquina anda del 0 al 1000
    CaminoError "002-0013"
    m_Volumen = New_Volumen * 10 ' * 30 - 3000
    CaminoError "002-0014"
    Ret = mciSendString("SetAudio MP3Play Volume To " + CStr(m_Volumen), 0, 0, 0)
    CaminoError "002-0015"
    If Ret <> 0 Then
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner el volumen en " + CStr(New_Volumen) + ". Tema: " + m_FileName + ". Property Let Volume", False
    End If
    CaminoError "002-0018"
    PropertyChanged "Volumen"
    Exit Property
ERmp3:
    WriteLog "Enabled" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get FileName() As String
    On Error GoTo ERmp3
    'CaminoError "002-0019"
    FileName = m_FileName
    Exit Property
ERmp3:
    WriteLog "Enabled" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

Public Property Let FileName(ByVal New_FileName As String)
    On Error GoTo ERmp3
    CaminoError "002-0020"
    m_FileName = New_FileName
    CaminoError "002-0021"
    PropertyChanged "FileName"
    Exit Property
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    On Error GoTo ERmp3
    CaminoError "002-0022"
    m_FileName = m_def_FileName
    'Visible = True
    Exit Sub
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error GoTo ERmp3
    CaminoError "002-0023"
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    CaminoError "002-0024"
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    Exit Sub
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error GoTo ERmp3
    CaminoError "002-0025"
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    CaminoError "002-0026"
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
    Exit Sub
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
End Sub

Public Function IsPlaying() As Boolean
    On Error GoTo ERmp3
    CaminoError "002-0027"
    If m_FileName = "" Then
        CaminoError "002-0028"
        IsPlaying = False
    Else
        CaminoError "002-0029"
        Static s As String * 30
        CaminoError "002-0030"
        Ret = mciSendString("status MP3Play mode", s, Len(s), 0)
        CaminoError "002-0031"
        If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
            LogErrorMCI Ret
            'no se pudo modificar el volumen
            WriteLog "No se pudo definir el estado de ejecucion." + ". Tema: " + m_FileName + " Function IsPlaying", False
        End If
        CaminoError "002-0034"
        IsPlaying = (Mid$(s, 1, 7) = "playing")
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoOpen()
    'DoStop    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    On Error GoTo ERmp3
    CaminoError "002-0035"
    Dim Ret As String * 128
    CaminoError "002-0036"
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    CaminoError "002-0037"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    CaminoError "002-0040"
    Dim cmdToDo As String * 255
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    CaminoError "002-0041"
    If FSO.FileExists(m_FileName) = False Then        '
        CaminoError "002-0042"
        WriteLog "No existe el archivo mp3 que se intenta abrir." + m_FileName + " Function DoOpen", True
        CaminoError "002-0043"
        Exit Function
    End If
    CaminoError "002-0044"
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    CaminoError "002-0045"
    FileNameSHORT = Left$(TMP, lenShort)
    CaminoError "002-0046"
    glo_hWnd = hwnd
    CaminoError "002-0047"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style child"
    CaminoError "002-0048"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0049"
    If dwReturn <> 0 Then
        If dwReturn = 263 Then
            'si da el error 263 es probable que la máquina no tenga MCI, lo que le paso a Mauro con W98 PE y a efren con ME
            WriteLog "WINDOWS NO REPRODUCE MP3!!!. INSTALE EL REPRODUCTOR CORESPONDIENTE " + _
                "A SU VERSION DE WINDOWS", True
        
            MsgBox "No se ha podido abrir el fichero debido a un problema existente en Windows. " + vbCrLf + _
                "Revise que el reproductor multimedia de Windows este instalado y funcione correctamente." + _
                "Notifique a tbrSoft de esto para más detalles"
                
        End If
        
        LogErrorMCI dwReturn
        'no se puedo abrir!!!
        WriteLog "No se pudo abrir un fichero mp3." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    CaminoError "002-0054"
    'uso todo en milisegundos
    dwReturn = mciSendString("set MP3Play time format milliseconds", 0, 0, 0)
    CaminoError "002-0055"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer el formato a milisegundos." + ". Tema: " + m_FileName + " Function DoOpen", False
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoOpenVideo(Style As String, HWind As Long, _
    X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
    
    On Error GoTo ERmp3
    'DoStop
    'DoClose
    'si uso esos dos mando el evento endPlay y se arma un kilombo
    CaminoError "002-0058"
    Dim Ret As String * 128
    CaminoError "002-0059"
    Ret = mciSendString("Close MP3Play", 0, 0, 0)
    CaminoError "002-0060"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        'no se pudo modificar el volumen
        WriteLog "No se pudo cerrar MCI (video) para reabrir tema." + ". Tema: " + m_FileName + "Function DoOpenVideo", False
    End If
    CaminoError "002-0063"
    Dim cmdToDo As String * 255

    CaminoError "002-0064"
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    CaminoError "002-0065"
    If Dir(m_FileName) = "" Then
        WriteLog "No existe el archivo de video que se intenta abrir." + m_FileName + " Function DoOpenVideo", True
        Exit Function
    End If
    CaminoError "002-0068"
    lenShort = GetShortPathName(m_FileName, TMP, 255)
    'la funcion transforma todo a 8.3 por que con espacioes
    'el reproductor no anda. JOYA JOYA JOYA
    CaminoError "002-0069"
    'volu = mciGetDeviceID(lenShort)
    FileNameSHORT = Left$(TMP, lenShort)
    CaminoError "002-0070"
    glo_hWnd = hwnd
    CaminoError "002-0071"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3Play style " + Style + " parent " + CStr(HWind)
    CaminoError "002-0072"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0073"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo abrir un fichero de video." + ". Tema: " + m_FileName + " Function DoOpenVideo", False
    End If
    CaminoError "002-0076"
    cmdToDo = "put MP3Play window at " + CStr(X1) + " " + CStr(Y1) + " " + CStr(X2) + " " + CStr(Y2) + " "
    CaminoError "002-0077"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner el video en la posicion y el tamaño especificado." + ". Tema: " + m_FileName + " Function DoOpenVideo (put MP3Play window at)", False
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
    CaminoError "002-0081"
    mciSendString "set MP3Play time format milliseconds", 0, 0, 0
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoPlay(Optional FullScreen As Boolean = False)
    On Error GoTo ERmp3
    CaminoError "002-0082"
    If FullScreen Then
        dwReturn = mciSendString("play MP3Play fullscreen", 0, 0, 0)
    Else
        dwReturn = mciSendString("play MP3Play", 0, 0, 0)
    End If
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo ejecutar un fichero." + m_FileName + " Function DoPlay", False
    End If
    CaminoError "002-0086"
    RELOJ.Interval = 1000
    CaminoError "002-0087"
    RaiseEvent BeginPlay
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoPause()
    On Error GoTo ERmp3
    CaminoError "002-0088"
    dwReturn = mciSendString("pause MP3Play", 0, 0, 0)
    CaminoError "002-0089"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo poner en pausa un fichero." + m_FileName + " Function DoPause", False
    End If
    CaminoError "002-0092"
    RELOJ.Interval = 0
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoStop() As String
    On Error GoTo ERmp3
    CaminoError "002-0093"
    dwReturn = mciSendString("stop MP3Play", 0, 0, 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo parar un fichero." + m_FileName + " Function DoStop", False
    End If
    CaminoError "002-0097"
    RELOJ.Interval = 0
    CaminoError "002-0098"
    RaiseEvent EndPlay
    
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function DoClose() As String
    On Error GoTo ERmp3
    CaminoError "002-0099"
    dwReturn = mciSendString("close MP3Play", 0, 0, 0)
    If dwReturn <> 0 And dwReturn <> 263 Then '263 ES CUANDO NO HAY NADA ABIERTO
        LogErrorMCI dwReturn
        WriteLog "No se pudo cerrar MCI." + ". Tema: " + m_FileName + " Function DoClose", False
    End If
    'SI SIGUE EL RELOJ SE MARCAN 1000 errores!!!!!!!!!!
    CaminoError "002-0103"
    RELOJ.Interval = 0
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function PercentPlay()
    On Error GoTo ERmp3
    CaminoError "002-0104"
    PercentPlay = PositionInSec / LengthInSec * 100
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function PositionInSec()
    CaminoError "002-0105"
    On Local Error GoTo ErrFunc
    Static s As String * 30
    CaminoError "002-0106"
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la posicion." + m_FileName + " Function PositionInSec", False
    End If
    CaminoError "002-0109"
    'esta funcion anda joya!!!
    PositionInSec = CLng(SoloNumeros(s)) / 1000
    'porqueria
    'PositionInSec = Int(Mid$(s, 1, Len(s)) / 1000)
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para PositionInSec no es válido." + ". Tema: " + m_FileName + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    PositionInSec = Int(Mid$(s, 2, Len(s)) / 1000)
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para PositionInSec no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    Resume Next
End Function

Public Function Position()
    On Local Error GoTo ErrFunc
    Static s As String * 30
    CaminoError "002-0113"
    dwReturn = mciSendString("status MP3Play position", s, Len(s), 0)
    CaminoError "002-0114"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la posicion." + m_FileName + " Function Position", False
    End If
    CaminoError "002-0117"
    sec = Int(Mid$(s, 1, Len(s)) / 1000)
    CaminoError "002-0118"
    If sec < 60 Then Position = "0:" & Format(sec, "00")
    CaminoError "002-0119"
    If sec > 59 Then
        CaminoError "002-0120"
        mins = Int(sec / 60)
        CaminoError "002-0121"
        sec = sec - (mins * 60)
        CaminoError "002-0122"
        Position = Format(mins, "00") & ":" & Format(sec, "00")
    End If
    
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para PositionInSec no es válido." + ". Tema: " + m_FileName + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    Position = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para Position no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    Resume Next
End Function

Public Function FaltaInSec()
    On Error GoTo ERmp3
    CaminoError "002-0129"
    Static s As String * 30
    CaminoError "002-0130"
    FaltaInSec = LengthInSec - PositionInSec 'llamo a la funcion para que se manejen los errores desde ahi
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function Falta()
    On Error GoTo ERmp3
    CaminoError "002-0131"
    sec = FaltaInSec
    CaminoError "002-0132"
    If sec < 60 Then Falta = "0:" & Format(sec, "00")
    CaminoError "002-0133"
    If sec > 59 Then
        CaminoError "002-0134"
        mins = Int(sec / 60)
        CaminoError "002-0135"
        sec = sec - (mins * 60)
        CaminoError "002-0136"
        Falta = Format(mins, "00") & ":" & Format(sec, "00")
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function LengthInSec()
    CaminoError "002-0137"
    On Local Error GoTo ErrFunc
    Static s As String * 30
    CaminoError "002-0138"
    dwReturn = mciSendString("status MP3Play length", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo establecer la duracion." + m_FileName + " Function LengthInSec", False
    End If
    CaminoError "002-0141"
    'esta funcion anda joya!!!!
    LengthInSec = CLng(SoloNumeros(s)) / 1000
    'esto era una porqueria!!!!
    'LengthInSec = Int(Trim(Mid$(s, 1, Len(s)) / 1000))
    Exit Function
ErrFunc:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI para Length no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    Resume Next
End Function

Public Function Length()
    On Error GoTo ERmp3
    CaminoError "002-0148"
    sec = LengthInSec 'pateo posibles errores a LengthInSec
    CaminoError "002-0149"
    If sec < 60 Then Length = "0:" & Format(sec, "00")
    CaminoError "002-0150"
    If sec > 59 Then
        CaminoError "002-0151"
        mins = Int(sec / 60)
        CaminoError "002-0152"
        sec = sec - (mins * 60)
        CaminoError "002-0153"
        Length = Format(mins, "00") & ":" & Format(sec, "00")
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Public Function SeekTo(Second)
    On Error GoTo ERmp3
    CaminoError "002-0154"
    If IsPlaying = True Then
        CaminoError "002-0155"
        dwReturn = mciSendString("play MP3Play from " & Second, 0, 0, 0)
        CaminoError "002-0156"
        If dwReturn <> 0 Then
            CaminoError "002-0157"
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            CaminoError "002-0158"
            WriteLog "No se pudo seek mientras se ejecutaba." + m_FileName + " Function SeekTo", False
        End If
    Else
        CaminoError "002-0159"
        dwReturn = mciSendString("seek MP3Play to " & Second, 0, 0, 0)
        If dwReturn <> 0 Then
            LogErrorMCI dwReturn
            'no se pudo modificar el volumen
            WriteLog "No se pudo seek mientras estaba detenida la ejecucion." + m_FileName + " Function SeekTo", False
        End If
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Function Record()
    On Error GoTo ERmp3
    CaminoError "002-0162"
    dwReturn = mciSendString("Close MP3rec", 0, 0, 0)
    CaminoError "002-0163"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        CaminoError "002-0164"
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        CaminoError "002-0165"
        WriteLog "No se pudo cerrar para comenzar a grabar." + m_FileName + " Function Record", False
    End If
    CaminoError "002-0166"
    Dim cmdToDo As String * 255
    CaminoError "002-0167"
    'abrir nuevo
    CaminoError "002-0168"
    cmdToDo = "open new type WaveAudio Alias MP3rec"
    CaminoError "002-0169"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo abrir un nuevo WaveAudio para comenzar a grabar." + m_FileName + " Function Record", False
        Exit Function
    End If
    'iniciar grabacion
    CaminoError "002-0174"
    cmdToDo = "record MP3rec"
    CaminoError "002-0175"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo comenzar a grabar." + m_FileName + " Function Record", False
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Function

Function StopRecord()
    On Error GoTo ERmp3
    Dim cmdToDo As String * 255
    CaminoError "002-0178"
    'parar nuevo
    cmdToDo = "stop MP3rec"
    CaminoError "002-0179"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0180"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo detener la grabacion." + m_FileName + " Function StopRecord", False
    End If
    'grabar grabacion
    CaminoError "002-0182"
    cmdToDo = "save MP3rec 3pm.wav"
    CaminoError "002-0183"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0184"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo grabar en archivo al detener la grabacion." + m_FileName + " Function StopRecord", False
    End If
    
    'cerrra grabacion
    CaminoError "002-0185"
    cmdToDo = "Close MP3rec "
    CaminoError "002-0186"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0187"
    If dwReturn <> 0 And dwReturn <> 263 Then '263 es cuando no hay nada abierto
        LogErrorMCI dwReturn
        'no se pudo modificar el volumen
        WriteLog "No se pudo cerrar MCI de grabacion ." + m_FileName + " Function StopRecord", False
    End If
    Exit Function
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
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
    WriteLog "Error MCI nº " + Trim(Str(CodeErrMCI)) + ": " + ErrTEXT, True
    Exit Sub
ERmp3:
    WriteLog "-" + vbCrLf + _
        "Desc: " + Err.Description + " (" + CStr(Err.Number) + ")", True
    Resume Next
End Sub

Public Sub WriteLog(TXT As String, PonerFecha As Boolean)
    
    TXT = "Linea: " + LineaError + vbCrLf + TXT
    If FSO.FileExists(AP + "log.txt") = False Then
        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
        TE.Close
    End If
    
    'ver si no es demasiado grande
    If FileLen(AP + "log.txt") > 100000 Then 'hasta 100 KB aguanto
        'pasarlo a otro archivo y volver a vrearlo
        If FSO.FileExists(AP + "OLDlog.txt") Then FSO.DeleteFile AP + "OLDlog.txt", True
        FSO.MoveFile AP + "log.txt", AP + "OLDlog.txt"
        Set TE = FSO.CreateTextFile(AP + "log.txt", False)
        TE.Close
    End If
    'finalmente escribir
    
    
    Set TE = FSO.OpenTextFile(AP + "log.txt", ForAppending, False)
    If PonerFecha Then
        TE.WriteLine Trim(Str(Date)) + " / " + Trim(Str(time)) + vbCrLf + TXT
    Else
        TE.WriteLine TXT
    End If
    TE.Close
End Sub

Public Function QuickLargoDeTema(TemaQuick As String) As String
    On Local Error GoTo ErrFunc
    CaminoError "002-0192"
    QuickLargoDeTema = "N/S"
    '------------cerrar si estaba abierto--------------
    CaminoError "002-0193"
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    CaminoError "002-0194"
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no hay abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema MP3 quick." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
    End If
    '------------abrir--------------
    Dim cmdToDo As String * 255
    
    CaminoError "002-0195"
    Dim TMP As String * 255
    Dim lenShort As Long
    Dim FileNameSHORT As String
    CaminoError "002-0196"
    If Dir(TemaQuick) = "" Then       '
        CaminoError "002-0197"
        WriteLog "No existe el archivo mp3 que se intenta abrir (QUICK)." + TemaQuick + " Function QuickLargoDeTema", True
        Exit Function
    End If
    CaminoError "002-0198"
    lenShort = GetShortPathName(TemaQuick, TMP, 255)
    CaminoError "002-0199"
    FileNameSHORT = Left$(TMP, lenShort)
    CaminoError "002-0200"
    glo_hWnd = hwnd
    CaminoError "002-0201"
    cmdToDo = "open " & FileNameSHORT & " type MPEGVideo Alias MP3quick"
    CaminoError "002-0202"
    dwReturn = mciSendString(cmdToDo, 0&, 0&, 0&)
    CaminoError "002-0203"
    If dwReturn = 264 Then 'no hay memoria sufuciente!!!
        LogErrorMCI dwReturn
        WriteLog "No se pudo abrir MCI (QUICK) ." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
        'En este caso queda tildado y no puede volver a mostrar hasta que se
        'cierre el MCI original (el que reproduce)
        Exit Function
    End If
    CaminoError "002-0204"
    If dwReturn <> 0 And dwReturn <> 264 Then
        LogErrorMCI dwReturn
        WriteLog "No se pudo abrir un fichero mp3 (QUICK)." + ". Tema: " + TemaQuick + " Function QuickLargoDeTema", False
        Exit Function
    End If
    CaminoError "002-0205"
    '------------poner en milisegundos--------------
    dwReturn = mciSendString("set MP3quick time format milliseconds", 0, 0, 0)
    CaminoError "002-0206"
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        WriteLog "No se pudo establecer el formato a milisegundos." + ". Tema: " + temacuick + " Function QuickLargoDeTema", False
    End If
    '------------ver el largo--------------
    CaminoError "002-0207"
    Static s As String * 30
    CaminoError "002-0208"
    dwReturn = mciSendString("status MP3quick length", s, Len(s), 0)
    If dwReturn <> 0 Then
        LogErrorMCI dwReturn
        WriteLog "No se pudo establecer la duracion." + TemaQuick + " Function QuickLargoDeTema", False
    End If
    CaminoError "002-0209"
    sec = CLng(SoloNumeros(s)) / 1000
    'sec = Int(Mid$(s, 1, Len(s)) / 1000)
    CaminoError "002-0210"
    If sec < 60 Then QuickLargoDeTema = "00:" & Format(sec, "00")
    CaminoError "002-0211"
    If sec > 59 Then
        CaminoError "002-0212"
        mins = Int(sec / 60)
        CaminoError "002-0213"
        sec = sec - (mins * 60)
        CaminoError "002-0214"
        QuickLargoDeTema = Format(mins, "00") & ":" & Format(sec, "00")
        CaminoError "002-0215"
    End If
    Exit Function
ErrFunc:
    On Local Error GoTo ErrFunc2
    WriteLog "El Valor devuelto por MCI para Length (QUICK) no es válido." + ". Tema: " + TemaQuick + " Valor= " + s, True
    'el error puede ser que el primer caracter de S no sea valido
    LargoQuick = Int(Mid$(s, 2, Len(s)) / 1000)
    Resume 'volver a ver que pasa
    Exit Function
ErrFunc2:
    WriteLog "El Valor devuelto (preuba desde el segundo caracter) por MCI (QUICK) para Length no es válido. Valor 2° prueba = " + Mid$(s, 2, Len(s)), False
    'Dim Ret As String * 128
    Ret = mciSendString("Close MP3quick", 0, 0, 0)
    If Ret <> 0 And Ret <> 263 Then '263 es cuando no ha abierto nada
        LogErrorMCI Ret
        WriteLog "No se pudo cerrar MCI para reabrir tema." + ". Tema: " + m_FileName + " Function DoOpen", False
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




