VERSION 5.00
Begin VB.UserControl tbrWEB 
   BackColor       =   &H00404000&
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8355
   ScaleHeight     =   4875
   ScaleWidth      =   8355
   Begin VB.ListBox lstBytes 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   5490
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   1860
      Width           =   1425
   End
   Begin VB.TextBox txtLOG 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   600
      Width           =   5325
   End
   Begin VB.ListBox lstFOLDERS 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   1
      Top             =   1860
      Width           =   2775
   End
   Begin VB.ListBox lstFILES 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2550
      IntegralHeight  =   0   'False
      Left            =   2880
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   1860
      Width           =   2595
   End
   Begin VB.Label lblFULLPATH 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ".."
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   4470
      Width           =   6855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   5550
      TabIndex        =   7
      Top             =   1590
      Width           =   855
   End
   Begin VB.Label lblINFO 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INFO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   30
      TabIndex        =   5
      Top             =   60
      Width           =   5235
   End
   Begin VB.Shape SH 
      BackColor       =   &H00FFFFC0&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   165
      Left            =   30
      Shape           =   4  'Rounded Rectangle
      Top             =   390
      Width           =   345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Carpetas"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   1560
      Width           =   2835
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Archivos"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2910
      TabIndex        =   2
      Top             =   1590
      Width           =   2685
   End
End
Attribute VB_Name = "tbrWEB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event FinfFile(sFile As String)
Public Event FinfFolder(sFolder As String)

'para todas las funciones devuelve 0 en OK
'los numeros siguintes son diferentes fallas

Dim hFile As Long
Const sReadBuffer = 1024

Dim KLIC As String 'path despues de ftp.ftpserver.com
'si es solo la barra / es en la raiz. siempre termina en /

Dim SESSION As Long
Dim SERVER As Long 'handle de la conexion (¿?)
Dim ADR As String 'aparentemente carpeta actual

Public strPath As String

'lo hago publico para que se pueda agregar detalle a esta barra de progreso desde los progrmas que lo uan
Public Sub LOG(TXT As String, ShapeWidth As Single)
    txtLOG = txtLOG + CStr(Time) + "->" + TXT + vbCrLf
    'mostrar lo ultimo
    txtLOG.SelStart = Len(txtLOG) - 1
    txtLOG.SelLength = 1
    SH.Width = lblINFO.Width * ShapeWidth
    SH.Refresh
    lblINFO = TXT
    lblINFO.Refresh
    UserControl.Refresh
End Sub

Public Function Connect(ADRESS As String, ID As String, PSW As String, PORT As Integer, _
    Optional TransBinary As Boolean = True, Optional ftpPASSIVE As Boolean = True) As Long
    
    LOG "Intenta conectar", 0.1
    
    KLIC = ""
    Dim TRANSFER As Long 'tipo de transferencia
    If TransBinary Then
        TRANSFER = FTP_TRANSFER_TYPE_BINARY
    Else
        TRANSFER = FTP_TRANSFER_TYPE_ASCII
    End If
    
    Dim Service As Long 'tipo de servicio
    If ftpPASSIVE Then
        Service = INTERNET_FLAG_PASSIVE
    Else
        Service = INTERNET_FLAG_EXISTING_CONNECT
    End If
    LOG "Intenta Abrir sesion", 0.6
    SESSION = InternetOpen("SiteName", INTERNET_OPEN_TYPE_DIRECT, "", "", INTERNET_FLAG_NO_CACHE_WRITE)
    
    If SESSION <> 0 Then
        LOG "Sesion OK=" + CStr(SESSION), 0.9
        
        SERVER = InternetConnect(SESSION, ADRESS, PORT, ID, PSW, _
            INTERNET_SERVICE_FTP, Service, &H0)
        
        If SERVER = 0 Then
            
           LOG "Server FALLO!=", 0
           InternetCloseHandle SESSION
           
           Connect = 2 'falla server
           Exit Function
        Else
            LOG "Server OK=" + CStr(SERVER), 1
            ADR = Space(260)
            FtpGetCurrentDirectory SERVER, ADR, Len(ADR)
    
            ADR = Left(ADR, InStr(1, ADR, Chr(0)) - 1)
            ADR = ADR & IIf((Right(ADR, 1) = "\"), "*.*", "\*.*")
            LOG "CONECTADO OK ADR=" + ADR, 0.1
            KLIC = "\"
        End If
    Else
        Connect = 1 'falla sesion
        LOG "Fallo conexion, cerrando", 0
        InternetCloseHandle SESSION
        Exit Function
    End If
    Connect = 0 'todo OK

End Function

Public Function List(Optional sFILTER As String = "*.*") As Long
    
    LOG "Listando...", 0.1
    Dim hFile2 As Long
    Dim udtWFD As WIN32_FIND_DATA
    Dim strFile As String
    Dim Img As Integer, r As Integer
    Dim L&
    Dim sTime As SYSTEMTIME, lTime As FILETIME
    
    If SESSION = 0 Or SERVER = 0 Then
        LOG "Lista falla por no conexion", 0
        List = 1 'no esta conectado
        Exit Function
    End If
    
    'ver si es con filtro!!!
    If sFILTER <> "*.*" Then ADR = Left(ADR, Len(ADR) - 3) + sFILTER
    
    lstBytes.Clear: lstFILES.Clear: lstFOLDERS.Clear
    hFile2 = FtpFindFirstFile(SERVER, ADR, udtWFD, INTERNET_FLAG_RELOAD Or INTERNET_FLAG_NO_CACHE_WRITE, 0&)
    If hFile2 Then
        Do
            strFile = Left(udtWFD.cFileName, InStr(1, udtWFD.cFileName, Chr(0)) - 1)
            If Len(strFile) > 0 Then
                If udtWFD.dwFileAttributes And vbDirectory Then
                    'ES CARPETA = strfile
                    If strFile <> "." Then
                        lstFOLDERS.AddItem strFile
                        'como no se cuantos archivos va a haber pongo como si fueran 15
                        LOG "Carpeta encontrada: " + strFile, (lstFOLDERS.ListCount + lstFILES.ListCount) / 15
                        RaiseEvent FinfFolder(strFile)
                    End If
                Else
                    'ES ACHIVO = strfile
                    'tengo el tamaño = Format((udtWFD.nFileSizeLow / 1024), "### ### ###.##") & "Kb"
                    'tengo time  lTime = udtWFD.ftLastWriteTime
                    'tengo = FileTimeToSystemTime(lTime, sTime)
                    'lsItem.SubItems(2) = CalcFTime(sTime)
                    'lsItem.SubItems(3) = udtWFD.nFileSizeLow
                    lstFILES.AddItem strFile
                    lstBytes.AddItem CStr(udtWFD.nFileSizeLow)
                    'elegir todos de manera predeterminada
                    lstFILES.Selected(lstFILES.ListCount - 1) = True
                    LOG "Archivo encontrado: " + strFile, (lstFOLDERS.ListCount + lstFILES.ListCount) / 15
                    RaiseEvent FinfFile(strFile)
                End If
            End If
        Loop While InternetFindNextFile(hFile2, udtWFD)
    End If
    InternetCloseHandle hFile2
    LOG "Se listaron todos los datos OK", 1
    List = 0 'todo ok
End Function

Public Function GetIndexListFilesName(Index As Long) As String
    GetIndexListFilesName = lstFILES.List(Index)
End Function

Public Function GetIndexListFilesBytes(Index As Long) As Long
    GetIndexListFilesBytes = CLng(lstBytes.List(Index))
End Function

Public Function GetIndexMaxFiles() As Long
    GetIndexMaxFiles = lstFILES.ListCount
End Function

Public Function ExisteFolder(sFolder As String) As Boolean
    'busca si existe la carpeta especificada en el directorio actual
    Dim A As Long
    ExisteFolder = False
    For A = 0 To lstFOLDERS.ListCount - 1
        If sFolder = lstFOLDERS.List(A) Then
            ExisteFolder = True
            Exit For
        End If
    Next A
End Function

Public Function ExisteFile(sFile As String) As Boolean
    'busca si existe la carpeta especificada en el directorio actual
    Dim A As Long
    ExisteFile = False
    For A = 0 To lstFILES.ListCount - 1
        If sFile = lstFILES.List(A) Then
            ExisteFile = True
            Exit For
        End If
    Next A
End Function


Public Function UpLoad(ListaFiles() As String, LocalePath As String) As Long
    
    'listafiles no debe tener el path completo en sus elementos!!
    'localepath tiene la carpeta que los contiene con la barra al final
    'el destino es la carpeta actual!
        
    If Right(LocalePath, 1) <> "\" Then LocalePath = LocalePath + "\"
    
    'LEER INFO DE LO QUE SE VA A SUBIR
    Dim TotBytes As Long 'total a subir
    Dim LenFileBytes As Long 'total de cada archivo
    
    TotBytes = 0
    
    Dim i As Integer
    
    If SESSION = 0 Or SERVER = 0 Then
        LOG "No upload, no Server", 0
        UpLoad = 1 'no hay coneccion
        Exit Function
    End If
    
    LOG "Leyendo Info Upload", 0.2
    For i = 0 To UBound(ListaFiles)
        LenFileBytes = FileLen(LocalePath + ListaFiles(i))
        TotBytes = TotBytes + LenFileBytes
        LOG "Leyendo Info Upload " + CStr(i + 1) + "/" + CStr(UBound(ListaFiles) + 1), (i + 1) / (UBound(ListaFiles) + 1)
    Next i
    
    'SUBIR
    
    LOG "SUBIENDO " + CStr(TotBytes) + " BYTES", 0.1
    
    Dim Cnt As Long
    Dim nFileLen As Long 'cantidad de KB que van subidos
    Dim nRet As Long 'valor de retorno para funciones
    Dim nTotFileLen As Long 'totala  subir de un archivo
    Dim sBuffer As String * 1024 'buffer de 1 KB que va subiendo
    Dim SentBytes As Long 'bytes enviados del archivo actual
    Dim sAllBytes As Long 'bytes enviados total
    
    Dim Kam As String 'destino del upload
    Dim Ode As String 'origen del upload
    
    Dim Fs As Long 'Prog.MAX de cada archivo
    Dim StartT As Long 'startTime para ver el RATE de upload
    Dim t As Long 'tiempo actual para caklcular el rate
    Dim p As Long 'es el t/1000
    Dim spRate As Single 'kb/seg
    
    spRate = 0
    sAllBytes = 0
    p = 0
    
    For i = 0 To UBound(ListaFiles)
        
        Fs = FileLen(LocalePath + ListaFiles(i)) 'tamaño del archivo actual
        
        Ode = LocalePath + ListaFiles(i)
        Kam = KLIC + ListaFiles(i)
        
        'escribe el archivo antes de subirle el contenido
        hFile = FtpOpenFile(SERVER, Kam, GENERIC_WRITE, FTP_TRANSFER_TYPE_BINARY, 0)
        If hFile = 0 Then
            LOG "DETENIDO, no se puede crear el archivo en el servidor" + vbCrLf + Ode + vbCrLf + Kam, 0
            UpLoad = 2
            Exit Function
        End If
        SentBytes = 0
        nFileLen = 0
        StartT = GetTickCount 'es como el timer (funcion de VB)
        Open Ode For Binary As #1
            nTotFileLen = LOF(1)
            Do
                Get #1, , sBuffer
                If nFileLen < nTotFileLen - sReadBuffer Then
                    If InternetWriteFile(hFile, sBuffer, sReadBuffer, nRet) = 0 Then
                        LOG "Falla Subiendo el archivo " + Kam + " (byte:" + CStr(nFileLen) + ")", 0
                        UpLoad = 3
                        Exit Do
                    End If
                    SentBytes = SentBytes + sReadBuffer
                    sAllBytes = sAllBytes + sReadBuffer
                    nFileLen = nFileLen + sReadBuffer
                Else
                    If InternetWriteFile(hFile, sBuffer, nTotFileLen - nFileLen, nRet) = 0 Then
                        LOG "Falla subiendo el archivo " + Kam + " (byte:" + CStr(nFileLen) + ")", 0
                        UpLoad = 4
                        Exit Do
                    End If
                    SentBytes = SentBytes + (nTotFileLen - nFileLen)
                    sAllBytes = sAllBytes + (nTotFileLen - nFileLen)
                    nFileLen = nTotFileLen
                End If
                
                'no uso el rate KB/S
                'If SentBytes <> 0 Then
                '    t = GetTickCount - StartT
                '    If t <> 0 Then
                '        spRate = (spRate + ((SentBytes / 1000) / (t / 1000))) / 2
                '        lbSPEED.Caption = Format(spRate, "#.##") & " Kb/S"
                '        lbSPEED.Refresh
                '    End If
                'End If
                LOG "Subiendo " + CStr(sAllBytes) + "/" + CStr(TotBytes), sAllBytes / TotBytes
            Loop Until nFileLen >= nTotFileLen
        Close
        p = t / 1000
        InternetCloseHandle hFile
    Next i
    UpLoad = 0
    LOG "TRANSFERENCIA COMPLETA", 1
    
End Function

Public Function Download(Destino As String, Optional Solo1Archivo As String = "")
    
    'por ahora solo la lista desde lstFiles que tiene el lstBytes
    'destino tiene la carpeta que los contiene con la barra al final
    
    'si Solo1Archivo es <>"" es que no me intyeresa los check y si solo un archivo
    'en particular!
            
    If Right(Destino, 1) <> "\" Then Destino = Destino + "\"
    
    'LEER INFO DE LO QUE SE VA A SUBIR
    Dim TotBytes As Long 'total a subir
    Dim LenFileBytes As Long 'total de cada archivo
    
    TotBytes = 0
    
    Dim i As Integer
    
    If SESSION = 0 Or SERVER = 0 Then
        LOG "No Download, no Server", 0
        Download = 1
        Exit Function
    End If
    
    If lstFILES.SelCount = 0 Then
        Download = 2
        MsgBox "No hay archivos elegdos para descargar!"
        Exit Function
    End If
    
    LOG "Leyendo Info Download", 0.1
    
    For i = 0 To lstFILES.ListCount - 1
        If lstFILES.Selected(i) Then
            If Solo1Archivo <> "" And lstFILES.List(i) <> Solo1Archivo Then GoTo SIG2
            LenFileBytes = CLng(lstBytes.List(i))
            TotBytes = TotBytes + LenFileBytes
            LOG "Leyendo Info download " + CStr(i + 1) + "/" + CStr(lstFILES.ListCount), (i + 1) / lstFILES.ListCount
        End If
SIG2:
    Next i
        
    Dim sBuffer As String 'buffer de 1024 Bytes que va bajando los datos
    Dim FileData As String 'todos los bytes descargados hasta el momento. desde aqui se genera el archivo
    Dim RET As Long 'valor de RETorno para funciones API
    Dim SentBytes As Long 'los bytes de cada archvo que van siendo bajados
    Dim sAllBytes As Long 'todos los bytes bajados de todos los archivos hasta cada momento
    Dim FF As Integer 'para encontrar FreeFile
    Dim Kam As String, Ode As String 'ode es de la WEB y kam el archivo bajado al disco
    Dim StartT As Long 'Star TIME para calcular el RATE (Kb/seg)
    Dim t As Long 'tiempo actual para calcular el RATE (Kb/seg)
    Dim p As Long 'p es segundos (t/1000)
    Dim spRate As Single 'kb/seg
    
    spRate = 0
    sAllBytes = 0
    p = 0
        
    For i = 0 To lstFILES.ListCount - 1
        'solo bajar los elegidos
        If lstFILES.Selected(i) = False Then GoTo SIG
        'ver si entro solo a bajar algun archivo en particular
        If Solo1Archivo <> "" And lstFILES.List(i) <> Solo1Archivo Then GoTo SIG
        
        Ode = KLIC & lstFILES.List(i)
        Kam = Destino & lstFILES.List(i)
        
        hFile = FtpOpenFile(SERVER, Ode, GENERIC_READ, FTP_TRANSFER_TYPE_BINARY, 0)
        If hFile = 0 Then
            LOG "DETENIDO, no se puede abrir el archivio WEB para bajar" + vbCrLf + Ode + vbCrLf + Kam, 0
            Exit Function
        End If
        sBuffer = Space(sReadBuffer)
        FileData = ""
        SentBytes = 0
        StartT = GetTickCount
        Do
            InternetReadFile hFile, sBuffer, sReadBuffer, RET
            If RET <> sReadBuffer Then
                sBuffer = Left$(sBuffer, RET)
            End If
            FileData = FileData + sBuffer
            SentBytes = SentBytes + RET
            sAllBytes = sAllBytes + RET
            
            'rate que no uso kb/s
'            If SentBytes <> 0 Then
'                t = GetTickCount - StartT
'                If t <> 0 Then
'                    spRate = (spRate + ((SentBytes / 1000) / (t / 1000))) / 2
'                    lbSPEED.Caption = Format(spRate, "#.##") & " Kb/S"
'                    lbSPEED.Refresh
'                End If
'            End If
            LOG "Bajando " + lstFILES.List(i), sAllBytes / TotBytes
        Loop Until RET <> sReadBuffer
        FF = FreeFile
        Open Kam For Binary As #FF
            Put #FF, , FileData
        Close #FF
        p = t / 1000
        InternetCloseHandle hFile
SIG:
    Next i
    Download = 0
    LOG "TRANSFERENCIA COMPLETA", 1
    
End Function

Public Function CrearCarpetaWEB(sCarpeta As String)
    Dim sRet As String
    
    On Error GoTo Err2
    
    LOG "Creando carpeta", 0.1
        
    If SESSION = 0 Or SERVER = 0 Then
        CrearCarpetaWEB = 1
        LOG "Crear carpeta Falla por falta de conexion", 0
        Exit Function
    End If
    LOG "Creando carpeta -> Enviando pedido", 0.45
    
    If FtpCreateDirectory(SERVER, KLIC & sCarpeta) = False Then
        CrearCarpetaWEB = 2
        LOG "Crear carpeta Falla", 0
        Exit Function
    End If
    CrearCarpetaWEB = 0
    LOG "Crear carpeta OK", 1
    Exit Function
Err2:
    If Err.Number = 75 Then
        CrearCarpetaWEB = 3
        LOG "Crear carpeta falla, quizas la carpeta ya exista", 0
    End If
End Function

Public Function BorrarCarpetaWEB(sCarpeta As String)
    'borra una carpeta dentro de la carpeta actual (KLIC).
    'No puedo estar adentro y la carpeta debe estar vacia
    
    Dim i As Integer
    If SESSION = 0 Or SERVER = 0 Then
        BorrarCarpetaWEB = 1
        LOG "Borrar carpeta Falla por falta de conexion", 0
        Exit Function
    End If
    
    LOG "Borrando carpeta", 0.4
    
    If FtpRemoveDirectory(SERVER, KLIC + sCarpeta) = False Then
        BorrarCarpetaWEB = 2
        LOG "Borrar carpeta Falla por archivos dentro de la carpeta", 0
        Exit Function
    End If
    BorrarCarpetaWEB = 0
    LOG "Borrar carpeta OK", 1
End Function

Public Function BorrarFileWEB(sFile As String)
    
    LOG "Borrando archivo " + sFile, 0.4
    
    If SESSION = 0 Or SERVER = 0 Then
        BorrarFileWEB = 1
        LOG "Borrar carpeta Falla por falta de conexion", 0
        Exit Function
    End If
    
    If FtpDeleteFile(SERVER, KLIC & sFile) = False Then
        BorrarFileWEB = 2
        LOG "Borrar archivo Falla ", 0
        Exit Function
    End If
    
    BorrarFileWEB = 0
    LOG "Borrar archivo OK", 1
    
End Function

Public Function RenameFileWEB(oldName As String, NewName As String)
    
    LOG "Renombrando archivo " + oldName, 0.4
    
    On Error GoTo Err
    
    If SESSION = 0 Or SERVER = 0 Then
        RenameFileWEB = 1
        LOG "Rename Falla por falta de conexion", 0
        Exit Function
    End If
    
    oldName = KLIC & oldName
    NewName = KLIC & NewName
    
    If FtpRenameFile(SERVER, oldName, NewName) = False Then
        RenameFileWEB = 2
        LOG "Renombrando archivo FALLA" + oldName, 0
        Exit Function
    End If
    
    RenameFileWEB = 0
    LOG "Renombrando archivo OK", 1
    Exit Function
    
Err:
    If Err.Number = 58 Then
        RenameFileWEB = 3
        LOG "Renombrando archivo falla, puede haber otro archivo con el mismo nombre", 0
    End If

End Function

Public Sub DisConnect()
    LOG "Conexion: Cerrando Server...", 0.2
    InternetCloseHandle SERVER
    
    LOG "Conexion: Cerrando Sesion...", 5
    InternetCloseHandle SESSION
        
    SERVER = 0: SESSION = 0
    LOG "DESCONECTADO", 1
    
    lstFOLDERS.Clear
    lstBytes.Clear
    lstFILES.Clear
End Sub

Private Sub Label3_DblClick()
    If SERVER = 0 Then
        Connect "ftp.psiap.com", "zlg", "zlg90", 21
        List
    Else
        DisConnect
    End If
End Sub

Private Sub lstFILES_Click()
    lstBytes.Visible = False
    lstFILES.Visible = False
    lstBytes.ListIndex = lstFILES.ListIndex
    lstBytes.Visible = True
    lstFILES.Visible = True
End Sub

Private Sub lstFILES_Scroll()
    lstBytes.Visible = False
    lstFILES.Visible = False
    lstBytes.TopIndex = lstFILES.Top
    lstBytes.Visible = True
    lstFILES.Visible = True
End Sub

Private Sub lstFOLDERS_DblClick()
    lstFOLDERS.Enabled = False
    If lstFOLDERS = "." Or lstFOLDERS = ".." Then
        'buscar la barra anterior para volvber una carpeta atras
        Dim Lastbarra As Long
        Lastbarra = InStrRev(KLIC, "\")
        'si la ultima es el primer caracrter es que estaba en el primer nivel de capetas
        If Lastbarra = 1 Then
            KLIC = "\"
        Else
            Lastbarra = InStrRev(KLIC, "\", Lastbarra - 1)
            KLIC = Mid(KLIC, 1, Lastbarra)
        End If
    Else
        KLIC = KLIC + lstFOLDERS + "\"
    End If
    UbicarseEnFolder KLIC
    List
    lstFOLDERS.Enabled = True
End Sub

Private Sub UserControl_Initialize()
    LOG "Iniciado", 0
End Sub

Public Function UbicarseEnFolder(sFolder As String) As Long
    
    LOG "Cambiando a carpeta " + sFolder, 0.2
    
    If SESSION = 0 Or SERVER = 0 Then
        UbicarseEnFolder = 1
        LOG "ChgDIR Falla por falta de conexion", 0
        Exit Function
    End If
    
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder + "\"
    
    KLIC = sFolder
    ADR = KLIC + "*.*"
    
    ' DA FALSO PERO FUNCIONA LA PUTA MADRE
'    If FtpSetCurrentDirectory(SESSION, ADR) = False Then
'        LOG "No se pudo ubicar en la carpeta!", 0
'        UbicarseEnFolder = 2
'        Exit Function
'    End If
    FtpSetCurrentDirectory SESSION, ADR
    
    LOG "Ubicarse OK", 1
    UbicarseEnFolder = 0
    lblFULLPATH = KLIC
End Function

Public Function GetFolderWEBName() As String
    LOG "Verfificando FOLDER...", 0.3
    'ver en que carpeta estoy
    ADR = Space(260)
    FtpGetCurrentDirectory SERVER, ADR, Len(ADR)

    ADR = Left(ADR, InStr(1, ADR, Chr(0)) - 1)
    ADR = ADR & IIf((Right(ADR, 1) = "\"), "*.*", "\*.*")

    'ver si esta ok
    GetFolderWEBName = ADR
    LOG "Verfificando OK= " + ADR, 0.3
End Function

Private Sub UserControl_Resize()
    lblINFO.Width = UserControl.Width - lblINFO.Left - 50
    txtLOG.Width = lblINFO.Width
End Sub
