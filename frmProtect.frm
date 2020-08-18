VERSION 5.00
Begin VB.Form frmProtect 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Left            =   9270
      Top             =   2580
   End
   Begin VB.Frame FR 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0FFFF&
      Height          =   8940
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      Begin VB.Label lblDISCO 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Protecci�n de pantalla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF80&
         Height          =   555
         Left            =   60
         TabIndex        =   2
         Top             =   8430
         UseMnemonic     =   0   'False
         Width           =   11535
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   2
         Left            =   3375
         Stretch         =   -1  'True
         Top             =   405
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   0
         Left            =   180
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   1
         Left            =   360
         Stretch         =   -1  'True
         Top             =   3330
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   3
         Left            =   3330
         Stretch         =   -1  'True
         Top             =   3420
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   4
         Left            =   6390
         Stretch         =   -1  'True
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Image PicProtec 
         BorderStyle     =   1  'Fixed Single
         Height          =   2640
         Index           =   5
         Left            =   6435
         Stretch         =   -1  'True
         Top             =   3420
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label lblTIT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Se esta protegiendo la pantalla. Presione cualquier tecla para continuar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   2265
         Left            =   10290
         TabIndex        =   1
         Top             =   210
         Width           =   1545
      End
   End
End
Attribute VB_Name = "frmProtect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TiempoEnProtect As Long 'en segundos
Dim Intervalo As Long 'en segundos
Dim NumFotoIni  As Long
Dim PropAchicar As Double
Dim PROP As Double, PROP2 As Double
Dim TopTit As Long, movTit As Integer
Dim IndPicVisible As Integer
Dim MTXtapas() As String
Dim IndMtxTapaVisible As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            End
    End Select
    SecSinTecla = 0
    frmIndex.lblNoTecla = 0
    Unload Me
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    On Error GoTo MiErr
    tERR.Anotar "acom", KeyCode, Shift
    
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            ShowCredits
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "radilav.cfg", CStr(CreditosValidar)
            
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
    tERR.Anotar "acon"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acok"
    Resume Next
End Sub

Private Sub Form_Load()
    On Error GoTo MiErr
    tERR.Anotar "acoo"
    Intervalo = 2
    AjustarFRM Me, 12000
    FR.Left = Screen.Width / 2 - FR.Width / 2
    FR.Top = Screen.Height / 2 - FR.Height / 2
    TopTit = 150: movTit = 40
    IndPicVisible = 0
    IndMtxTapaVisible = 0
    Dim NombreDir As String, ContadorArch As Long
    PicProtec(0).Stretch = (Protector = 1)
    PicProtec(1).Stretch = (Protector = 1)
    PicProtec(2).Stretch = (Protector = 1)
    PicProtec(3).Stretch = (Protector = 1)
    PicProtec(4).Stretch = (Protector = 1)
    PicProtec(5).Stretch = (Protector = 1)
    lblDISCO.Visible = (Protector = 1)
    'VER POR QUE NUMERO DE FOTO IVA
    NumFotoIni = Val(ReadSimpleFile)
    If (Protector = 1) Then
        tERR.Anotar "acop"
        ContadorArch = 0
        'hacer una lista de las tapas disponibles
        ruta = AP + "discos\"
        NombreDir = Dir$(ruta & "*.*", vbDirectory)
        Do While Len(NombreDir)
            If NombreDir = "." Or NombreDir = ".." Then
                ' excluir las entradas "." y ".."
            ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                ' este es un archivo normal
            Else
                tERR.Anotar "acoq", ruta + NombreDir
                ' es un directorio
                If FSO.FileExists(ruta & NombreDir + "\tapa.jpg") Then
                    ContadorArch = ContadorArch + 1
                    ReDim Preserve MTXtapas(ContadorArch) As String
                    MTXtapas(ContadorArch) = ruta & NombreDir & "\tapa.jpg"
                End If
            End If
            NombreDir = Dir$
            tERR.Anotar "acor", NombreDir
        Loop
    End If
    If (Protector = 2) Then
        tERR.Anotar "acos"
        ContadorArch = 0
        'hacer una lista de las fotos disponibles
        ruta = AP + "fotos\"
        NombreDir = Dir$(ruta & "*.jpg")
        Do While Len(NombreDir)
            If NombreDir = "." Or NombreDir = ".." Then
                ' excluir las entradas "." y ".."
            ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                
                'este es un archivo normal
                ContadorArch = ContadorArch + 1
                ReDim Preserve MTXtapas(ContadorArch) As String
                MTXtapas(ContadorArch) = ruta & NombreDir
                tERR.Anotar "acot", ContadorArch, MTXtapas(ContadorArch)
            Else
                ' es un directorio
            End If
            NombreDir = Dir$
            tERR.Anotar "acou", NombreDir
        Loop
    End If
    'si no hay archivos en fotos da error!!!!
    tERR.Anotar "acov", ContadorArch
    If ContadorArch = 0 Then
        lblDISCO = "!!!!!!No hay fotos para mostrar!!!!"
        lblDISCO.Visible = True
    Else
        TiempoEnProtect = 0
        Timer1.Interval = Intervalo * 1000
        IndMtxTapaVisible = NumFotoIni
    End If
    tERR.Anotar "acow"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acol"
    Resume Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'dejar grabada la ultima foto que se vio
    WriteSimpleFile CStr(IndMtxTapaVisible)
End Sub

Private Sub Timer1_Timer()
    On Error GoTo MiErr
    TiempoEnProtect = TiempoEnProtect + (Intervalo * 1000)
    tERR.Anotar "acoy", TiempoEnProtect
    If DuracionProtect > 0 Then 'si duuracion protect=0 no sale hasta que toquen tecla
        If TiempoEnProtect > DuracionProtect * 1000 Then
            tERR.Anotar "acoz", TiempoEnProtect, DuracionProtect
            Timer1.Interval = 0
            Unload Me
        End If
    End If
        
    TopTit = TopTit + movTit
    lblTIT.Top = TopTit
    If TopTit > 5130 Then movTit = movTit * (-1)
    If TopTit < 100 Then movTit = movTit * (-1)
    tERR.Anotar "acpa"
    PicProtec(IndPicVisible).Visible = False
    IndPicVisible = IndPicVisible + 1
    IndMtxTapaVisible = IndMtxTapaVisible + 1
    If IndPicVisible = 6 Then IndPicVisible = 0
    'ver que no se pase del total de fotos
    If IndMtxTapaVisible > UBound(MTXtapas) Then IndMtxTapaVisible = 1
    
    PicProtec(IndPicVisible).Stretch = (Protector = 1)
    PicProtec(IndPicVisible).Picture = LoadPicture(MTXtapas(IndMtxTapaVisible))
    PROP = PicProtec(IndPicVisible).Height / PicProtec(IndPicVisible).Width
    tERR.Anotar "acpb", Protector
    If (Protector = 1) Then
        Dim DISCO As String
        DISCO = Left(MTXtapas(IndMtxTapaVisible), Len(MTXtapas(IndMtxTapaVisible)) - 9)
        DISCO = FSO.GetBaseName(DISCO)
        lblDISCO = DISCO
        PicProtec(IndPicVisible).Stretch = True
    End If
    If (Protector = 2) Then
        'si es muy grande
        If PicProtec(IndPicVisible).Height > FR.Height * 0.8 Or PicProtec(IndPicVisible).Width > FR.Width * 0.8 Then
            'llevar a un tama�o decente
            PicProtec(IndPicVisible).Stretch = True
            tERR.Anotar "acpc", PROP
            If PROP > 1 Then
                'fr es menor al alto
                PropAchicar = (FR.Height * 0.8) / PicProtec(IndPicVisible).Height
            Else
                'fr es menor al ancho
                PropAchicar = (FR.Width * 0.8) / PicProtec(IndPicVisible).Width
            End If
            PicProtec(IndPicVisible).Height = PicProtec(IndPicVisible).Height * PropAchicar
            PicProtec(IndPicVisible).Width = PicProtec(IndPicVisible).Width * PropAchicar
            PROP2 = PicProtec(IndPicVisible).Height / PicProtec(IndPicVisible).Width
            If PROP - PROP2 > 0.1 Then
                tERR.Anotar "acpd.NOSEVE=", MTXtapas(IndMtxTapaVisible)
                'WriteTBRLog "La imagen del protect no se mostro correctamente. Foto:" + _
                    CStr(MTXtapas(IndMtxTapaVisible)), True
            End If
            
        Else
            PicProtec(IndPicVisible).Stretch = False
        End If
    End If
    
    Randomize Timer
    B = lblDISCO.Top - PicProtec(IndPicVisible).Height
    If B < 150 Then B = 150 '150 es el tope del frmae
    tERR.Anotar "acpe"
    A = Int(Rnd * B)
    PicProtec(IndPicVisible).Top = A
    
    Randomize Timer
    B = lblTIT.Left - PicProtec(IndPicVisible).Width
    A = Int(Rnd * B)
    PicProtec(IndPicVisible).Left = A
    
    PicProtec(IndPicVisible).Visible = True
    tERR.Anotar "acpf"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acom"
    Resume Next
End Sub

Private Sub WriteSimpleFile(TXT As String)
    
    If FSO.FileExists(AP + "protect.tbr") = False Then
        Set TE = FSO.CreateTextFile(AP + "protect.tbr", False)
        TE.Close
    End If
    Set TE = FSO.OpenTextFile(AP + "protect.tbr", ForWriting, False)
    
    TE.WriteLine TXT
    
    TE.Close
End Sub

Private Function ReadSimpleFile() As String
    
    If FSO.FileExists(AP + "protect.tbr") = False Then
        Set TE = FSO.CreateTextFile(AP + "protect.tbr", False)
        TE.Close
        ReadSimpleFile = "0"
        Exit Function
    End If
    
    Set TE = FSO.OpenTextFile(AP + "protect.tbr", ForReading, False)
    If TE.AtEndOfStream Then
        ReadSimpleFile = ""
    Else
        ReadSimpleFile = TE.ReadLine
    End If
    
    TE.Close
End Function

