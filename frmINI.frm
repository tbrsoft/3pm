VERSION 5.00
Begin VB.Form frmINI 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmINI.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9225
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Label lblTipoLIC 
      Alignment       =   2  'Center
      BackColor       =   &H00000040&
      Caption         =   "Iniciando 3PM. Licencia Full"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7500
      TabIndex        =   4
      Top             =   60
      Width           =   4440
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1095
      Left            =   3900
      Picture         =   "frmINI.frx":0442
      Stretch         =   -1  'True
      Top             =   7860
      Width           =   3570
   End
   Begin VB.Label pBar 
      BackColor       =   &H0000FFFF&
      Height          =   240
      Left            =   30
      TabIndex        =   3
      Top             =   6870
      Width           =   11895
   End
   Begin VB.Label VVV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "v 8.8.88"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   1740
      TabIndex        =   0
      Top             =   2070
      Width           =   2460
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Contando Discos: 00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   30
      TabIndex        =   2
      Top             =   7170
      Width           =   11910
   End
   Begin VB.Label lblProces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscando discos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Top             =   6540
      UseMnemonic     =   0   'False
      Width           =   11850
   End
   Begin VB.Image TapaCD 
      Height          =   4215
      Left            =   7560
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   4305
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4020
      Left            =   1710
      Picture         =   "frmINI.frx":1EB0
      Stretch         =   -1  'True
      Top             =   1410
      Width           =   4710
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    MostrarCursor False
    ClaveAdmin = "ERO77701192FF"
    VVV = "v " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    '--------
    If TypeVersion = "SL" Then
        If FSO.FileExists(WINfolder + "\SL\imgbig.tbr") Then Image1.Picture = LoadPicture(WINfolder + "\SL\imgbig.tbr")
        If FSO.FileExists(WINfolder + "\SL\imgtbr.tbr") Then Image2.Picture = LoadPicture(WINfolder + "\SL\imgtbr.tbr")
                
    End If
    '--------
    Select Case TypeVersion
        Case "SL"
            lblTipoLIC = "Iniciando SUPERLICENCIA"
        Case "FULL"
            lblTipoLIC = "Iniciando 3PM. Licencia Full"
        Case "DEMO2"
            lblTipoLIC = "Iniciando Demo gratuito"
        Case "DEMO"
            lblTipoLIC = "Iniciando demo 3PM"
    End Select
    lblTipoLIC.Refresh
            
    AjustarFRM Me, 12000
    'leer el archivo de configuracion SYSfolder + "3pmcfg.tbr"
    CargarIMGinicio = LeerConfig("CargarImagenInicio", "1")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI", "LISTA")
    VolumenIni = CLng(LeerConfig("Volumen", "50"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    FASTini = LeerConfig("FastIni", "1")
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro", "1")
    TapasMostradasH = Val(LeerConfig("DiscosH", "3"))
    TapasMostradasV = Val(LeerConfig("DiscosV", "2"))
    PasarHoja = LeerConfig("Pasarhoja", "1")
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    RankToPeople = LeerConfig("RankToPeople", "1")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosCuestaTema = LeerConfig("CreditosCuestaTema", "1")
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    
    'cargar variables de claves
    'archivo de claves
    If FSO.FileExists(WINfolder + "\sevalc.dll") = False Then
        Set TE = FSO.CreateTextFile(WINfolder + "\sevalc.dll", True)
        TE.WriteLine "Config:12345612345612345612"
        TE.WriteLine "Close:45612345612345612345"
        TE.WriteLine "Credit:1234441234441234561"
        TE.Close
    End If
    Set TE = FSO.OpenTextFile(WINfolder + "\sevalc.dll", ForReading, False)
    'config/close/credit es el orden del archivo
    ClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    ClaveClose = txtInLista(TE.ReadLine, 1, ":")
    ClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    Me.Show
    Me.Refresh
    
    
    'ver si ya estaba cargado
    If App.PrevInstance Then MsgBox "No se pueden abrir dos instancias de 3pm": End
        
    'ASEGURARSE QUE EXISTA la carpeta del ranking y la imagen que le corresponde
    If FSO.FolderExists(AP + "discos") = False Then
        FSO.CreateFolder AP + "discos"
    End If
    If FSO.FolderExists(AP + "discos\01- Los mas escuchados") = False Then
        FSO.CreateFolder AP + "discos\01- Los mas escuchados"
     End If
    'siempre copiarlo
    'If FSO.FileExists(AP + "discos\01- Los mas escuchados\tapa.jpg") = False Then
        If FSO.FileExists(AP + "top10.jpg") Then
            FSO.CopyFile AP + "top10.jpg", AP + "discos\01- Los mas escuchados\tapa.jpg", True
        Else
            MsgBox "No se encuentra el archivo de imagen del Ranking!. La instalacion de 3PM no es corecta!"
            End
        End If
    'End If
    If FSO.FileExists(AP + "tapa.jpg") = False Then
        MsgBox "No se encuentra el archivo de imagen de las portadas predeterminadas!. La instalacion de 3PM no es corecta!"
        End
    End If
    'carpeta del protector
    If FSO.FolderExists(AP + "fotos") = False Then
        FSO.CreateFolder AP + "fotos"
    End If
    
    TECLAS_PRES = "11111222223333344444" 'arranca con 20 pulsaciones
    
    
    '===================ORDENAR EL RANKING================================
    On Error GoTo notop
    'ver si existe ranking.tbr
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
        'si me quedo da error
        GoTo FinOrden
    End If
    lblINI = "Inicializando 3PM..."
    lblINI.Refresh
    pBar.Width = 0
    pBar.Refresh
    Dim TT As String
    Dim mtxTOP10() As String, z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    
    'sin esto los archivos vacios se clavan
    ReDim Preserve mtxTOP10(0)
    
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        TT = TE.ReadLine
        If TT <> "" Then
            z = z + 1
            pBar.Width = z * 10
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            ReDim Preserve mtxTOP10(z)
            mtxTOP10(z) = TT
        End If
    Loop
    TE.Close
    'ordenar la matriz
    'tomar la matriz (con valores separador) y ordenala en base a la
    'columna indicada. en este caso el separador es "," y la columna es 0.
    'seria los mismo que tomara 1 ya que todos tienen el mismo path
    Dim MaxPT As Long 'comparacoin de cadenas. Empiezo con el máximo
    Dim ubicMAX As Long 'indice en la matriz del menor encontrado cada vuelta
    MaxPT = 0
    Dim c As Long, mtx As Long, ValComp As Long
    c = 0 'cantidad de minimos encontrados
    Dim Ordenados() As Long 'matriz con los indices ordenados
    pBar.Width = 0
    pBar.Refresh
    Do
        pBar.Width = c * 10
        For mtx = 1 To UBound(mtxTOP10)
            'se compara por los puntos
            ValComp = Val(txtInLista(mtxTOP10(mtx), 0, ","))
            If ValComp > MaxPT Then
                'nunca uno sumara mas de dos puntos (legalmente)
                MaxPT = ValComp
                ubicMAX = mtx
            End If
        Next
        'al mayor lo quito para que no salga de nuevo
        mtxTOP10(ubicMAX) = "0," + mtxTOP10(ubicMAX)
        c = c + 1
        ReDim Preserve Ordenados(c)
        Ordenados(c) = ubicMAX
        If c >= UBound(mtxTOP10) Then Exit Do
        MaxPT = 0
    Loop
    'cargar todos y sacar la primera columna de las zetas
    pBar.Width = 0
    pBar.Refresh
            
    Dim MTXsort() As String
    Set TE = FSO.CreateTextFile(AP + "ranking.tbr", True)
    For mtx = 1 To UBound(mtxTOP10)
        ReDim Preserve MTXsort(mtx)
        'como se agrego un indice mas en archivo esta en el indice2
        'ver si existe si si no no cargarlo
        If Dir(txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",")) <> "" Then
            MTXsort(mtx) = txtInLista(mtxTOP10(Ordenados(mtx)), 1, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 2, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 3, ",") + "," + _
                txtInLista(mtxTOP10(Ordenados(mtx)), 4, ",")
            TE.WriteLine MTXsort(mtx)
            pBar.Width = mtx * 10
        Else
            WriteTBRLog "limpiado del Rank: " + txtInLista(mtxTOP10(Ordenados(mtx)), 2, ","), False
            Limpiaron = Limpiaron + 1
        End If
    Next
    TE.Close
    
    If Limpiaron > 0 Then WriteTBRLog "Se limpiaron " + CStr(Limpiaron) + " temas", True
    '==================================================================
FinOrden:
    'se inicializa el contador para que la variable CONTADOR tenga el
    'valor de todas las fichas cargadas
    'si este es cero esta en los primeros usos entonces mostrar el CLUF
    SumarContadorCreditos 0
    frmIndex.Show 1
    Exit Sub
notop:
    WriteTBRLog "frmINI - Ranking ordenar. " + vbCrLf + Err.Description, True
    Resume Next
End Sub

