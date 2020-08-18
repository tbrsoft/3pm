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
   Begin VB.Label pBar 
      BackColor       =   &H0000FFFF&
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   7155
      Width           =   11985
   End
   Begin VB.Label VVV 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "v 8.8.88"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   420
      Left            =   2835
      TabIndex        =   0
      Top             =   3420
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000080&
      Height          =   1005
      Left            =   0
      TabIndex        =   3
      Top             =   8415
      Width           =   12120
   End
   Begin VB.Label lblINI 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Contando Discos: 00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   825
      Left            =   0
      TabIndex        =   2
      Top             =   7470
      Width           =   11910
   End
   Begin VB.Image Image3 
      Height          =   1800
      Left            =   8955
      Picture         =   "frmINI.frx":0442
      Stretch         =   -1  'True
      Top             =   180
      Width           =   1650
   End
   Begin VB.Label lblProces 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buscando discos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   0
      TabIndex        =   1
      Top             =   6480
      UseMnemonic     =   0   'False
      Width           =   12000
   End
   Begin VB.Image TapaCD 
      Height          =   4215
      Left            =   7650
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   4305
   End
   Begin VB.Image Image1 
      Height          =   6300
      Left            =   135
      Picture         =   "frmINI.frx":1346
      Stretch         =   -1  'True
      Top             =   90
      Width           =   7500
   End
End
Attribute VB_Name = "frmINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    VVV = "v " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    AjustarFRM Me, 12000
    
    'leer el archivo de configuracion ap+"config.tbr"
    CargarIMGinicio = LeerConfig("CargarImagenInicio")
    AutoReDibuj = LeerConfig("AutoReDraw")
    TeclaDER = Val(LeerConfig("TeclaDerecha"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt"))
    TeclaOK = Val(LeerConfig("TeclaOK"))
    TeclaESC = Val(LeerConfig("TeclaESC"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha"))
    TeclaConfig = Val(LeerConfig("TeclaConfig"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre")
    MaximoFichas = Val(LeerConfig("MaximoFichas"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos"))
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI")
    VolumenIni = CLng(LeerConfig("Volumen"))
    EsperaTecla = Val(LeerConfig("EsperaTecla"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema"))
    FASTini = LeerConfig("FastIni")
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro")
    TapasMostradasH = Val(LeerConfig("DiscosH"))
    TapasMostradasV = Val(LeerConfig("DiscosV"))
    verTiempoRestante = LeerConfig("VerTiempoRestante")
    verTemasEnLista = LeerConfig("verTemasEnLista")
    verCreditos = LeerConfig("verCreditos")
    verTOTdiscos = LeerConfig("verTotDiscos")
    verPuesto = LeerConfig("verPuesto")
    verLista = LeerConfig("verLista")
    PasarHoja = LeerConfig("Pasarhoja")
    DistorcionarTapas = LeerConfig("DistorcionarTapas")
    ProtectOriginal = LeerConfig("ProtectOriginal")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas")
    MostrarRotulos = LeerConfig("MostrarRotulos")
    RotulosArriba = LeerConfig("RotulosArriba")
    DuracionProtect = LeerConfig("DuracionProtect")
    RankToPeople = LeerConfig("RankToPeople")
    
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
    If FSO.FileExists(AP + "discos\01- Los mas escuchados\tapa.jpg") = False Then
        FSO.CopyFile AP + "top10.jpg", AP + "discos\01- Los mas escuchados\tapa.jpg", True
    End If
    'carpeta del protector
    If FSO.FolderExists(AP + "fotos") = False Then
        FSO.CreateFolder AP + "fotos"
    End If
    
    TECLAS_PRES = "11111222223333344444" 'arranca con 20 pulsaciones
    
    
    '===================ORDENRA EL RANKING================================
    On Error GoTo notop
    'ver si existe ranking.tbr
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        FSO.CreateTextFile AP + "ranking.tbr", True
        'si me quedo da error
        GoTo FinOrden
    End If
    lblINI = "Inicializando 3PM..."
    lblINI.Refresh
    PBar.Width = 0
    PBar.Refresh
    Dim TT As String
    Dim mtxTOP10() As String, Z As Integer
    Dim ThisArch As String
    Dim ThisTEMA As String
    Dim ThisDISCO As String
    Dim ThisPTS As Long
    Dim Encontrado As Boolean
    Encontrado = False
    'abrir el archivo y CARGARLO A UNA MATRIZ
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    'leerlo cargarlo en matriz y ordenar por mas escuchado
    
    'sin esto los archivos vacios se clavan
    ReDim Preserve mtxTOP10(0)
    
    Do While Not TE.AtEndOfStream
        'cada linea es "puntos,arch,nombretema,nombredisco"
        TT = TE.ReadLine
        If TT <> "" Then
            Z = Z + 1
            PBar.Width = Z * 10
            ThisPTS = Val(txtInLista(TT, 0, ","))
            ThisArch = txtInLista(TT, 1, ",")
            ThisTEMA = txtInLista(TT, 2, ",")
            ThisDISCO = txtInLista(TT, 3, ",")
            ReDim Preserve mtxTOP10(Z)
            mtxTOP10(Z) = TT
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
    PBar.Width = 0
    PBar.Refresh
    Do
        PBar.Width = c * 10
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
    PBar.Width = 0
    PBar.Refresh
            
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
            PBar.Width = mtx * 10
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
    frmINDEX.Show 1
    Exit Sub
notop:
    WriteTBRLog Err.Description, True
    Resume Next
End Sub

