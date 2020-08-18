VERSION 5.00
Begin VB.Form frmTemasDeDisco 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   8805
      Left            =   90
      TabIndex        =   0
      Top             =   30
      Width           =   11805
      Begin VB.ListBox lstTemas 
         BackColor       =   &H00404080&
         BeginProperty Font 
            Name            =   "HandelGotDLig"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0E0FF&
         Height          =   7710
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   150
         Width           =   7035
      End
      Begin VB.Label lblDataDisco 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "No hay datos adicionales del disco"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2895
         Left            =   7200
         TabIndex        =   5
         Top             =   4950
         UseMnemonic     =   0   'False
         Width           =   4545
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   9720
         X2              =   9720
         Y1              =   910
         Y2              =   210
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   7770
         X2              =   9720
         Y1              =   900
         Y2              =   180
      End
      Begin VB.Label lblDisco 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   7200
         TabIndex        =   4
         Top             =   4290
         UseMnemonic     =   0   'False
         Width           =   4545
      End
      Begin VB.Label lblIndicaciones 
         Alignment       =   2  'Center
         BackColor       =   &H00000080&
         Caption         =   "Utilize las flechas para desplazarse sobre los distintos temas, OK = escuchar. Escape = Salir"
         BeginProperty Font 
            Name            =   "HandelGotDLig"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   7890
         Width           =   11565
      End
      Begin VB.Image TapaCD 
         BorderStyle     =   1  'Fixed Single
         Height          =   3300
         Left            =   7740
         Stretch         =   -1  'True
         Top             =   930
         Width           =   3465
      End
      Begin VB.Label lblTemaSonando 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmTemasDeDisco.frx":0000
         BeginProperty Font 
            Name            =   "Humanst531 BT"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   2
         Top             =   8520
         Width           =   11775
      End
   End
End
Attribute VB_Name = "frmTemasDeDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyW
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS < 9 Then
                CREDITOS = CREDITOS + 1
                frmINDEX.lblCreditos = "Creditos: 0" + Str(CREDITOS)
            End If
        Case vbKeyE
            ESTOY = 0
            Unload Me
        Case vbKeyP
            'ver si esta habilitado
            If CREDITOS > 0 Then
                CREDITOS = CREDITOS - 1
                frmINDEX.lblCreditos = "Creditos: 0" + Str(CREDITOS)
                Dim temaElegido As String
                temaElegido = UbicDiscoActual + lstTemas + ".mp3"
                'temaElegido = txtInLista(MATRIZ_TEMAS(lstTemas.ListIndex + 1), 0, ",")
                
                'si esta ejecutando pasa a la lista de reproducción
                If ESTOY_REPRODUCIENDO Then
                    'pasar a la lista de reproducción
                    Dim NewIndLista As Long
                    NewIndLista = UBound(MATRIZ_LISTA)
                    ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                    'se graba en Matriz_Listas como patah, nombre(sin ".mp3")
                    MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTemas + " / " + FSO.GetBaseName(UbicDiscoActual)
                    CargarProximosTemas
                Else
                    'TEMA_REPRODUCIENDO y ESTOY_REPRODUCIENDO se cargan en ejecutartema
                    EjecutarTema temaElegido
                End If
                'pase lo que pase me vuelvo a los discos y cierro ventana actual
                ESTOY = 0
                Unload Me
            End If
        
        Case vbKeyI
            If lstTemas.ListIndex < lstTemas.ListCount - 1 Then lstTemas.ListIndex = lstTemas.ListIndex + 1
        
        Case vbKeyU
            If lstTemas.ListIndex > 0 Then lstTemas.ListIndex = lstTemas.ListIndex - 1
    End Select
End Sub

Private Sub Form_Load()
    Frame1.Left = Screen.Width / 2 - Frame1.Width / 2
    Frame1.Top = Screen.Height / 2 - Frame1.Height / 2
    Dim ArchTapa As String
    ArchTapa = UbicDiscoActual + "tapa.jpg"
    If FSO.FileExists(ArchTapa) Then
        TapaCD.Picture = LoadPicture(ArchTapa)
    Else
        TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
    End If
    lblDisco = FSO.GetBaseName(UbicDiscoActual)
    Dim ArchData As String
    ArchData = UbicDiscoActual + "data.txt"
    If FSO.FileExists(ArchData) Then
        
        Dim a As TextStream
        Set a = FSO.OpenTextFile(ArchData, ForReading, False)
        lblDataDisco = a.ReadAll
    Else
        lblDataDisco = "No hay datos adicionales de este disco"
    End If
    
End Sub

