VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00404080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de 3PM"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recuperar Licencia, ya estaba cargada (tecla Izquierda 6 veces)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7080
      Width           =   1905
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Cancel          =   -1  'True
      Caption         =   "COMPRAR AHORA!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7590
      Width           =   5580
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir Manual de uso. RECOMENDADO si es su primer uso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2580
      TabIndex        =   19
      Top             =   6570
      Width           =   6915
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404080&
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia de Usuario Final"
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
      Height          =   375
      Left            =   1920
      TabIndex        =   18
      Top             =   6150
      Value           =   1  'Checked
      Width           =   8655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ver CLUF"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7020
      TabIndex        =   17
      Top             =   7050
      Width           =   1785
   End
   Begin VB.TextBox LBL 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   90
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "frmREG.frx":0000
      Top             =   90
      Width           =   10245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5160
      TabIndex        =   10
      Top             =   7050
      Width           =   1785
   End
   Begin VB.ComboBox cmbCountry 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4410
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5790
      Width           =   2805
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   7
      Left            =   8970
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   6
      Left            =   7890
      MaxLength       =   5
      TabIndex        =   6
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   5
      Left            =   6810
      MaxLength       =   5
      TabIndex        =   5
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   4
      Left            =   5730
      MaxLength       =   5
      TabIndex        =   4
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   3
      Left            =   4650
      MaxLength       =   5
      TabIndex        =   3
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   2
      Left            =   3570
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   1
      Left            =   2490
      MaxLength       =   5
      TabIndex        =   1
      Top             =   3960
      Width           =   1050
   End
   Begin VB.TextBox lblGUID 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Aqui va el codigo"
      Top             =   3030
      Width           =   11805
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   0
      Left            =   1410
      MaxLength       =   5
      TabIndex        =   0
      Top             =   3960
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3270
      TabIndex        =   9
      Top             =   7050
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "07.50"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   405
      Index           =   4
      Left            =   60
      TabIndex        =   21
      Top             =   7740
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":0006
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Index           =   3
      Left            =   60
      TabIndex        =   15
      Top             =   4560
      Width           =   11805
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   10380
      Stretch         =   -1  'True
      Top             =   30
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique su pais de residencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   14
      Top             =   5580
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Si ya recibio el código de tbrSoft cárguelo aqui. Respete mayúsculas y minúsculas!!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   13
      Top             =   3750
      Width           =   11505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":00B9
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   1710
      Width           =   11805
   End
End
Attribute VB_Name = "frmREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GGG As String
Dim LastTeclas As String
Dim CarpetaSys As String

Private Sub Check1_Click()
    Command1.Enabled = Check1
End Sub

Private Sub Command1_Click()
    
    Dim CodigoUnido As String
    CodigoUnido = txtCOD(0) + "-" + txtCOD(1) + "-" + txtCOD(2) + "-" + _
        txtCOD(3) + "-" + txtCOD(4) + "-" + txtCOD(5) + "-" + _
        txtCOD(6) + "-" + txtCOD(7)

    If cmbCountry = "(SELECCIONE PAIS)" Then
        MsgBox "Debe cargar el pais de residencia"
        Exit Sub
    End If
    
   'dar ingreso a la clave y la grabo solo aqui
    Dim zz
    zz = K.IngresaClave(CodigoUnido, True)
        
    If K.LICENCIA = aSinCargar Then
        Dim TXTmsg As String
        TXTmsg = "3PM en version Demo tiene limite de discos, " + _
            "y trunca los temas a los 2 minutos." + vbCrLf + _
            "¿Desea ejecutar 3PM en version demo?"
        If MsgBox(TXTmsg, vbQuestion + vbYesNo, "3PM demo") = vbNo Then End
    End If
        
    If K.LICENCIA = BErronea Then
        MsgBox "Existen datos erroneos de la licencia. Si ingresa claves equivocadas o ha" + _
            " reemplazado componentes de su PC debe comunicarse con tbrSoft o su proveedor" + _
            " de 3PM para solucionar este inconveniente"
            Exit Sub
    End If
    
    If K.LICENCIA = CGratuita Then MsgBox "Clave gratuita de 3PM. " + vbCrLf + "VARIACION: " + CStr(K.VariacionClave)
    
    If K.LICENCIA = GFull Then MsgBox "El codigo se ha cargado correctamente. Bienvenido a 3PM " + vbCrLf + "VARIACION: " + CStr(K.VariacionClave)
    
    If K.LICENCIA = HSuperLicencia Then MsgBox "SUPERLICENCIA de 3PM. El codigo de SuperLicencia se ha cargado correctamente. Bienvenido a Super3PM" + vbCrLf + "VARIACION: " + CStr(K.VariacionClave)
    
    Unload Me
    frmINI.Show 1
                    
    tERR.Anotar "acpg"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acph"
    Resume Next
                    
End Sub

Private Sub Command2_Click()
    Unload Me
    End
End Sub

Private Sub Command3_Click()
    AbrirArchivo AP + "manual.doc", Me
End Sub

Private Sub Command4_Click()
    frmCLUF.Show 1
End Sub

Private Sub Command5_Click()
    frmCompraYA.Show 1
End Sub

Private Sub Command6_Click()
    A = Shell(AP + "repair.exe", vbNormalFocus)
    'MsgBox A
    Unload Me
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    LastTeclas = LastTeclas + Chr(KeyCode)
    LastTeclas = Right(LastTeclas, 6)
    Dim YoBusco As String
    YoBusco = String(6, Chr(TeclaIZQ))
    If UCase(LastTeclas) = UCase(YoBusco) Then
        Command6_Click
    End If
End Sub

Private Sub Form_Load()
    LastTeclas = "??????"
       
    If FSO.FileExists(GPF("origs")) = False Then
        'ESCRIBIRLO!!!
        EscribirArch1Linea GPF("origs"), AP + "discos"
    End If
    
    'para recuperaciones
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    IDIOMA = LeerConfig("Idioma", "Español")
    'descomprimir el pakage de imágenes siemrpe que se inicia para evitar
    'violaciones. La version exclusiva puede ser un paquete generado especialmente
    'todas se descomprimen a system
    'las imágenes que se necesitan son
    
    'En Frm Reg una chiquita tipo la chica = index = tapa _
        f61.dlw
    'En frmIni: _
        una grande: f52.dlw _
    'En frmIndex se necesita _
        'El fondo grande: f53.dlw
        'El fondo chico de abajo: f55.dlw (para exclusivo el mismo!!!)
        'tbrPassImg: es el mismo f61.dlw !!!
    'en frmTop10-RANK: el mismo f61.dlw
    'En frmSuperLic se necesitan: _
        los 3 archivos de Windows _
        logo.sys = f56.dlw _
        logos.sys = f57.dlw _
        logow.sys = f58.dlw _
        las imagenes del frmINI _
        f52.dlw _
        Imagen del index en tbrPassIMG _
        tapa.jpg = f61.dlw _
        TOP10.jpg = f54.dlw
    
    'además el manual.doc NO VA!!!!!!!! _
        f1ya.nac
        
    Dim JuSe As New clsJuntaSepara
    'leerlo
    JuSe.ReadFile GPF("pdis233")
    'extraer todo en System
    Dim A As Long
    For A = 1 To JuSe.CantArchs
        JuSe.Extract GPF("extr233"), A
    Next
    'cerrar todo
    Set JuSe = Nothing
    
    Image1.Picture = LoadPicture(GPF("extr233_61"))
    
    '------------------------------------------------------
    'dejar cragado el frmVideo
    Load frmVIDEO
    'ubicarlo joia para ir mostrando cosas por fuera
    frmVIDEO.Left = Screen.Width
    frmVIDEO.Width = Screen.Width
    frmVIDEO.Top = 0
    frmVIDEO.Height = Screen.Height
    frmVIDEO.Show
    frmVIDEO.Refresh
    
    frmVIDEO.picVideo.Left = 0
    frmVIDEO.picVideo.Top = 0
    frmVIDEO.picVideo.Width = frmVIDEO.Width
    frmVIDEO.picVideo.Height = frmVIDEO.Height
    frmVIDEO.picVideo.Visible = False
    '------------------------------------------------------
    
    If frmVIDEO.Left = Screen.Width Then
        TvOn = 1
    Else
        TvOn = 0
    End If
    
    AjustarFRM Me, 12000
    'se graba en win y system
    If UCase(App.EXEName) <> "3PM" Then
        MsgBox "No puede cambiar el nombre del programa"
        End
    End If
    'VER SI existe el archivo con los datos de las
    'imágenes de inicio y de cierre
    Dim ArchImgIni As String
    ArchImgIni = GPF("iit17222")
    'este archivo de inicio se genera la primera vez para tomas las imagenes de windows
    'al momento de instalar 3PM
    If FSO.FileExists(ArchImgIni) Then
        GoTo YaEstaIMG
    Else
        Set TE = FSO.CreateTextFile(ArchImgIni, True)
        'ver imagen de inicio
        If FSO.FileExists("c:\logo.sys") Then
            TE.WriteLine "ImgIni=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(GPF("ildw9m")) Then
                FSO.DeleteFile GPF("ildw9m"), True
            End If
            FSO.CopyFile "c:\logo.sys", GPF("ildw9m"), True
        Else
            TE.WriteLine "ImgIni=0"
        End If
        
        'ver imagen de cerrando
        If FSO.FileExists(WINfolder + "logow.sys") Then
            TE.WriteLine "ImgCerrando=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(GPF("ildw9m3")) Then FSO.DeleteFile GPF("ildw9m3"), True
            FSO.CopyFile WINfolder + "logow.sys", GPF("ildw9m3"), True
        Else
            TE.WriteLine "ImgCerrando=0"
        End If
        
        'ver imagen de apagar
        If FSO.FileExists(WINfolder + "logos.sys") Then
            TE.WriteLine "ImgApagar=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(GPF("ildw9m2")) Then FSO.DeleteFile GPF("ildw9m2"), True
            FSO.CopyFile WINfolder + "logos.sys", GPF("ildw9m2"), True
        Else
            TE.WriteLine "ImgApagar=0"
        End If
        'escribir que todas las imagenes se cargan desde windows
        TE.WriteLine "LoadImgIni=w"
        TE.WriteLine "LoadImgCerrando=w"
        TE.WriteLine "LoadImgApagar=w"
        TE.Close
    End If
    
YaEstaIMG:
    'VER si ya se pasaron las imagenes de 3pm
    'a la carpeta que corresponde
    
    'copiar a la carpeta primero la original....
    If FSO.FileExists(GPF("extr233_56")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm")) Then FSO.DeleteFile GPF("ild3pm"), True
        FSO.CopyFile GPF("extr233_56"), GPF("ild3pm"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If FSO.FileExists(GPF("233_56_b")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm")) Then FSO.DeleteFile GPF("ild3pm"), True
        FSO.CopyFile GPF("233_56_b"), GPF("ild3pm"), True
    End If
    
    'copiar a la carpeta primero la original....
    If FSO.FileExists(GPF("extr233_58")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm3")) Then FSO.DeleteFile GPF("ild3pm3"), True
        FSO.CopyFile GPF("extr233_58"), GPF("ild3pm3"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If FSO.FileExists(GPF("233_58_b")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm3")) Then FSO.DeleteFile GPF("ild3pm3"), True
        FSO.CopyFile GPF("233_58_b"), GPF("ild3pm3"), True
    End If
    
    'copiar a la carpeta primero la original....
    If FSO.FileExists(GPF("extr233_57")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm2")) Then FSO.DeleteFile GPF("ild3pm2"), True
        FSO.CopyFile GPF("extr233_57"), GPF("ild3pm2"), True
    End If
    'que sera reemplazada si existe la de SL.....
    If FSO.FileExists(GPF("233_57_b")) Then
        'siempre copiarlo si esta
        If FSO.FileExists(GPF("ild3pm2")) Then FSO.DeleteFile GPF("ild3pm2"), True
        FSO.CopyFile GPF("233_57_b"), GPF("ild3pm2"), True
    End If
    
    cmbCountry.AddItem "Argentina"
    cmbCountry.AddItem "(SELECCIONE PAIS)"
    cmbCountry.AddItem "España"
    cmbCountry.AddItem "Chile"
    cmbCountry.AddItem "Uruguay"
    cmbCountry.AddItem "Honduras"
    cmbCountry.AddItem "El Salvador"
    cmbCountry.AddItem "Mexico"
    cmbCountry.AddItem "Venezuela"
    cmbCountry.AddItem "Bolivia"
    cmbCountry.AddItem "Colombia"
    cmbCountry.AddItem "Perú"
    cmbCountry.AddItem "Nicaragua"
    cmbCountry.AddItem "EEUU"
    cmbCountry.AddItem "Paraguay"
    cmbCountry.AddItem "Costa Rica"
    cmbCountry.AddItem "Ecuador"
    cmbCountry.AddItem "Guatemala"
    cmbCountry.AddItem "Panama"
    cmbCountry.AddItem "Puerto Rico"
    cmbCountry.AddItem "Republica Dominicana"
    cmbCountry.ListIndex = 0
    
    
    'no se pone el primer componente por que este cambia con cada inicio
    lblGUID = K.UniquePC
        
    TXT = "Bienvenido a 3PM. Gracias por confiar en tbrSoft Argentina" + vbCrLf + vbCrLf + _
    "Puede utilizar esta version demo con algunas restricciones simplemente " + _
    "indicando su pais de residencia y presionando OK ahora" + vbCrLf + vbCrLf + _
    "3PM no incluye en ninguna de sus licencia el derecho de venta del software. Por lo que " + _
    "solo usted tendra una copia LEGAL si compra este software a " + _
    "tbrSoft Argentina." + vbCrLf + _
    "Para adquirir la version definitiva deberá solicitarlo a tbrSoft " + _
    "via email a info@tbrsoft.com o a tbrsoft@hotmail.com (tambien Messenger)"
    
    LBL = TXT
    
    'ver primero quien es para saber si esta habilitado licenciarse
    'si ClaveAdmin = "demo" quiere decir que lo bajo de internet y por
    'lo tanto no puede licenciar NI BOSTA!!!JAJAJAJAJA
    ClaveAdmin = LeerConfig("ClaveAdmin", "ADMIN")
    'ClaveAdmin = "sncMEX098181y"
    'ERO77701192FF / MARC777
    
    Select Case ClaveAdmin
        Case "xx"
            DatosLicencia = "Licencia propiedad de Miguel Angel Cozzi. " + vbCrLf + _
                "Venado Tuerto - Santa Fe - Argentina"
                
    End Select

'roberto cpaz   RCP888
'diego antonio sanchez corr DASC717771090
'Edgar Giovanni Valdez Hernandez GUAT HGVHG34771000
'rio ceballos rioceballos88
'Clifton Forde PANAMA CFP7118820192
'Melina Gieco StaFe MGSF711905621
'Caludia Sala BsAs MRCSR81172660
'Andres Giamello AGBA718829540/AG31
'Alejandro Maltez NICARAGUA AMN5102991732
'Abraham Grenberg Valle Verde SA GUAT AGVVSA8177109
'humberto segundo cruces CHI HSCC719288012
'juan miguel RepDominic JMRD611885094
'Pablo Duvos UY PDUY210098881
'FRANCISCO JAVIER GONZALEZ LAZCANO CHILE FJGLC71625551
'Oscar Figeuroa Martinez Salvador OEFMES810001
'Hector Amigo BsAs HABSAS8281901
'Jesus Alexander SALV JAS7166290011
'william Obando y Francisco Vielman Flores GUAT WOFVFGU918812
'Damian Ostuni BsAs DOBA811726300
'carlos salas JJY CSJJYAR719922
'daniel omar herrera robles chile DOHRCH6199201
'paulo garcia CHILE PGC5220119851
'Jorge Horta CHILE JOCH102881276
'Alejandro Carmona Oliveros Chile ACOCH3217729
'edison ariel caceres chile EACCH81032772
'Daniel Martinez Chicago DMCE183745510
'Wilmer Fidel Marquez Silva WFMSPR2981109
'julio papetti TUCUMAN JPT1077594731
'jorge albin JACP719283001 jorge albin FEDERICODANIEL
'jorge albin RCP888
'Dardo maidana DARDOMAIDANA
'german peier BsAs GPBSAS7812003
'luis iglesias BsAs LIBA6152896R
'eduardo alberti UY EAJCM2987889h
'juan francisco gonzalez COL JFG729432119q
'Mauricio Levuy Sergio Davo MEX MLSD61846362e
'Jesus Andrès Mata Jimenez MEX JAMG67298187r
'Thomas Hernadez MEX THH635478111g
'rolando torres honduras RTMH523142567z
'david gonzalez MEX DGM652253435y
'Jose Juan Martinez Arguello JIMM MEX JJMA81948572y
'Giovanne Barrios MEX GB6156901836y
'tommy corrientes TC194736251438y
'Luis Enrique Ruiz Chaparro MEX LERC8711101yy
'onofre inda OI71909081125y
'efrain solarte COL EFS50091hgyurr
'henry soto ECU HS611ecu119yh
'Alex Herrera COL AHQ54COL52hyy
'VICTOR HUGO DE LA ROSA (de JMFC) vhdlr5001787y"
'Tomas Nuñez Gonzalez    sncMEX098181y
'Mig    uel Angel Santos Hernandez MEX MASH81090011y
'JUAN MARTIN FLORES CRUZ MEX JMFCm6511yyyq
'Mauro Villaroel     MV541CHQ9090Y
'Chirstian Beltra    cb9811191ujY
'Miguel Angel Cozzi  grAS981aATTy6
'Marcos Sepulveda    bsaHH0981AWqQ
'Rigoberto Matamoros - Oscar Otero Cartagena (El Salvador)   fRF4247L000wZ
'Jose Luis Dorado    33Ccq0151AxqF
'Ramon Daniel Cruz   RMLVF00012yqq
'Miguel Angel Cozzi  grAS981aATTy6
'Ivan Vera   LOpaFE1701666
'Carlos Alberto Montaña Alvarado Caa9107g8s811
'Gabriel Pablo de Rosa   AFD076qwnn100
'Santiago Vignolo    LIQ3661SV0909
'Humberto Breton BR7ME2jGtt981
'Juan Carlos Monsegu MONS7111yHu66
'Jorge Andres Gonzales Torres    JAGT61098Saa6
'Hugo Kollman    KOLL717109888
'Eduardo Rodriguez   ERO77701192FF / MARC777
'Victor Rocha    VR541SLP11MEX
'Roberto Hurtado RHUR28177MEXy
'Alberto Devit   AlDe1098MXca5
'German Becley   GerBKL00198AA
'Abelardo Garcia Morales ABG011boCO1ky
'Judith Rodriguez    ROD0906mx143u
'Guillermo Milian    gMIL991Mex199
'Jesus Andres Mata Jimenez MG611mex0909a
'Juan Serano JaS0106uuw103
'Sergio Sosa Mendoza SeSo711922yh6
' Leandro Visciarelli CBA7111levi09
'Eduardo Rodriguez Uruguay ERO77701192FF
'Favio Martinez Gomez COL fa61MG52COL91
'Oscar Armenta Soberanis OAS81090Mx880
'(RIGMAT)Francisco Somoza  FrHN0102099yi
'Ernesto isidro vazquez MEX eiv767611iJAA
'Cesar Gordillo GERSA de CV MEX cg5510978AByR
'Rene escrich SALV REES91210909u
'Allan orante Martinez MEX AOM519090hnYa
'guille2p españa G2Pk9111900ES
' Miguel Angel Santos hernandez dice que le di yo??? ms6511comp9ME
'ES EL MISMO DE MSCOMPU GARKA!!!!!!!!!!!!!!!!
'enrique israel mora suarez EIMS611609yyw
'desiderio meneseres BOL Des911BOL1011
    
    'ver si hay registro
    
    If K.ReleerLICENCIA = BErronea Then
        MsgBox "Existen datos erroneos de la licencia. Si ingresa claves equivocadas o ha" + _
            " reemplazado componentes de su PC debe comunicarse con tbrSoft o su proveedor" + _
            " de 3PM para solucionar este inconveniente"
            Exit Sub
    End If
    
    If K.LICENCIA = aSinCargar Then
        'darle la oportunidad de que cargue algo
        Exit Sub
    End If
    
    If K.LICENCIA = CGratuita Then
        'que siga de largo he ingrese
        'MsgBox "Clave gratuita de 3PM."
    End If
    
    If K.LICENCIA = GFull Then
        If ClaveAdmin = "demo" Then
            'o ha crakeado o todavia no ha instalado la actualizacion
            'que correspónde que le envie si compro
            MsgBox "No es posible licenciar esta versión demo de 3PM" + _
                " descargada de internet. Solicite su actualización o " + _
                "instalador para validar este software como corresponde"
            'NO INGRESAR
            'SI SIGUE DE LARGO ENTRA A 3PM, se queda en bolas
            Exit Sub
        Else
            'que siga de largo he ingrese
            'MsgBox "El codigo se ha cargado correctamente. Bienvenido a 3PM"
        End If
    End If
    
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("61conf")) Then
            Image1.Picture = LoadPicture(GPF("61conf"))
        End If
        If ClaveAdmin = "demo" Then
            'o ha crakeado o todavia no ha instalado la actualizacion
            'que correspónde que le envie si compro
            MsgBox "No es posible licenciar esta versión demo de 3PM" + _
                " descargada de internet. Solicite su actualización o " + _
                "instalador para validar este software como corresponde"
            'NO INGRESAR
            'SI SIGUE DE LARGO ENTRA A 3PM, se queda en bolas
            Exit Sub
        Else
            'que siga de largo he ingrese
            'MsgBox "SUPERLICENCIA de 3PM. El codigo de SuperLicencia se ha cargado correctamente. Bienvenido a Super3PM"
            '-----------------------------------------
            '-----------------------------------------
            Is3pmExclusivo = LeerConfig("3pmExcl", "0")
            '-----------------------------------------
            '-----------------------------------------
        End If
    End If
    On Local Error GoTo noPuede
    Me.Hide
    Me.Refresh
    Unload Me
    frmINI.Show 1
    tERR.Anotar "acpi"
    
    Exit Sub
noPuede:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acpp"
    
    Resume Next
    
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acpj"
    Resume Next
    
End Sub

Private Sub txtCOD_Change(Index As Integer)
    If Index < 7 And Len(txtCOD(Index)) = 5 Then
        'pasar a la casilla siguinete
        txtCOD(Index + 1).SetFocus
    End If
    If Index = 7 And Len(txtCOD(Index)) = 5 Then
        cmbCountry.SetFocus
    End If
End Sub

Private Sub txtCOD_GotFocus(Index As Integer)
    'pintar todo
    txtCOD(Index).SelStart = 0
    txtCOD(Index).SelLength = Len(txtCOD(Index))
End Sub
    
