VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00404080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de 3PM"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7590
      Width           =   3720
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
      Left            =   2460
      TabIndex        =   19
      Top             =   6480
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
      Left            =   1980
      TabIndex        =   18
      Top             =   5520
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
      Left            =   4770
      TabIndex        =   17
      Top             =   5940
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
      Height          =   1905
      Left            =   510
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Text            =   "frmREG.frx":0000
      Top             =   60
      Width           =   9255
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
      Left            =   5760
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
      Top             =   5010
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
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
      Top             =   3720
      Width           =   1050
   End
   Begin VB.TextBox lblGUID 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Aqui va el codigo"
      Top             =   2760
      Width           =   11535
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
      Top             =   3720
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
      Left            =   3840
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
      Height          =   705
      Index           =   3
      Left            =   510
      TabIndex        =   15
      Top             =   4110
      Width           =   10605
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   10110
      Picture         =   "frmREG.frx":009C
      Stretch         =   -1  'True
      Top             =   390
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
      Top             =   4800
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
      Top             =   3510
      Width           =   11505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":2B97
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
      Height          =   765
      Index           =   0
      Left            =   540
      TabIndex        =   12
      Top             =   1980
      Width           =   9255
   End
End
Attribute VB_Name = "frmREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GGG As String

Dim SysF As Folder 'ubicacion de la carpeta de Windows
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
    
   'dar ingreso a la clave
    K.IngresaClave CodigoUnido
        
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
                    
End Sub

Private Sub Command2_Click()
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

Private Sub Form_Load()
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
    '------------------------------------------------------
    
    
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    AjustarFRM Me, 12000
    'se graba en win y system
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(WINfolder, 1) <> "\" Then WINfolder = WINfolder + "\"
    If Right(SYSfolder, 1) <> "\" Then SYSfolder = SYSfolder + "\"
    
    If UCase(App.EXEName) <> "3PM" Then
        MsgBox "No puede cambiar el nombre del programa"
        End
    End If
    'VER SI existe el archivo con los datos de las
    'imágenes de inicio y de cierre
    Dim ArchImgIni As String
    ArchImgIni = AP + "imgini.tbr"
    If FSO.FileExists(ArchImgIni) Then
        GoTo YaEstaIMG
    Else
        Set TE = FSO.CreateTextFile(ArchImgIni, True)
        If FSO.FolderExists(WINfolder + "img3PM") = False Then FSO.CreateFolder WINfolder + "img3PM"
        If FSO.FolderExists(WINfolder + "img3PM\w") = False Then FSO.CreateFolder WINfolder + "img3pm\w"
        'ver imagen de inicio
        If FSO.FileExists("c:\logo.sys") Then
            TE.WriteLine "ImgIni=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "img3pm\w\logo.sys") Then FSO.DeleteFile WINfolder + "img3pm\w\logo.sys", True
            FSO.CopyFile "c:\logo.sys", WINfolder + "img3pm\w\logo.sys", True
        Else
            TE.WriteLine "ImgIni=0"
        End If
        
        'ver imagen de cerrando
        If FSO.FileExists(WINfolder + "logow.sys") Then
            TE.WriteLine "ImgCerrando=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "img3pm\w\logow.sys") Then FSO.DeleteFile WINfolder + "img3pm\w\logow.sys", True
            FSO.CopyFile WINfolder + "logow.sys", WINfolder + "img3pm\w\logow.sys", True
        Else
            TE.WriteLine "ImgCerrando=0"
        End If
        
        'ver imagen de apagar
        If FSO.FileExists(WINfolder + "logos.sys") Then
            TE.WriteLine "ImgApagar=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "img3pm\w\logos.sys") Then FSO.DeleteFile WINfolder + "img3pm\w\logos.sys", True
            FSO.CopyFile WINfolder + "logos.sys", WINfolder + "img3pm\w\logos.sys", True
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
    If FSO.FolderExists(WINfolder + "img3pm") = False Then FSO.CreateFolder (WINfolder + "img3pm")
    If FSO.FolderExists(WINfolder + "img3pm\3") = False Then FSO.CreateFolder (WINfolder + "img3pm\3")
    If FSO.FileExists(AP + "logo.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "img3pm\3\logo.sys") Then FSO.DeleteFile WINfolder + "img3pm\3\logo.sys", True
        FSO.CopyFile AP + "logo.sys", WINfolder + "img3pm\3\logo.sys", True
        'If FSO.FileExists(WINfolder + "img3pm\3\logo.sys") = False Then FSO.CopyFile AP + "logo.sys", WINfolder + "img3pm\3\logo.sys", True
    End If
    If FSO.FileExists(AP + "logow.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "img3pm\3\logow.sys") Then FSO.DeleteFile WINfolder + "img3pm\3\logow.sys", True
        FSO.CopyFile AP + "logow.sys", WINfolder + "img3pm\3\logow.sys", True
    End If
    If FSO.FileExists(AP + "logos.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "img3pm\3\logos.sys") Then FSO.DeleteFile WINfolder + "img3pm\3\logos.sys", True
        FSO.CopyFile AP + "logos.sys", WINfolder + "img3pm\3\logos.sys", True
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
    "El costo de 3PM con licencia para un equipo es de U$S 75" + vbCrLf + _
    "3PM no incluye en ninguna de sus licencia el derecho de venta del software. Por lo que " + _
    "solo usted tendra una copia LEGAL si compra este software a " + _
    "tbrSoft Argentina." + vbCrLf + _
    "Para adquirir la version definitiva deberá solicitarlo a tbrSoft " + _
    "via email a info@tbrsoft.com o a tbrsoft@hotmail.com (tambien Messenger)"
    
    LBL = TXT
    
    'ver primero quien es para saber si esta habilitado licenciarse
    'si ClaveAdmin = "demo" quiere decir que lo bajo de internet y por
    'lo tanto no puede licenciar NI BOSTA!!!JAJAJAJAJA
    ClaveAdmin = "sncMEX098181y"
    
    Select Case ClaveAdmin
        Case "grAS981aATTy6"
            DatosLicencia = "Licencia propiedad de Miguel Angel Cozzi. " + vbCrLf + _
                "Venado Tuerto - Santa Fe - Argentina"
                
    End Select
'VICTOR HUGO DE LA ROSA (de JMFC) vhdlr5001787y"
'Tomas Nuñez Gonzalez    sncMEX098181y
'Miguel Angel Santos Hernandez MEX MASH81090011y
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
            '-----------------------------------------
            '-----------------------------------------
            'habra algunas que ingresen como exlcusivo!!!!!!!!
            'revise cuales entregue y aparentemente nunca del 1 al 10!!!!!
            If K.VariacionClave <= 10 Then
                Is3pmExclusivo = True
            Else
                Is3pmExclusivo = False
            End If
            '-----------------------------------------
            '-----------------------------------------
            '-----------------------------------------
            '-----------------------------------------
        End If
    End If
    
    Unload Me
    frmINI.Show 1
        
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
