VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00000000&
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
   Begin VB.CommandButton Command8 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3930
      Picture         =   "frmREG.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4500
      Width           =   645
   End
   Begin VB.TextBox LBL 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   3045
      Left            =   5190
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   90
      Width           =   5865
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3960
      Picture         =   "frmREG.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7410
      Width           =   645
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3960
      Picture         =   "frmREG.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5340
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3930
      Picture         =   "frmREG.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3810
      Width           =   645
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia de Usuario Final"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   345
      Left            =   4680
      TabIndex        =   3
      Top             =   6900
      Value           =   1  'Checked
      Width           =   7035
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3960
      Picture         =   "frmREG.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6750
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3960
      Picture         =   "frmREG.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5970
      Width           =   645
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3930
      Picture         =   "frmREG.frx":123C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3090
      Width           =   645
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cargar archivo de licencia"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   6
      Left            =   4710
      TabIndex        =   15
      Top             =   4650
      Width           =   4005
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   1
      X1              =   4200
      X2              =   11610
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      Index           =   0
      X1              =   4170
      X2              =   11580
      Y1              =   5220
      Y2              =   5220
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recuperar/Reparar (tecla Izquierda 6 veces)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   5
      Left            =   4680
      TabIndex        =   12
      Top             =   7440
      Width           =   6525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ver contrato de licencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   4
      Left            =   4710
      TabIndex        =   11
      Top             =   6720
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPRAR AHORA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   3
      Left            =   4710
      TabIndex        =   10
      Top             =   5490
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   2
      Left            =   4710
      TabIndex        =   9
      Top             =   6060
      Width           =   2745
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INICIAR PROGRAMA"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Index           =   1
      Left            =   4710
      TabIndex        =   8
      Top             =   3240
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Abrir Manual de uso. RECOMENDADO si es su primer uso"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   555
      Index           =   0
      Left            =   4680
      TabIndex        =   7
      Top             =   3780
      Width           =   5985
   End
   Begin VB.Image Image1 
      Height          =   7800
      Left            =   480
      Picture         =   "frmREG.frx":1546
      Stretch         =   -1  'True
      Top             =   270
      Width           =   4185
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

   'dar ingreso a la clave y la grabo solo aqui
    K.IngresaClave 'aqui se carga mLicencia !!!
        
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
    
    If K.LICENCIA = CGratuita Then MsgBox "Clave gratuita de 3PM. "
    
    If K.LICENCIA = GFull Then MsgBox "El codigo se ha cargado correctamente. Bienvenido a 3PM "
    
    If K.LICENCIA = HSuperLicencia Then MsgBox "SUPERLICENCIA de 3PM. El codigo de SuperLicencia se ha cargado correctamente. Bienvenido a Super3PM"
    
    Unload Me
    frmINI.Show 1
                    
    tERR.Anotar "acpg"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acph"
    Resume Next
                    
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
    'ejecutar el sistema de compras en AP
    
    'frmCompraYA.Show 1
End Sub

Private Sub Command6_Click()
    A = Shell(AP + "repair.exe", vbNormalFocus)
    'MsgBox A
    Unload Me
    End
End Sub

'Private Sub Command7_Click()
'    K.CrearClave
'End Sub

Private Sub Command8_Click()
    'leer algun archivo de licecnia
    Dim CM As New CommonDialog
    CM.DialogTitle = "Cargar licencia de 3PM v7 ..."
    CM.Filter = "Licencia de 3PM v7 (*.L37)|*.L37"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'ponerlo como original ...
    FSO.CopyFile F, GPF("cd5pm"), True
    ' y como copia ...
    FSO.CopyFile F, GPF("cd6pm"), True
    
    K.IngresaClave
    
    'apretar el boton iniciar programa
    Command1_Click
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
        
    Dim JuSe As New tbrJUSE.clsJUSE
    'leerlo
    JuSe.ReadFile GPF("pdis233")
    'extraer todo en System
    Dim A As Long
    For A = 1 To JuSe.CantArchs
        JuSe.Extract GPF("extr233"), A
    Next
    'cerrar todo
    Set JuSe = Nothing
    
    Image1.Picture = LoadPicture(GPF("extr233_62"))
    
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
        
    TXT = "Bienvenido a 3PM." + vbCrLf + "Gracias por confiar en tbrSoft Argentina" + vbCrLf + vbCrLf + _
    "Puede utilizar esta version demo con algunas restricciones simplemente " + _
    "presionando 'Iniciar programa' ahora" + vbCrLf + vbCrLf + _
    "Si desea adquirir definitivamente este software presione el boton " + _
    "'COMPRAR AHORA' o siga los pasos indicados " + _
    "en la herramienta creada para este fin en Inicio/Programas/tbrSoft/3PM/Licencia" + vbCrLf + vbCrLf + _
    "Si desea quitar esta pantalla de bienvenida y otras limitaciones " + _
    "puede obtener una clave gratuita utilizando la misma herramienta de compra" + vbCrLf + vbCrLf + _
    "Si esta PC ya contaba con licencia de 3PM la funcion de 'COMPRAR LICENCIA' " + _
    "lo resolvera." + vbCrLf + vbCrLf + _
    "Si ya ha adquirido y dispone de su archivo de licencia use la opción" + _
    "'Cargar archivo de licencia'" + vbCrLf + vbCrLf + _
    "Cualquier duda envie un email a info@tbrsoft.com"
    
    
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
    
    'ver si hay registro
    
    If K.LICENCIA = BErronea Then
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
