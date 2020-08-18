VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigCart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración del carrito de compras"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkTengoCD 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   480
      TabIndex        =   36
      Top             =   1320
      Width           =   195
   End
   Begin VB.CheckBox chkTengoInfra 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   480
      TabIndex        =   34
      Top             =   1050
      Width           =   195
   End
   Begin VB.CheckBox chkTengoUSB 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   480
      TabIndex        =   32
      Top             =   510
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chkTengoBluetooth 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   480
      TabIndex        =   30
      Top             =   780
      Width           =   195
   End
   Begin VB.CheckBox chkVendoMusica 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   180
      TabIndex        =   27
      Top             =   120
      Width           =   195
   End
   Begin VB.CheckBox chkSaveCart 
      BackColor       =   &H00533422&
      Caption         =   "Conservar el carrito de compras elegido al cerrar 3PM (persistente a cortes de luz)."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   180
      TabIndex        =   18
      Top             =   2250
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkNOMUSIC 
      BackColor       =   &H00533422&
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   480
      TabIndex        =   15
      Top             =   1680
      Width           =   195
   End
   Begin VB.CheckBox chkShowDemoMusic 
      BackColor       =   &H00533422&
      Caption         =   "Hacer una muestra de 20 segundos de canciones al presionar 'OK'."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   210
      Left            =   750
      TabIndex        =   14
      Top             =   1980
      Width           =   195
   End
   Begin VB.Frame frSELL 
      BackColor       =   &H00000000&
      Caption         =   "Precios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4995
      Left            =   150
      TabIndex        =   0
      Top             =   2520
      Width           =   9975
      Begin VB.ListBox lstPromosV 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1020
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   19
         Top             =   3840
         Width           =   8595
      End
      Begin VB.VScrollBar vsPrecioTotalCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   8460
         Max             =   0
         Min             =   400
         TabIndex        =   6
         Top             =   1185
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioTotalCarrito 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   7860
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1185
         Width           =   600
      End
      Begin VB.VScrollBar vsCantFileCarrito 
         Height          =   330
         LargeChange     =   10
         Left            =   5280
         Max             =   0
         Min             =   100
         TabIndex        =   4
         Top             =   1230
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCantFileCarrito 
         Alignment       =   2  'Center
         Enabled         =   0   'False
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
         Left            =   4680
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1230
         Width           =   600
      End
      Begin VB.ComboBox cmbTipoFile 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmConfigCart.frx":0000
         Left            =   120
         List            =   "frmConfigCart.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2985
      End
      Begin VB.ListBox lstPROMOSA 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1080
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   1
         Top             =   2460
         Width           =   8595
      End
      Begin tbrFaroButton.fBoton fBoton2 
         Height          =   540
         Left            =   8190
         TabIndex        =   7
         Top             =   1710
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   953
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Agregar esta promoción"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton3 
         Height          =   1140
         Left            =   8730
         TabIndex        =   8
         Top             =   2430
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   2011
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Eliminar promoción elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton5 
         Height          =   1050
         Left            =   8730
         TabIndex        =   20
         Top             =   3840
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1852
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Eliminar promoción elegida"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label lblDetPrec 
         BackColor       =   &H00C0C0C0&
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   1740
         Width           =   8055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS DE VIDEO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   22
         Top             =   3660
         Width           =   5625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRECIOS DE AUDIO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   2280
         Width           =   6345
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Se toma como referencia los créditos por señal de monedero especificado en la sección 'Créditos' de esta configuración."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Index           =   19
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   9795
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Puede agregar la cantidad de promociones que desee. Asegúrese que sean coherentes."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   12
         Top             =   810
         Width           =   9765
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de archivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   21
         Left            =   90
         TabIndex        =   11
         Top             =   1080
         Width           =   2985
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad de archivos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   495
         Index           =   22
         Left            =   3180
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Precio total en créditos"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   435
         Index           =   23
         Left            =   6150
         TabIndex        =   9
         Top             =   1140
         Width           =   1665
      End
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   450
      Left            =   8670
      TabIndex        =   16
      Top             =   7590
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   794
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   450
      Left            =   6120
      TabIndex        =   17
      Top             =   7590
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   794
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   210
      X2              =   4230
      Y1              =   1590
      Y2              =   1590
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   210
      X2              =   4230
      Y1              =   420
      Y2              =   420
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar Grabación en CD"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   35
      Top             =   1320
      Width           =   2370
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar Infrarrojos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   33
      Top             =   1050
      Width           =   1860
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar USB automático"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   31
      Top             =   510
      Width           =   2295
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Activar Bluetooth (solo dispositivos compatibles con BlueSoleil 1.6.4 o superior)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   29
      Top             =   780
      Width           =   7845
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar la venta de música."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   450
      TabIndex        =   28
      Top             =   90
      Width           =   3720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Conservar el carrito de compras elegido al cerrar 3PM (persistente a cortes de luz)."
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      TabIndex        =   26
      Top             =   2250
      Visible         =   0   'False
      Width           =   8220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hacer una muestra de 20 segundos de canciones al presionar 'OK'."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1020
      TabIndex        =   25
      Top             =   1980
      Width           =   6525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NO REPRODUCIR MÚSICA. ESTE EQUIPO SOLO VENDE."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   750
      TabIndex        =   24
      Top             =   1680
      Width           =   5055
   End
End
Attribute VB_Name = "frmConfigCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkNOMUSIC_Click()
    UpCHKS
End Sub

Private Sub chkVendoMusica_Click()
    UpCHKS
End Sub

Private Sub cmbTipoFile_Click()
    UpDesc
End Sub

Private Sub fBoton1_Click()

    'grabar las configuracviones basicas de carrito
    ChangeConfig "VendoMusica", CStr(chkVendoMusica)
    ChangeConfig "NOMUSIC", CStr(chkNOMUSIC)
    ChangeConfig "ShowDemoMusic", CStr(chkShowDemoMusic)
    ChangeConfig "SaveCart", CStr(chkSaveCart)
    
    ChangeConfig "TengoUsb", CStr(Abs(chkTengoUSB))
    ChangeConfig "TengoBluetooth", CStr(Abs(chkTengoBluetooth))
    ChangeConfig "TengoInfra", CStr(Abs(chkTengoInfra))
    ChangeConfig "TengoCD", CStr(Abs(chkTengoCD))

    VendoMusica = chkVendoMusica
    NOMUSIC = chkNOMUSIC
    ShowDemoMusic = chkShowDemoMusic
    SaveCart = chkSaveCart
    TengoBluetooth = chkTengoBluetooth
    
    Carrito.SavePrices GPF("promocart")
    Unload Me
End Sub

Private Sub fBoton2_Click()
    If vsCantFileCarrito.Value = 0 Then
        MsgBox "No puede especificar el precio para cero selecciones!"
        Exit Sub
    End If
    
    If cmbTipoFile.ListIndex = 0 Then 'ES AUDIO
        Carrito.SetPricesAudio vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    If cmbTipoFile.ListIndex = 1 Then 'ES VIDEO
        Carrito.SetPricesVideo vsCantFileCarrito.Value, vsPrecioTotalCarrito.Value
    End If
    
    VerPromos
End Sub

Private Sub fBoton3_Click()
    
    If lstPROMOSA.ListIndex = -1 Then Exit Sub
    
    'si solo queda una no se puede eliminar
    If lstPROMOSA.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPROMOSA
    End If
End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub

Private Sub fBoton5_Click()
    If lstPromosV.ListIndex = -1 Then Exit Sub
    
    If lstPromosV.ListCount = 1 Then
        MsgBox "No puede dejar sin promociones!"
    Else
        KillPromo lstPromosV
    End If
    
End Sub

Private Sub KillPromo(S As String)

    'leer que es lo que se quiere eliminar
    Dim SP() As String, TP As Long, Cant As Long
    SP = Split(S)
    TP = CLng(SP(0))
    Cant = CLng(SP(2))
    
    If Cant = 1 Then
        MsgBox "No puede borrar el precio de una cancion!"
        Exit Sub
    End If
    
    If TP = 1 Then ' ES MUSICA
        Carrito.KillPricesAudioBase Cant
    End If
    
    If TP = 2 Then ' ES MUSICA
        Carrito.KillPricesVideoBase Cant
    End If
    
    VerPromos
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    'cargar la configuracion
    chkVendoMusica.Value = CLng(LeerConfig("VendoMusica", "0"))
    chkNOMUSIC.Value = CLng(LeerConfig("NOMUSIC", "0"))
    chkShowDemoMusic = CLng(LeerConfig("ShowDemoMusic", "0"))
    chkSaveCart = CLng(LeerConfig("SaveCart", "0"))
    
    chkTengoUSB.Value = CLng(LeerConfig("TengoUsb", "1"))
    chkTengoBluetooth.Value = CLng(LeerConfig("TengoBluetooth", "0"))
    chkTengoInfra.Value = CLng(LeerConfig("TengoInfra", "0"))
    chkTengoCD.Value = CLng(LeerConfig("TengoCD", "0"))
    'xxxxxx poner todas estas en todos los lugares que correspondan!!!!
        
    VerPromos
    
    cmbTipoFile.Clear
    cmbTipoFile.AddItem "AUDIO"
    cmbTipoFile.AddItem "VIDEO"
    cmbTipoFile.ListIndex = 0
    
    vsCantFileCarrito.Value = 0
    vsPrecioTotalCarrito.Value = 0
    
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    
    UpCHKS
End Sub

Private Sub UpCHKS()

    If chkVendoMusica.Value = 0 Then
        chkNOMUSIC.Value = 0
        chkShowDemoMusic.Value = 0
    End If
    
    chkNOMUSIC.Enabled = (chkVendoMusica.Value = 1)
    chkShowDemoMusic.Enabled = (chkNOMUSIC.Value = 1) And (chkVendoMusica.Value = 1)
    'chkSaveCart.Enabled = (chkVendoMusica.Value = 1)
    
    If chkNOMUSIC.Enabled Then
        Label2.ForeColor = vbWhite
    Else
        Label2.ForeColor = &H808080
    End If
    
    If chkShowDemoMusic.Enabled Then
        Label3.ForeColor = vbWhite
    Else
        Label3.ForeColor = &H808080
    End If
    
    chkTengoUSB.Enabled = False
    Label7.ForeColor = &H808080
    
    chkTengoBluetooth.Enabled = (chkVendoMusica.Value = 1)
    If chkVendoMusica.Value = 1 Then
        Label6.ForeColor = vbWhite
    Else
        Label6.ForeColor = &H808080
    End If
    
    chkTengoInfra.Enabled = False
    Label8.ForeColor = &H808080
    
    chkTengoCD.Enabled = False
    Label9.ForeColor = &H808080
    
End Sub

Private Sub VerPromos()
    'ver las promociones ya grabadas
    Dim H As Long
    
    lstPROMOSA.Clear
    
    Dim SN As Single
    
    For H = 1 To Carrito.GetTotalPricesAudio
        If Carrito.GetPricesAudioBase(H) > 0 Then
            SN = Round(Carrito.GetPricesAudioBase(H) * PrecioBase / TemasPorCredito, 2)
            
            If H = 1 Then
                lstPROMOSA.AddItem "1 - " + CStr(H) + " fichero de AUDIO " + _
                    " por $ " + CStr(SN)
                    
            Else
                lstPROMOSA.AddItem "1 - " + CStr(H) + " ficheros de AUDIO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
    
    lstPromosV.Clear
    For H = 1 To Carrito.GetTotalPricesVideo
        If Carrito.GetPricesVideoBase(H) > 0 Then
            SN = Round(Carrito.GetPricesVideoBase(H) * PrecioBase / TemasPorCredito, 2)
            If H = 1 Then
                lstPromosV.AddItem "2 - " + CStr(H) + " fichero de VIDEO " + _
                    " por $ " + _
                    CStr(SN)
            Else
                lstPromosV.AddItem "2 - " + CStr(H) + " ficheros de VIDEO " + _
                    " por $ " + CStr(SN) + _
                    " ($ " + CStr(Round(SN / H, 2)) + " cada uno)"
            End If
        End If
    Next H
End Sub

Private Sub vsCantFileCarrito_Change()
    txtCantFileCarrito.Text = vsCantFileCarrito.Value
    UpDesc
End Sub

Private Sub vsPrecioTotalCarrito_Change()
    txtPrecioTotalCarrito.Text = vsPrecioTotalCarrito.Value
    UpDesc
End Sub

Private Sub UpDesc()
    Dim PP As Single
    PP = Round(vsPrecioTotalCarrito.Value * PrecioBase / TemasPorCredito, 2)
    
    Dim TP As String
    If cmbTipoFile.ListIndex = 0 Then TP = "Audio"
    If cmbTipoFile.ListIndex = 1 Then TP = "Video"
    
    If vsCantFileCarrito.Value = 0 Or vsPrecioTotalCarrito.Value = 0 Then
        lblDetPrec.Caption = "Paquete no válido"
        Exit Sub
    End If
    
    If vsCantFileCarrito.Value = 1 Then
        lblDetPrec.Caption = "Paquete de " + CStr(vsCantFileCarrito.Value) + _
            " fichero de " + TP + ": $ " + CStr(PP)
    Else
        lblDetPrec.Caption = "Paquete de " + CStr(vsCantFileCarrito.Value) + _
            " ficheros de " + TP + ": $ " + _
            CStr(PP) + " cada uno $ " + CStr(Round(PP / vsCantFileCarrito.Value, 2))
    End If
End Sub
