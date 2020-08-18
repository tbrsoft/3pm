VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigCart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración del carrito de compras"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   10575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkVentaExtras 
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
      Left            =   90
      TabIndex        =   27
      Top             =   2010
      Width           =   195
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Muestras de musica"
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
      Height          =   2505
      Left            =   4980
      TabIndex        =   14
      Top             =   60
      Width           =   5505
      Begin VB.TextBox txtMaxListaTestMusic 
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
         Left            =   4500
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1890
         Width           =   600
      End
      Begin VB.VScrollBar vsMaxListaTestMusic 
         Height          =   330
         LargeChange     =   10
         Left            =   5130
         Max             =   0
         Min             =   400
         TabIndex        =   24
         Top             =   1890
         Width           =   330
      End
      Begin VB.TextBox txtCreditForTestMusic 
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
         Left            =   4500
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditForTestMusic 
         Height          =   330
         LargeChange     =   10
         Left            =   5130
         Max             =   0
         Min             =   50
         TabIndex        =   20
         Top             =   1200
         Width           =   330
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
         Left            =   150
         TabIndex        =   16
         Top             =   780
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
         Left            =   150
         TabIndex        =   15
         Top             =   330
         Width           =   195
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "canciones"
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
         Index           =   3
         Left            =   4530
         TabIndex        =   26
         Top             =   2190
         Width           =   885
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar el máximo de canciones que pueden quedar en lista (cero permite todo)"
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
         Height          =   465
         Left            =   60
         TabIndex        =   23
         Top             =   1860
         Width           =   4425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "créditos"
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
         Index           =   2
         Left            =   4500
         TabIndex        =   22
         Top             =   1530
         Width           =   705
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Exigir que haya créditos cargados para permitir muestras musicales (incluye ringtones)"
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
         Height          =   585
         Left            =   90
         TabIndex        =   19
         Top             =   1170
         Width           =   4365
      End
      Begin VB.Label Label2 
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
         Left            =   420
         TabIndex        =   18
         Top             =   330
         Width           =   5055
      End
      Begin VB.Label Label3 
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
         Height          =   435
         Left            =   420
         TabIndex        =   17
         Top             =   690
         Width           =   4965
      End
   End
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
      Left            =   330
      TabIndex        =   13
      Top             =   1590
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
      Left            =   330
      TabIndex        =   11
      Top             =   1260
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
      Left            =   330
      TabIndex        =   9
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
      Left            =   330
      TabIndex        =   7
      Top             =   900
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
      Left            =   210
      TabIndex        =   4
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
      Left            =   300
      TabIndex        =   2
      Top             =   2820
      Visible         =   0   'False
      Width           =   195
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   660
      Left            =   5670
      TabIndex        =   0
      Top             =   2640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar todo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   660
      Left            =   6870
      TabIndex        =   1
      Top             =   2640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton9 
      Height          =   660
      Left            =   3840
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   4210752
      fCapt           =   "GRABAR KARAOKES"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   4210816
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   660
      Left            =   8310
      TabIndex        =   31
      Top             =   2640
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Configurar precios venta de multimedia"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   270
      X2              =   4290
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpapers, ringtones, juegos y aplicaciones JAVA"
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
      Height          =   435
      Left            =   510
      TabIndex        =   29
      Top             =   2220
      Width           =   3150
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar la venta de extras"
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
      Left            =   390
      TabIndex        =   28
      Top             =   1980
      Width           =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   210
      X2              =   4230
      Y1              =   2730
      Y2              =   2730
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
      BackStyle       =   0  'Transparent
      Caption         =   "Activar Grabación en CD (requiere nero 7 o superior)"
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
      Height          =   435
      Left            =   600
      TabIndex        =   12
      Top             =   1500
      Width           =   4125
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
      Left            =   570
      TabIndex        =   10
      Top             =   1260
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
      Left            =   570
      TabIndex        =   8
      Top             =   510
      Width           =   2295
   End
   Begin VB.Label Label6 
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
      Height          =   525
      Left            =   570
      TabIndex        =   6
      Top             =   780
      Width           =   4305
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Habilitar expendedor."
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
      TabIndex        =   5
      Top             =   90
      Width           =   3720
   End
   Begin VB.Label Label4 
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
      Height          =   615
      Left            =   570
      TabIndex        =   3
      Top             =   2820
      Visible         =   0   'False
      Width           =   4230
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

Private Sub chkShowDemoMusic_Click()
    UpCHKS
End Sub

Private Sub chkVendoMusica_Click()
    UpCHKS
End Sub

Private Sub fBoton1_Click()

    'grabar las configuracviones basicas de carrito
    ChangeConfig "VendoMusica", CStr(chkVendoMusica)
    ChangeConfig "NOMUSIC", CStr(chkNOMUSIC)
    ChangeConfig "ShowDemoMusic", CStr(chkShowDemoMusic)
    ChangeConfig "SaveCart", CStr(chkSaveCart)
    
    ChangeConfig "VentaExtras", CStr(chkVentaExtras)
    
    ChangeConfig "TengoUsb", CStr(Abs(chkTengoUSB))
    ChangeConfig "TengoBluetooth", CStr(Abs(chkTengoBluetooth))
    ChangeConfig "TengoInfra", CStr(Abs(chkTengoInfra))
    ChangeConfig "TengoCD", CStr(Abs(chkTengoCD))

    ChangeConfig "CreditForTestMusic", txtCreditForTestMusic.tExt
    ChangeConfig "MaxListaTestMusic", txtMaxListaTestMusic.tExt

    VendoMusica = chkVendoMusica
    NOMUSIC = chkNOMUSIC
    ShowDemoMusic = chkShowDemoMusic
    SaveCart = chkSaveCart
    'los que estan comentados no les cambio por que en YaCerrar3PM necesito saber tal como estaban al inicio
    'TengoBluetooth = chkTengoBluetooth
    VentaExtras = chkVentaExtras
    TengoUSB = chkTengoUSB
    'TengoCD = chkTengoCD
    
    CreditForTestMusic = vsCreditForTestMusic.Value
    MaxListaTestMusic = vsMaxListaTestMusic.Value

    Unload Me
End Sub

Private Sub fBoton2_Click()
    frmConfigCartPrecios.Show 1
End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub


Private Sub fBoton9_Click()
    Unload Me
    frmConfigGrabarKar.Show 1
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    'cargar la configuracion
    chkVendoMusica.Value = CLng(LeerConfig("VendoMusica", "0"))
    chkNOMUSIC.Value = CLng(LeerConfig("NOMUSIC", "0"))
    chkShowDemoMusic = CLng(LeerConfig("ShowDemoMusic", "0"))
    chkSaveCart = CLng(LeerConfig("SaveCart", "0"))
    
    chkVentaExtras.Value = CLng(LeerConfig("VentaExtras", "0"))
    
    chkTengoUSB.Value = CLng(LeerConfig("TengoUsb", "1"))
    chkTengoBluetooth.Value = CLng(LeerConfig("TengoBluetooth", "0"))
    chkTengoInfra.Value = CLng(LeerConfig("TengoInfra", "0"))
    chkTengoCD.Value = CLng(LeerConfig("TengoCD", "0"))
    
    vsCreditForTestMusic.Value = CLng(LeerConfig("CreditForTestMusic", "0"))
    vsMaxListaTestMusic.Value = CLng(LeerConfig("MaxListaTestMusic", "0"))
    
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
    
    vsCreditForTestMusic.Enabled = (chkShowDemoMusic.Value <> 0) And (chkShowDemoMusic.Enabled)
        txtCreditForTestMusic.Enabled = vsCreditForTestMusic.Enabled
        vsMaxListaTestMusic.Enabled = vsCreditForTestMusic.Enabled
        txtMaxListaTestMusic.Enabled = vsCreditForTestMusic.Enabled
    
    If (chkShowDemoMusic.Value = 0) Or (chkShowDemoMusic.Enabled = False) Then
        Label10.ForeColor = &H808080
        Label11.ForeColor = &H808080
    Else
        Label10.ForeColor = vbWhite
        Label11.ForeColor = vbWhite
    End If
    
    'ocultar los dependedientes de vender musica
    chkTengoBluetooth.Enabled = (chkVendoMusica.Value = 1)
    If chkVendoMusica.Value = 1 Then
        Label6.ForeColor = vbWhite
    Else
        Label6.ForeColor = &H808080
    End If
    
    chkTengoCD.Enabled = (chkVendoMusica.Value = 1)
    If chkVendoMusica.Value = 1 Then
        Label9.ForeColor = vbWhite
    Else
        Label9.ForeColor = &H808080
    End If
    
    chkTengoUSB.Enabled = (chkVendoMusica.Value = 1)
    If chkVendoMusica.Value = 1 Then
        Label7.ForeColor = vbWhite
    Else
        Label7.ForeColor = &H808080
    End If
    
    chkTengoInfra.Enabled = False
    Label8.ForeColor = &H808080
    
End Sub

Private Sub vsCreditForTestMusic_Change()
    txtCreditForTestMusic.tExt = vsCreditForTestMusic.Value
End Sub

Private Sub vsMaxListaTestMusic_Change()
    txtMaxListaTestMusic.tExt = vsMaxListaTestMusic.Value
End Sub
