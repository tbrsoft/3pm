VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigCart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración del carrito de compras"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkBloquearTecladosUSB 
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
      Left            =   4380
      TabIndex        =   32
      Top             =   780
      Value           =   1  'Checked
      Width           =   195
   End
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
      Left            =   180
      TabIndex        =   27
      Top             =   2250
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
      Height          =   2685
      Left            =   150
      TabIndex        =   14
      Top             =   3390
      Width           =   10395
      Begin VB.VScrollBar vsMaxMuestrasToAddCredit 
         Height          =   330
         LargeChange     =   10
         Left            =   10020
         Max             =   0
         Min             =   400
         TabIndex        =   35
         Top             =   2130
         Width           =   330
      End
      Begin VB.TextBox txtMaxMuestrasToAddCredit 
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
         Left            =   9390
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2130
         Width           =   600
      End
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
         Left            =   9390
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1560
         Width           =   600
      End
      Begin VB.VScrollBar vsMaxListaTestMusic 
         Height          =   330
         LargeChange     =   10
         Left            =   10020
         Max             =   0
         Min             =   400
         TabIndex        =   24
         Top             =   1560
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
         Left            =   9330
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   960
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditForTestMusic 
         Height          =   330
         LargeChange     =   10
         Left            =   9960
         Max             =   0
         Min             =   50
         TabIndex        =   20
         Top             =   930
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
         Top             =   630
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
         Caption         =   "Esto evita que con una moneda pasen todo el día realizando pruebas"
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
         Index           =   4
         Left            =   2790
         TabIndex        =   39
         Top             =   2370
         Width           =   6555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Esto evita que aquellos que colocan monedas dejen programadas cientos de canciones en lista"
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
         Left            =   660
         TabIndex        =   38
         Top             =   1770
         Width           =   8490
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Esto evita que escuchen pruebas aquellos sin intenciones de comprar"
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
         Left            =   2280
         TabIndex        =   37
         Top             =   1230
         Width           =   6555
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Limitar el maximo de muestras musicales sin que se agregue crédito (cero permite todo)"
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
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2160
         Width           =   9225
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
         Left            =   9450
         TabIndex        =   26
         Top             =   1860
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
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1590
         Width           =   9225
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
         Left            =   9630
         TabIndex        =   22
         Top             =   1290
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
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   1050
         Width           =   9195
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
         Height          =   255
         Left            =   420
         TabIndex        =   17
         Top             =   630
         Width           =   6795
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
      Left            =   300
      TabIndex        =   13
      Top             =   1830
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
      Left            =   300
      TabIndex        =   11
      Top             =   1560
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
      Left            =   240
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
      Left            =   300
      TabIndex        =   7
      Top             =   1290
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
      Left            =   5820
      TabIndex        =   2
      Top             =   6780
      Visible         =   0   'False
      Width           =   195
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   660
      Left            =   630
      TabIndex        =   0
      Top             =   6510
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
      Left            =   1830
      TabIndex        =   1
      Top             =   6510
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
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   660
      Left            =   3270
      TabIndex        =   30
      Top             =   6510
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
   Begin tbrFaroButton.fBoton fBoton3 
      Height          =   660
      Left            =   9630
      TabIndex        =   40
      Top             =   1440
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Revisar JAVA"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfigCart.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   150
      TabIndex        =   33
      Top             =   2670
      Width           =   10410
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bloquear cualquier teclado USB"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4650
      TabIndex        =   31
      Top             =   750
      Width           =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   60
      X2              =   7560
      Y1              =   2130
      Y2              =   2130
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Wallpapers, ringtones, juegos, aplicaciones JAVA, imágenes iso / nrg, temas para móviles, videos 3GP, etc."
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
      Height          =   435
      Left            =   4620
      TabIndex        =   29
      Top             =   2190
      Width           =   5880
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
      ForeColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   390
      TabIndex        =   28
      Top             =   2220
      Width           =   3720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   60
      X2              =   7800
      Y1              =   3240
      Y2              =   3240
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
      Caption         =   "Activar Grabación en CD/DVD (requiere nero 7 o superior)"
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
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1830
      Width           =   7875
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
      Left            =   600
      TabIndex        =   10
      Top             =   1560
      Width           =   1860
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Activar USB automático (si va a exponer una conexion USB al público se reocmienda anular el uso de teclados USB)"
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
      Height          =   495
      Left            =   540
      TabIndex        =   8
      Top             =   450
      Width           =   9840
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
      Height          =   255
      Left            =   600
      TabIndex        =   6
      Top             =   1260
      Width           =   8205
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
      Height          =   315
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
      Height          =   1035
      Left            =   6270
      TabIndex        =   3
      Top             =   6330
      Visible         =   0   'False
      Width           =   2520
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

Private Sub chkTengoUSB_Click()
    If chkTengoUSB.Enabled And chkTengoUSB.Value <> 0 Then
        Label14.ForeColor = vbWhite
        chkBloquearTecladosUSB.Enabled = True
    Else
        Label14.ForeColor = &H808080
        chkBloquearTecladosUSB.Enabled = False
    End If
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
    ChangeConfig "BloquearTecladosUSB", CStr(Abs(chkBloquearTecladosUSB))
    
    ChangeConfig "TengoBluetooth", CStr(Abs(chkTengoBluetooth))
    ChangeConfig "TengoInfra", CStr(Abs(chkTengoInfra))
    ChangeConfig "TengoCD", CStr(Abs(chkTengoCD))

    ChangeConfig "CreditForTestMusic", txtCreditForTestMusic.tExt
    ChangeConfig "MaxListaTestMusic", txtMaxListaTestMusic.tExt
    ChangeConfig "MaxMuestrasToAddCredit", txtMaxMuestrasToAddCredit.tExt

    VendoMusica = chkVendoMusica
    NOMUSIC = chkNOMUSIC
    ShowDemoMusic = chkShowDemoMusic
    SaveCart = chkSaveCart
    'los que estan comentados no les cambio por que en YaCerrar3PM necesito saber tal como estaban al inicio
    'TengoBluetooth = chkTengoBluetooth
    VentaExtras = chkVentaExtras
    TengoUSB = chkTengoUSB
    'TengoCD = chkTengoCD
    'BloquearTecladosUSB=chkBloquearTecladosUSB'no por que si lo desactiva el teclado no se va a desactivar!
    
    CreditForTestMusic = vsCreditForTestMusic.Value
    MaxListaTestMusic = vsMaxListaTestMusic.Value
    MaxMuestrasToAddCredit = vsMaxMuestrasToAddCredit.Value

    Unload Me
End Sub

Private Sub fBoton2_Click()
    frmConfigCartPrecios.Show 1
End Sub

Private Sub fBoton3_Click()
    Unload Me
    frmCheckJAR.Show 1
End Sub

Private Sub fBoton4_Click()
    Unload Me
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
    chkBloquearTecladosUSB.Value = CLng(LeerConfig("BloquearTecladosUSB", "0"))
    
    chkTengoBluetooth.Value = CLng(LeerConfig("TengoBluetooth", "0"))
    chkTengoInfra.Value = CLng(LeerConfig("TengoInfra", "0"))
    chkTengoCD.Value = CLng(LeerConfig("TengoCD", "0"))
    
    vsCreditForTestMusic.Value = CLng(LeerConfig("CreditForTestMusic", "0"))
    vsMaxListaTestMusic.Value = CLng(LeerConfig("MaxListaTestMusic", "0"))
    vsMaxMuestrasToAddCredit.Value = CLng(LeerConfig("MaxMuestrasToAddCredit", "0"))
    
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
        vsMaxMuestrasToAddCredit.Enabled = vsCreditForTestMusic.Enabled
        txtMaxMuestrasToAddCredit.Enabled = vsCreditForTestMusic.Enabled
    
    If (chkShowDemoMusic.Value = 0) Or (chkShowDemoMusic.Enabled = False) Then
        Label10.ForeColor = &H808080
        Label11.ForeColor = &H808080
        Label16.ForeColor = &H808080
    Else
        Label10.ForeColor = vbWhite
        Label11.ForeColor = vbWhite
        Label16.ForeColor = vbWhite
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
    
    If chkTengoUSB.Enabled And chkTengoUSB.Value <> 0 Then
        Label14.ForeColor = vbWhite
        chkBloquearTecladosUSB.Enabled = True
    Else
        Label14.ForeColor = &H808080
        chkBloquearTecladosUSB.Enabled = False
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

Private Sub vsMaxMuestrasToAddCredit_Change()
    txtMaxMuestrasToAddCredit.tExt = vsMaxMuestrasToAddCredit.Value
End Sub
