VERSION 5.00
Begin VB.Form frmIndex 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12045
   Icon            =   "frmINDEX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   315
      Left            =   3510
      ScaleHeight     =   255
      ScaleWidth      =   4380
      TabIndex        =   36
      Top             =   4590
      Visible         =   0   'False
      Width           =   4440
   End
   Begin tbr3pm.VUMeter2 VU21 
      Height          =   1275
      Left            =   30
      TabIndex        =   37
      Top             =   5355
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   2249
      Begin tbr3pm.tbrProgressCircle Prog 
         Height          =   465
         Left            =   765
         TabIndex        =   43
         Top             =   540
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   820
      End
      Begin VB.Image TapaEjecutando 
         BorderStyle     =   1  'Fixed Single
         Height          =   840
         Left            =   5940
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1125
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   840
         Left            =   7110
         Stretch         =   -1  'True
         Top             =   45
         Width           =   1170
      End
      Begin VB.Label lblPrecios2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1 coin = 8 creditos / 8 creditos = 1 tema / 8 creditos = 1 VIDEO"
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
         Height          =   255
         Left            =   1440
         TabIndex        =   42
         Top             =   900
         Width           =   6855
      End
      Begin VB.Label lblPuesto2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rank #888"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   4545
         TabIndex        =   41
         Top             =   45
         Width           =   1365
      End
      Begin VB.Label lblREP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Reproduciendo:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1440
         TabIndex        =   40
         Top             =   45
         Width           =   2970
      End
      Begin VB.Label lblTemaSonando2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sin Reproducci�n actual Sin Reproducci�n actual Sin Reproducci�n actual Sin Reproducci�n actual"
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
         Height          =   570
         Left            =   1440
         TabIndex        =   39
         Top             =   315
         UseMnemonic     =   0   'False
         Width           =   4470
      End
      Begin VB.Label lblCreditos2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos 00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1110
         Left            =   8325
         TabIndex        =   38
         Top             =   45
         Width           =   2175
      End
   End
   Begin VB.Frame frDISCOS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3990
      Left            =   3510
      TabIndex        =   11
      Top             =   90
      Width           =   4485
      Begin VB.Timer Timer1 
         Left            =   180
         Top             =   2610
      End
      Begin VB.Timer Timer3 
         Interval        =   10000
         Left            =   180
         Top             =   2070
      End
      Begin tbr3pm.MP3Play MP3 
         Height          =   1620
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   2858
      End
      Begin VB.PictureBox picFondoDisco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00004080&
         Height          =   3735
         Left            =   45
         ScaleHeight     =   3675
         ScaleWidth      =   4245
         TabIndex        =   0
         Top             =   135
         Width           =   4305
         Begin VB.Label lblDisco 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Complete al menos la primera hoja de discos cargados"
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
            Index           =   0
            Left            =   720
            TabIndex        =   33
            Top             =   2790
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Image TapaCD 
            Height          =   2505
            Index           =   0
            Left            =   675
            Stretch         =   -1  'True
            Top             =   270
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Shape lblSel 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   6
            Height          =   555
            Left            =   405
            Shape           =   4  'Rounded Rectangle
            Top             =   1935
            Width           =   435
         End
      End
   End
   Begin tbr3pm.VUMeter VU1 
      Height          =   4425
      Left            =   30
      TabIndex        =   7
      Top             =   90
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   7805
   End
   Begin VB.PictureBox picFondo 
      AutoSize        =   -1  'True
      Height          =   4620
      Left            =   0
      ScaleHeight     =   4560
      ScaleWidth      =   15360
      TabIndex        =   13
      Top             =   6660
      Width           =   15420
      Begin tbr3pm.tbrPassImg tbrPassImg1 
         Height          =   2250
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   3969
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VERSION DEMOSTRATIVA. PROHIBIDA SU UTILIZACION. tbrSoft Argentina info@tbrsoft.com www.tbrsoft.com"
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
            Height          =   1665
            Left            =   90
            TabIndex        =   35
            Top             =   315
            Visible         =   0   'False
            Width           =   2085
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Touch"
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
         Height          =   2220
         Left            =   9780
         TabIndex        =   14
         Top             =   45
         Width           =   2200
         Begin VB.CommandButton cmdPagAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":0442
            Height          =   710
            Left            =   1140
            Picture         =   "frmINDEX.frx":1403
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1485
            Width           =   1000
         End
         Begin VB.CommandButton cmdDiscoAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":1E10
            Height          =   710
            Left            =   1110
            Picture         =   "frmINDEX.frx":2B0D
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   750
            Width           =   1000
         End
         Begin VB.CommandButton cmdDiscoAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":33E5
            Height          =   710
            Left            =   90
            Picture         =   "frmINDEX.frx":4157
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   750
            Width           =   1000
         End
         Begin VB.CommandButton cmdPagAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":4A9A
            Height          =   710
            Left            =   90
            Picture         =   "frmINDEX.frx":5AF9
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1485
            Width           =   1000
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
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
            Height          =   525
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   210
            Width           =   2030
         End
      End
      Begin VB.Label lblV 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "version 8.88.888"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   2340
         TabIndex        =   31
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lstProximos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sin Reproducci�n actual"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   945
         Left            =   4290
         TabIndex        =   32
         Top             =   510
         UseMnemonic     =   0   'False
         Width           =   5445
      End
      Begin VB.Label lblTBR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2265
         TabIndex        =   30
         Top             =   1740
         Width           =   7500
      End
      Begin VB.Label lblDEMO 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Solicite la version definitiva a info@tbrsoft.com / avazquez@cpcipc.org"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   2250
         TabIndex        =   29
         Top             =   2025
         UseMnemonic     =   0   'False
         Width           =   7470
      End
      Begin VB.Label lblTOTdiscos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         Caption         =   "Disco 188 de 188"
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
         Height          =   255
         Left            =   5325
         TabIndex        =   28
         Top             =   1470
         Width           =   2415
      End
      Begin VB.Label lblTiempoRestante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         Caption         =   "Falta: 00:00"
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
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   1470
         Width           =   1545
      End
      Begin VB.Label lblPag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         Caption         =   "Pagina 88 de 88"
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
         Height          =   255
         Left            =   7800
         TabIndex        =   26
         Top             =   1470
         Width           =   1980
      End
      Begin VB.Label lblTemaSonando 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sin Reproducci�n actual"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2280
         TabIndex        =   25
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   7485
      End
      Begin VB.Label lblCreditos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Creditos 00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   2280
         TabIndex        =   24
         Top             =   510
         Width           =   1995
      End
      Begin VB.Label lblPrecios 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1 coin=3 creditos  2 creditos= 1 tema 3 creditos= 1 VIDEO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   2250
         TabIndex        =   23
         Top             =   870
         Width           =   2040
      End
      Begin VB.Label lblPuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         Caption         =   "Rank #888"
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
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   1470
         Width           =   1455
      End
      Begin VB.Label lblValidar 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Validar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6390
         TabIndex        =   21
         Top             =   90
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label LBLpORCtEMA 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   120
         Left            =   2250
         TabIndex        =   20
         Top             =   360
         Width           =   3015
      End
   End
   Begin VB.Frame frModoVideo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   10290
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1715
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nombre del artista - nombre del disco"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.Frame frTEMAS 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   10290
      TabIndex        =   4
      Top             =   2280
      Visible         =   0   'False
      Width           =   1715
      Begin VB.Label T 
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         Caption         =   "Nombre del TEMA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   5
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Label lblTECLAS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "11111222223333344444"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   6060
      TabIndex        =   10
      Top             =   645
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblNoTecla 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   6060
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblNoUSO 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   6765
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lblTEMAS 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Temas del disco elegido"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   10260
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   1710
   End
   Begin VB.Label lblModoVideo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Discos en Modo Video"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   10260
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ModoVideoSelTema As Boolean 'si estoy en video
'saber si estoy eligiendo tema. Sino estoy en disco

Dim TemaElegidoModoVideo As Integer

Dim LastDiscoSel As Long
Dim DiscosEnPagina As Long

Dim VolBajando As Double 'bajando volumen para terminar tema demo
Dim LastpSeconds As Long 'comparador para bajar de a uno el volumen en demos

Dim Ancho As Long, Variacion As Long 'PARA la barra de proceso del tema
Public DuracionTema As Long 'duracion de todos los tenmas de un disco
Dim TotalTema As Long 'duracion total
Dim nDiscoSEL As Long 'del 0 al 5 o hasta donde coresponda!!

Private Function EnQueFilaEstoy() As Long
    'es la fila uno si es la primera
    'la baarra invertida devuelve solo la parte entera!!!
    EnQueFilaEstoy = (nDiscoSEL \ TapasMostradasH) + 1
End Function

Private Sub cmdDiscoAd_Click()
    If MostrarTouch Then
        LineaError = "000-0001"
        Form_KeyDown TeclaDER, 0
        LineaError = "000-0002"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    LineaError = "000-0003"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdDiscoAt_Click()
    If MostrarTouch Then
        LineaError = "000-0004"
        Form_KeyDown TeclaIZQ, 0
        LineaError = "000-0005"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
    LineaError = "000-0006"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAd_Click()
    If MostrarTouch Then
        LineaError = "000-0007"
        Form_KeyDown TeclaPagAd, 0
        LineaError = "000-0008"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
    LineaError = "000-0009"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAt_Click()
    If MostrarTouch Then
        LineaError = "000-0010"
        Form_KeyDown TeclaPagAt, 0
        LineaError = "000-0011"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
    LineaError = "000-0012"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Command1_Click()
    If MostrarTouch Then
        LineaError = "000-0023"
        Form_KeyDown TeclaOK, 0
    End If
End Sub

Private Sub Form_Activate()
    LineaError = "000-0024"
    MostrarCursor False
    'actualizar los precios
    LineaError = "000-0025"
    If TemasPorCredito = 1 Then
        LineaError = "000-0026"
        lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " credito"
        lblPrecios2 = "1 coin = " + CStr(TemasPorCredito) + " credito"
    Else
        LineaError = "000-0027"
        lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " creditos"
        lblPrecios2 = "1 coin = " + CStr(TemasPorCredito) + " creditos"
    End If
    LineaError = "000-0028"
    If CreditosCuestaTema = 1 Then
        LineaError = "000-0029"
        lblPrecios = lblPrecios + vbCrLf + "1 credito = 1 tema"
        lblPrecios2 = lblPrecios2 + " / " + "1 credito = 1 tema"
    Else
        LineaError = "000-0030"
        lblPrecios = lblPrecios + vbCrLf + CStr(CreditosCuestaTema) + " creditos = 1 tema"
        lblPrecios2 = lblPrecios2 + " / " + CStr(CreditosCuestaTema) + " creditos = 1 tema"
    End If
    'agreagr el precio de los videos!!!
    If CreditosCuestaTemaVIDEO = 1 Then
        LineaError = "000-0029"
        lblPrecios = lblPrecios + vbCrLf + "1 credito = 1 VIDEO"
        lblPrecios2 = lblPrecios2 + " / " + "1 credito = 1 VIDEO"
    Else
        LineaError = "000-0030"
        lblPrecios = lblPrecios + vbCrLf + CStr(CreditosCuestaTemaVIDEO) + " creditos = 1 VIDEO"
        lblPrecios2 = lblPrecios2 + " / " + CStr(CreditosCuestaTemaVIDEO) + " creditos = 1 VIDEO"
    End If
    
    'total ser�a
    '1 coin = 8 creditos /// " + "8 creditos = 1 tema /// 8 creditos = 1 VIDEO
        
    LineaError = "000-0031"
    If HabilitarVUMetro Then
        If Is3pmExclusivo Then
            If VU21.inHabilitado = False And VU21.IsPlaying = False Then
                VU21.DoStart
            End If
        Else
            If VU1.inHabilitado = False And VU1.IsPlaying = False Then
                VU1.DoStart
            End If
        End If
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    LineaError = "000-0033"
    Dim PagNum As Long
        
    Select Case KeyCode
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
                End
            End If
        Case vbKeyU
            frmOnlyContador.Show 1
        Case vbKeyJ 'avanzar 10 segundos
            Dim ToSec As Long
            ToSec = MP3.PositionInSec * 1000 + 10000
            MP3.SeekTo CStr(ToSec)
        'subir o bajar volumen
        Case 68 'es la D y baja el volumen
            LineaError = "000-0034"
            If frmIndex.MP3.IsPlaying Then
                LineaError = "000-0035"
                If VolumenIni <= 5 Then
                    LineaError = "000-0036"
                    frmIndex.MP3.Volumen = 0
                Else
                    LineaError = "000-0037"
                    frmIndex.MP3.Volumen = VolumenIni - 5
                End If
                LineaError = "000-0038"
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case 69 'es la E y sube el volumen
            LineaError = "000-0039"
            If frmIndex.MP3.IsPlaying Then
                LineaError = "000-0039"
                If VolumenIni >= 95 Then
                    LineaError = "000-0040"
                    frmIndex.MP3.Volumen = 100
                Else
                    LineaError = "000-0041"
                    frmIndex.MP3.Volumen = VolumenIni + 5
                End If
                LineaError = "000-0042"
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case 66 ' es la b y pasa al siguiente tema
            'si es video ocultar la pantalla de video
            LineaError = "000-0043"
            'If EsVideo Then
            '    picVideo.Visible = False
            'End If
            LineaError = "000-0044"
            EMPEZAR_SIGUIENTE
        Case TeclaPagAd
            'pase lo que pase registrar
            LineaError = "000-0054"
            TECLAS_PRES = TECLAS_PRES + "5"
            LineaError = "000-0055"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0056"
            lblTECLAS = TECLAS_PRES
            
            'es para abajo en el modo 5 y pagina adelante den el modo 46
            If IsMod46Teclas = 46 Then
                'esta tecla es pagina adelante en el modo 46 y abajo en el modo 5
                LineaError = "000-0045"
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                LineaError = "000-0046"
                Dim PrimeroDeLaPaginaQueSigue As Long
                LineaError = "000-0047"
                PrimeroDeLaPaginaQueSigue = (PagNum + 1) * (TapasMostradasH * TapasMostradasV)
                LineaError = "000-0048"
                If PrimeroDeLaPaginaQueSigue < TOTAL_DISCOS Then
                    LineaError = "000-0049"
                    If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                    LineaError = "000-0050"
                    DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueSigue, True, 1)
                    LineaError = "000-0051"
                    lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueSigue + 1) + " de " + CStr(TOTAL_DISCOS)
                    LineaError = "000-0052"
                    nDiscoSEL = 0
                End If
            End If
            'si esta eligiendo discos en modo video min es
            'totalmente desitinto, solo va al que sigue
            'no importann p�ginas ni nada
            'If EstoyEnModoVideoMiniSelDisco = False Then
            '    'xxxx
            '    Exit Sub
            'End If
            If IsMod46Teclas = 5 Then
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo As Long
                DiskToGo = nDiscoSEL + TapasMostradasH
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo < DiscosEnPagina Then
                    nDiscoGral = nDiscoGral + TapasMostradasH
                    LineaError = "000-0083"
                    UnSelDisco nDiscoSEL
                    LineaError = "000-0084"
                    SelDisco nDiscoSEL + TapasMostradasH
                End If
            End If
            
        Case TeclaPagAt
            If IsMod46Teclas = 46 Then
                'esta tecla es pagina atras en el modo 46 y arriba en el modo 5
                LineaError = "000-0056"
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                LineaError = "000-0057"
                If PagNum > 0 Then
                    LineaError = "000-0058"
                    Dim PrimeroDeLaPaginaQueAnterior As Long
                    PrimeroDeLaPaginaQueAnterior = (PagNum - 1) * (TapasMostradasH * TapasMostradasV)
                    LineaError = "000-0059"
                    If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                    LineaError = "000-0060"
                    DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueAnterior, False, 1)
                    LineaError = "000-0061"
                    lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueAnterior + 1) + " de " + CStr(TOTAL_DISCOS)
                    
                End If
            End If
            If IsMod46Teclas = 5 Then
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo2 As Long
                DiskToGo2 = nDiscoSEL - TapasMostradasH
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo2 >= 0 Then
                    nDiscoGral = nDiscoGral - TapasMostradasH
                    LineaError = "000-0083"
                    UnSelDisco nDiscoSEL
                    LineaError = "000-0084"
                    SelDisco nDiscoSEL - TapasMostradasH
                End If
            End If
            LineaError = "000-0062"
            TECLAS_PRES = TECLAS_PRES + "6"
            LineaError = "000-0063"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0064"
            lblTECLAS = TECLAS_PRES
        Case TeclaConfig
            LineaError = "000-0065"
             frmConfig.Show 1
        Case TeclaIZQ
            'esta tecla es IZQ en el modo 46 pasandpo de arriba aa abjo y siguiendo a la pag ant en el modo 5
            'ver si desplazo temas en modo video
            LineaError = "000-0066"
            If ModoVideoSelTema Then
                LineaError = "000-0067"
                If TemaElegidoModoVideo > 0 Then
                    LineaError = "000-0068"
                    UnSelTema TemaElegidoModoVideo
                    LineaError = "000-0069"
                    TemaElegidoModoVideo = TemaElegidoModoVideo - 1
                    LineaError = "000-0070"
                    SelTema TemaElegidoModoVideo
                    LineaError = "000-0071"
                    OrdenarListaTemaVideo
                End If
            Else
                'no ir a -1
                LineaError = "000-0072"
                'ver si es el primero
                If nDiscoSEL = 0 Then
                    'ver si hay que pasar hoja o no
                    LineaError = "000-0073"
                    If PasarHoja Then
                        LineaError = "000-0074"
                        'ver si hay p�ginas antes
                        'si el gral es mayor que cero entonces si hay
                        'en la primera p�gina gral y discosel son iguales
                        If nDiscoGral > 0 Then
                            'como si viene eligiendo desde la ultima fila
                            If IsMod46Teclas = 46 Then DiscosEnPagina = _
                                CargarDiscos(nDiscoGral - _
                                ((TapasMostradasH * TapasMostradasV)), False, TapasMostradasV)
                            
                            'busca solo la fila!!
                            If IsMod46Teclas = 5 Then DiscosEnPagina = _
                                CargarDiscos(nDiscoGral - _
                                ((TapasMostradasH * TapasMostradasV)), False, EnQueFilaEstoy)
                        End If
                                
                    Else
                        'NO NO NO!!!! nDiscoGral = (TapasMostradasH * TapasMostradasV) - 1
                        'estoy en una hoja al principio y debo elegir el disco del final
                        'sel y unsel trabajan con referencias de o al total de discos por pag
                        'nDiscoGral es el numero absoluto del disco
                        'ver si existe el disco al que voy
                        LineaError = "000-0075"
                        If TOTAL_DISCOS > nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1 Then
                            LineaError = "000-0076"
                            nDiscoGral = nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1
                            LineaError = "000-0077"
                            UnSelDisco nDiscoSEL
                            LineaError = "000-0078"
                            SelDisco (TapasMostradasH * TapasMostradasV) - 1
                        Else
                            LineaError = "000-0079"
                            nDiscoGral = TOTAL_DISCOS - 1
                            LineaError = "000-0080"
                            UnSelDisco nDiscoSEL
                            LineaError = "000-0081"
                            SelDisco DiscosEnPagina - 1
                        End If
                    End If
                Else
                    'si no es el primero ver si es
                    'el primero de una fila y esta en modo 5 el teclado
                    If nDiscoSEL = TapasMostradasH * (EnQueFilaEstoy - 1) Then
                        'si esta en el modo 5 me fijo si esta al final de una l�nea
                        If IsMod46Teclas = 5 Then
                            'el disco a iniciar ya no es nDiscoGral-(tapash*tapasv)!!!!!!
                            'hay que restar tambien el nOrden de esta pagina
                            Dim DiscoToIni As Long
                            'el primero de esta mas el total de esta!
                            DiscoToIni = nDiscoGral - nDiscoSEL - (TapasMostradasH * TapasMostradasV)
                            'ver que no se vaya a la mierda!!
                            If DiscoToIni >= 0 Then
                                DiscosEnPagina = CargarDiscos(DiscoToIni, False, EnQueFilaEstoy)
                            End If
                        Else
                            'tratarlo normalmente como el 46
                            GoTo Mod46IZQ
                        End If
                    Else
Mod46IZQ:
                        LineaError = "000-0082"
                        nDiscoGral = nDiscoGral - 1
                        LineaError = "000-0083"
                        UnSelDisco nDiscoSEL
                        LineaError = "000-0084"
                        SelDisco nDiscoSEL - 1
                    End If
                End If
                LineaError = "000-0085"
                lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            End If
            LineaError = "000-0086"
            TECLAS_PRES = TECLAS_PRES + "1"
            LineaError = "000-0087"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0088"
            lblTECLAS = TECLAS_PRES
        Case TeclaDER
            'esta tecla es IZQ en el modo 46 pasandpo de arriba a abjo
            'y siguiendo a la pag sig en el modo 5
            LineaError = "000-0089"
            If ModoVideoSelTema Then
                LineaError = "000-0090"
                If TemaElegidoModoVideo < UBound(MATRIZ_TEMAS) Then
                    LineaError = "000-0091"
                    UnSelTema TemaElegidoModoVideo
                    LineaError = "000-0092"
                    TemaElegidoModoVideo = TemaElegidoModoVideo + 1
                    LineaError = "000-0093"
                    SelTema TemaElegidoModoVideo
                    LineaError = "000-0094"
                    OrdenarListaTemaVideo
                End If
            Else
                'esta eligiendo discos ya sea en las portadas o en el modo video!!
                LineaError = "000-0095"
                If nDiscoSEL = DiscosEnPagina - 1 Then
                    'ver si hay que pasar hojas (segun config)
                    If PasarHoja Then
                        LineaError = "000-0096"
                        'ver que no se vaya a la mierda!!
                        If nDiscoGral + 1 < TOTAL_DISCOS Then
                            LineaError = "000-0097"
                            'si esta en el modtec 46 pasa al primero
                            'pero si esta en el modo 5 pasa a su mismo nivel
                            'vertical en la hoja que sigue
                            If IsMod46Teclas = 46 Then DiscosEnPagina = _
                                CargarDiscos(nDiscoGral + 1, True, 1) 'va a la primera fila!!
                            'busca solo la fila!!
                            If IsMod46Teclas = 5 Then DiscosEnPagina = _
                                CargarDiscos(nDiscoGral + 1, True, EnQueFilaEstoy)
                        End If
                    Else
                        '------------------------------
                        'si no esta configurado para pasar hojas entonces debe _
                        estar en el modo 46
                        'en el modo 5 no hay salto de p�gina...
                        '------------------------------
                        '!!!NO NO NO nDiscoGral = 0
                        'estoy en una hoja al final y debo elegir el disco del principio
                        'sel y unsel trabajan con referencias de o al total de discos por pag
                        'nDiscoGral es el numero absoluto del disco
                        LineaError = "000-0098"
                        nDiscoGral = nDiscoGral - DiscosEnPagina + 1
                        LineaError = "000-0099"
                        UnSelDisco nDiscoSEL
                        LineaError = "000-0100"
                        SelDisco 0
                    End If
                Else
                    'ver si llego al final de una linea horizontal para pasar a la hoja
                    'que sigue si esta en el modTeclado5
                    
                    LineaError = "000-0101"
                    'ver si el disco existe !!! o llegamos al final de todo !!!!
                    If nDiscoGral + 1 < TOTAL_DISCOS Then
                        'si esta en el modo 5 me fijo si esta al final de una l�nea
                        If IsMod46Teclas = 5 Then
                            'ver ahora si es el �ltimo de una l�nea!!!
                            If nDiscoSEL = (TapasMostradasH * EnQueFilaEstoy) - 1 Then
                                'el disco a iniciar ya no es nDiscoGral + 1  !!!!!!
                                Dim DiscoToIni2 As Long
                                'el primero de esta mas el total de esta!
                                DiscoToIni2 = nDiscoGral - nDiscoSEL + (TapasMostradasH * TapasMostradasV)
                                'ver que no se vaya a la mierda!!
                                If DiscoToIni2 < TOTAL_DISCOS Then
                                    DiscosEnPagina = CargarDiscos(DiscoToIni2, True, EnQueFilaEstoy)
                                End If
                                
                            Else
                                'tratarlo como el modo 46
                                GoTo Mod46
                            End If
                        Else
Mod46:
                            LineaError = "000-0102"
                            nDiscoGral = nDiscoGral + 1
                            LineaError = "000-0103"
                            UnSelDisco nDiscoSEL
                            LineaError = "000-0104"
                            SelDisco nDiscoSEL + 1
                        End If
                    End If
                End If
            End If
            LineaError = "000-0105"
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            LineaError = "000-0106"
            TECLAS_PRES = TECLAS_PRES + "2"
            LineaError = "000-0107"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0108"
            lblTECLAS = TECLAS_PRES
        Case TeclaOK
            LineaError = "000-0109"
            TECLAS_PRES = TECLAS_PRES + "3"
            LineaError = "000-0110"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0111"
            lblTECLAS = TECLAS_PRES
            'si estoy en video
            'saber si estoy eligiendo tema. Si no estoy en disco
            
            If ModoVideoSelTema Then
                'si esta en fullscreen NO EJECUTAR!!!
                'solo si no sale por la segunda salida!!!
                If EsVideo And vidFullScreen And Salida2 = False Then GoTo FinKD 'fin keydown
                'si no dice salir cargar tema
                LineaError = "000-0112"
                If T(TemaElegidoModoVideo) = "SALIR" Or T(TemaElegidoModoVideo) = "No hay temas" Then
                    'volver a elegir discos
                    LineaError = "000-0113"
                    frTEMAS.Visible = False
                    LineaError = "000-0114"
                    lblTEMAS.Visible = False
                    LineaError = "000-0115"
                    frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                    LineaError = "000-0116"
                    UnSelTema 0
                    LineaError = "000-0117"
                    ModoVideoSelTema = False
                Else
                    'ejecutar el tema
                    LineaError = "000-0118"
                    
                    'ANTES DE VER CUANTOS CREDITOS NECESITA TENGO QUE SABER SI QUIERE EJECUTAR
                    'MP3 O VIDEO!!!!!!
                    LineaError = "000-0126"
                    Dim temaElegido As String
                    'lstext es una lista oculta  con datos completos
                    temaElegido = txtInLista(MATRIZ_TEMAS(TemaElegidoModoVideo), 0, ",")
                    
                    If LCase(Right(temaElegido, 3)) = "mp3" Then
                        PideVideo = False
                    Else
                        PideVideo = True
                    End If
                    'ver si puede pagar lo que pide!!!
                    'que joyita papa!!!. Parece que supieras programar
                    '--------------------------------------------------------------
                    If (PideVideo = False And CREDITOS >= CreditosCuestaTema) Or _
                        (PideVideo And CREDITOS >= CreditosCuestaTemaVIDEO) Then
                    '--------------------------------------------------------------
                        LineaError = "000-0119"
                        'restar lo que corresponde!!!
                        If PideVideo Then
                            CREDITOS = CREDITOS - CreditosCuestaTemaVIDEO
                        Else
                            CREDITOS = CREDITOS - CreditosCuestaTema
                        End If
                        'siempre que se ejecute un credito estaremos por debajo de maximo
                        LineaError = "000-0120"
                        OnOffCAPS vbKeyScrollLock, True
                        'grabar cant de creditos
                        LineaError = "000-0121"
                        EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                        LineaError = "000-0122"
                        
                        ShowCredits
                        
                        LineaError = "000-0124"
                        'grabar credito para validar
                        'creditosValidar ya se cargo en load de frmindex
                        CreditosValidar = CreditosValidar + TemasPorCredito
                        LineaError = "000-0125"
                        EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
                        
                        LineaError = "000-0127"
                        'si esta ejecutando pasa a la lista de reproducci�n
                        If MP3.IsPlaying Then
                            'pasar a la lista de reproducci�n
                            Dim NewIndLista As Long
                            LineaError = "000-0128"
                            NewIndLista = UBound(MATRIZ_LISTA)
                            LineaError = "000-0129"
                            ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                            LineaError = "000-0130"
                            'se graba en Matriz_Listas como path, nombre(sin .mp3)
                            MATRIZ_LISTA(NewIndLista + 1) = _
                                temaElegido + "," + _
                                FSO.GetBaseName(T(TemaElegidoModoVideo)) + _
                                " / " + FSO.GetBaseName(UbicDiscoActual)
                            LineaError = "000-0131"
                            CargarProximosTemas
                            'graba en reini.tbr los datos que correspondan por si se corta la luz
                            LineaError = "000-0132"
                            CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                            'volver a elegir discos
                            LineaError = "000-0133"
                            frTEMAS.Visible = False
                            LineaError = "000-0134"
                            lblTEMAS.Visible = False
                            LineaError = "000-0135"
                            frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                            LineaError = "000-0136"
                            UnSelTema 0
                            LineaError = "000-0137"
                            ModoVideoSelTema = False
                        Else
                            'NUNCA ENTRARA AQUI, siempre esta rep video
                            'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                            'paciencia
                            LineaError = "000-0138"
                            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                            LineaError = "000-0139"
                            EjecutarTema temaElegido, True
                        End If
                        
                        VerSiTocaPUB
                        
                    End If
                End If
            Else
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                LineaError = "000-0140"
                'ver si es video deber�a desplegar los temas del disco elegido
                'en modo de texto
                'pero si estoy viendo el video en salida2 es video sera verdadero
                'pero de todas formas no veo als lista de texto y sigo igual
                'solo si esvideo y necesito el modo texto del video!!!!
                If EsVideo And Salida2 = False Then
                    frModoVideo.Height = frDISCOS.Height / 4
                    LineaError = "000-0141"
                    OrdenarListaModoVideo
                    LineaError = "000-0142"
                    lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
                    LineaError = "000-0143"
                    lblTEMAS.Left = lblModoVideo.Left
                    LineaError = "000-0144"
                    frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
                    LineaError = "000-0145"
                    frTEMAS.Height = frDISCOS.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
                    LineaError = "000-0146"
                    lblTEMAS.Visible = True
                    LineaError = "000-0147"
                    frTEMAS.Visible = True
                    LineaError = "000-0148"
                    'cargar los temas multimedia en t()
                    ReDim MATRIZ_TEMAS(0) 'matriz en blanco
                    'es una matriz global
                    LineaError = "000-0149"
                    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral + 1), 0, ",")
                    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
                    LineaError = "000-0150"
                    ReDim Preserve MATRIZ_TEMAS(0)
                    LineaError = "000-0151"
                    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
                    LineaError = "000-0152"
                    If UBound(MATRIZ_TEMAS) = 0 Then
                        LineaError = "000-0153"
                        T(0) = "No hay temas"
                        LineaError = "000-0154"
                        SelTema 0
                        LineaError = "000-0155"
                        ModoVideoSelTema = True
                        LineaError = "000-0156"
                        Exit Sub
                    End If
                    LineaError = "000-0157"
                    T(0) = "SALIR"
                    '----------------------------
                    'a daniel cruz le da un error como si se volviera a cargar algo que esta cargado
                    'por lo tanto tengo que poner un manejador de error aqui, unico lugar en que se carga esto
                    LineaError = "000-0158"
                    For Each LLL In frmIndex.T
                        LineaError = "000-0159"
                        If LLL.Index > 0 Then Unload LLL
                    Next
                    '----------------------------
                    LineaError = "000-0160"
                    For AA = 1 To UBound(MATRIZ_TEMAS)
                        LineaError = "000-0161"
                        Load T(AA)
                        LineaError = "000-0162"
                        T(AA) = FSO.GetBaseName(txtInLista(MATRIZ_TEMAS(AA), 1, ","))
                        LineaError = "000-0163"
                        T(AA).Top = T(AA - 1).Top + T(AA - 1).Height
                        LineaError = "000-0164"
                        T(AA).Left = T(AA - 1).Left
                        LineaError = "000-0165"
                        T(AA).Visible = True
                        LineaError = "000-0166"
                    Next
                    TemaElegidoModoVideo = 0
                    LineaError = "000-0167"
                    SelTema 0
                    LineaError = "000-0168"
                    ModoVideoSelTema = True
                Else
                    LineaError = "000-0169"
                    If lblDisco(nDiscoSEL) = "01- Los mas escuchados" Then GoTo TOP10Show
                    LineaError = "000-0170"
                    frmTemasDeDisco.Show 1
                End If
            End If
        Case TeclaCerrarSistema
            LineaError = "000-0171"
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            LineaError = "000-0172"
            MostrarCursor True
            LineaError = "000-0173"
            MP3.DoClose
            LineaError = "000-0174"
            If ApagarAlCierre Then APAGAR_PC
            LineaError = "000-0175"
            Unload Me
            End
        Case TeclaESC
            LineaError = "000-0176"
            TECLAS_PRES = TECLAS_PRES + "4"
            LineaError = "000-0177"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            LineaError = "000-0178"
            lblTECLAS = TECLAS_PRES
            LineaError = "000-0179"
            If ModoVideoSelTema Then
                LineaError = "000-0180"
                'volver a elegir discos
                frTEMAS.Visible = False
                LineaError = "000-0181"
                lblTEMAS.Visible = False
                LineaError = "000-0182"
                frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                LineaError = "000-0183"
                UnSelTema 0
                LineaError = "000-0184"
                ModoVideoSelTema = False
            End If
    End Select
FinKD:
    LineaError = "000-0185"
    VerClaves TECLAS_PRES
    LineaError = "000-0186"
    SecSinTecla = 0
    LineaError = "000-0187"
    lblNoTecla = 0
    Exit Sub
TOP10Show:
    LineaError = "000-0188"
    FRMTOP10.Show 1
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    LineaError = "000-0189"
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        LineaError = "000-0190"
        If CREDITOS <= MaximoFichas Then
            LineaError = "000-0191"
            OnOffCAPS vbKeyScrollLock, True
            LineaError = "000-0192"
            CREDITOS = CREDITOS + TemasPorCredito
            LineaError = "000-0193"
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            LineaError = "000-0194"
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            LineaError = "000-0195"
            
            ShowCredits
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            LineaError = "000-0198"
            CreditosValidar = CreditosValidar + TemasPorCredito
            LineaError = "000-0199"
            EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
        Else
            'apagar el fichero electronico
            LineaError = "000-0200"
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    
    'imagenes no cargadas
    picFondoDisco.Picture = LoadPicture(SYSfolder + "f3ya.nam")
    picFondo.Picture = LoadPicture(SYSfolder + "f4ya.nam")
    tbrPassImg1.Picture SYSfolder + "f8ya.nam"
    TapaEjecutando.Picture = LoadPicture(SYSfolder + "f8ya.nam")
    'la imagen chiquita del exclusivo es la misma!!
    Image1.Picture = LoadPicture(SYSfolder + "f8ya.nam")
    VU21.Picture SYSfolder + "f4ya.nam"
    
    If Is3pmExclusivo Then
        'poner el picfondo.top a la altura del VU21 ya que todo esta basado en ese top!!!
        VU21.Top = Me.Height - VU21.Height
        picFondo.Top = VU21.Top
        VU21.Visible = True
        picFondo.Visible = False
    Else
        VU21.Visible = False
        picFondo.Visible = True
    End If
    On Error GoTo NoLoadIndex
    
    Prog.MIN = 0 'barra de progreso circular
    picFondoDisco.Top = 0
    picFondoDisco.Left = 0
    
    LineaError = "000-0201"
    RegistroDiario 'anota la fecha, hora y numero del contador
    '--------
    LineaError = "000-0202"
    If K.LICENCIA = HSuperLicencia Then
        LineaError = "000-0203"
        If FSO.FileExists(WINfolder + "SL\indexchi.tbr") Then
            tbrPassImg1.Picture WINfolder + "SL\indexchi.tbr"
            'la imagen chiquita del exclusivo es la misma!!
            Image1.Picture = LoadPicture(WINfolder + "SL\indexchi.tbr")
        End If
    End If
    '--------
    LineaError = "000-0204"
    AjustarFRM Me, 12000
    LineaError = "000-0205"
    If K.LICENCIA = aSinCargar Then
        LineaError = "000-0206"
        lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        LineaError = "000-0207"
        lblDEMO = textoUsuario
    End If
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    
    LineaError = "000-0208"
    '-----------------
    If K.LICENCIA = HSuperLicencia Then
        LineaError = "000-0209"
        If FSO.FileExists(WINfolder + "SL\txtIDX.tbr") Then
            LineaError = "000-0210"
            Set TE = FSO.OpenTextFile(WINfolder + "SL\txtIDX.tbr", ForReading, False)
            LineaError = "000-0211"
            Dim NewT As String
            LineaError = "000-0212"
            NewT = TE.ReadAll
            LineaError = "000-0213"
            lblTBR = NewT
            LineaError = "000-0214"
            TE.Close
        Else
            LineaError = "000-0215"
            lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
        End If
    Else
        LineaError = "000-0216"
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
    End If
    '-----------------
    LineaError = "000-0217"
    VU1.Width = Screen.Width
    LineaError = "000-0218"
    VU1.Left = 0: VU1.Top = 0
    LineaError = "000-0219"
    VU1.Height = picFondo.Top - 25
    LineaError = "000-0220"
    
    'si es exclusivo inhabilito el vumetro GRANDE !!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        'que entre en el control
        LineaError = "000-0221"
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width
        LineaError = "000-0222"
        frDISCOS.Left = VU1.AnchoBarra + 25 '0
    Else
        LineaError = "000-0223"
        frDISCOS.Left = 0 ' tapa a las barras que no se usan 'VU1.Left + VU1.Width
        LineaError = "000-0224"
        frDISCOS.Width = VU1.Width ' Screen.Width - VU1.Width
    End If
    LineaError = "000-0225"
    frDISCOS.Top = 0
    LineaError = "000-0226"
    frDISCOS.Height = picFondo.Top
    LineaError = "000-0227"
    picFondoDisco.Height = frDISCOS.Height
    LineaError = "000-0228"
    picFondoDisco.Width = frDISCOS.Width
    
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    LineaError = "000-0231"
    If MostrarTouch = False Then
        LineaError = "000-0232"
        Frame1.Visible = False 'frame del touch
        lblTemaSonando.Width = Screen.Width - lblTemaSonando.Left - 250
        lstProximos.Width = Screen.Width - lstProximos.Left - 250
        lblTBR.Width = Screen.Width - lblTBR.Left - 250
        lblDEMO.Width = Screen.Width - lblDEMO.Left - 250
    End If
    'frDISCOS contiene los discos a mostrar
    'se debera calcualr el tama�o de cada discos asi como cantidad horizontal y vertical
    LineaError = "000-0241"
    Dim AnchoTapaDisco As Long
    Dim AltoTapaDisco As Long
    'el alto de estos incluye tambien el lbldisco
    LineaError = "000-0242"
    AnchoTapaDisco = (frDISCOS.Width * 0.98 / TapasMostradasH)
    LineaError = "000-0243"
    AltoTapaDisco = (frDISCOS.Height * 0.97 / TapasMostradasV)
    'ver cual es mayor para no permitir mucha distorsion
    'lo que se ajuste se agranda del espacio entrediscos
    LineaError = "000-0244"
    Dim EspacioEntreDiscosH As Long
    Dim EspacioEntreDiscosV As Long
    LineaError = "000-0245"
    EspacioEntreDiscosV = 50: EspacioEntreDiscosH = 50
    LineaError = "000-0246"
    If DistorcionarTapas = False Then
        LineaError = "000-0247"
        Dim DIFF As Double
        LineaError = "000-0248"
        DIFF = AnchoTapaDisco - AltoTapaDisco
        LineaError = "000-0249"
        If DIFF > 0 Then
            LineaError = "000-0250"
            'el ancho es mas que el alto
            AnchoTapaDisco = AltoTapaDisco
            LineaError = "000-0251"
            EspacioEntreDiscosH = DIFF
        Else
            LineaError = "000-0252"
            'el alto es mas que el ancho
            AltoTapaDisco = AnchoTapaDisco
            LineaError = "000-0253"
            EspacioEntreDiscosV = -DIFF
        End If
    End If
    LineaError = "000-0254"
    If MostrarRotulos Then
        LineaError = "000-0255"
        TapaCD(0).Width = AnchoTapaDisco
        LineaError = "000-0256"
        TapaCD(0).Height = AltoTapaDisco * 0.79 '80%disco, 20% lbldisco
        LineaError = "000-0257"
        lblDisco(0).Height = AltoTapaDisco * 0.19 '80%disco, 20% lbldisco
        LineaError = "000-0258"
        lblDisco(0).Width = AnchoTapaDisco
    Else
        LineaError = "000-0259"
        TapaCD(0).Width = AnchoTapaDisco
        LineaError = "000-0260"
        TapaCD(0).Height = AltoTapaDisco
        LineaError = "000-0261"
        lblDisco(0).Visible = False
    End If
    'centrar!!
    Dim IniCentrarH As Long
    IniCentrarH = EspacioEntreDiscosH
    Dim IniCentrarV As Long
    IniCentrarV = EspacioEntreDiscosV
    LineaError = "000-0262"
    lblDisco(0).Left = IniCentrarH
    LineaError = "000-0268"
    TapaCD(0).Left = IniCentrarH
    'ver si los rotulos van arriba o abajo
    If RotulosArriba Then
        LineaError = "000-0263"
        lblDisco(0).Top = IniCentrarV
        LineaError = "000-0265"
        If MostrarRotulos Then
            LineaError = "000-0266"
            TapaCD(0).Top = lblDisco(0).Top + lblDisco(0).Height + 50
        Else
            LineaError = "000-0267"
            TapaCD(0).Top = IniCentrarV
        End If
    Else
        LineaError = "000-0269"
        TapaCD(0).Top = IniCentrarV
        LineaError = "000-0271"
        lblDisco(0).Top = TapaCD(0).Top + TapaCD(0).Height + 50
    End If
    LineaError = "000-0272"
    Dim CantDiscos As Long
    LineaError = "000-0273"
    CantDiscos = TapasMostradasH * TapasMostradasV
    'cargar la cantidad de tapas correspondientes
    LineaError = "000-0274"
    C = 0
    LineaError = "000-0275"
    Do While C < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        LineaError = "000-0276"
        C = C + 1
        LineaError = "000-0277"
        Load TapaCD(C)
        LineaError = "000-0278"
        Load lblDisco(C)
        'ya toman el tama�o del original
        LineaError = "000-0279"
        If C / TapasMostradasH = C \ TapasMostradasH Then
            'es una tapa al principio de linea
            LineaError = "000-0280"
            If RotulosArriba Then
                LineaError = "000-0281"
                lblDisco(C).Left = IniCentrarH
                LineaError = "000-0282"
                lblDisco(C).Top = TapaCD(C - TapasMostradasH).Top + TapaCD(C - TapasMostradasH).Height + EspacioEntreDiscosV
                LineaError = "000-0283"
                TapaCD(C).Left = IniCentrarH
                If MostrarRotulos Then
                    LineaError = "000-0284"
                    TapaCD(C).Top = lblDisco(C).Top + lblDisco(C).Height + 50
                Else
                    LineaError = "000-0285"
                    TapaCD(C).Top = TapaCD(C - TapasMostradasH).Top + TapaCD(C - TapasMostradasH).Height + 50
                End If
                LineaError = "000-0286"
                TapaCD(C).Visible = True
                LineaError = "000-0287"
                If MostrarRotulos Then lblDisco(C).Visible = True
            Else
                LineaError = "000-0288"
                TapaCD(C).Left = IniCentrarH
                If MostrarRotulos Then
                    LineaError = "000-0289"
                    TapaCD(C).Top = lblDisco(C - TapasMostradasH).Top + lblDisco(C - TapasMostradasH).Height + EspacioEntreDiscosV
                Else
                    LineaError = "000-0290"
                    TapaCD(C).Top = TapaCD(C - TapasMostradasH).Top + TapaCD(C - TapasMostradasH).Height + EspacioEntreDiscosV
                End If
                LineaError = "000-0291"
                lblDisco(C).Left = IniCentrarH
                LineaError = "000-0292"
                lblDisco(C).Top = TapaCD(C).Top + TapaCD(C).Height + 50
                LineaError = "000-0293"
                TapaCD(C).Visible = True
                LineaError = "000-0294"
                If MostrarRotulos Then lblDisco(C).Visible = True
            End If
        Else
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                LineaError = "000-0295"
                lblDisco(C).Left = lblDisco(C - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                LineaError = "000-0296"
                lblDisco(C).Top = lblDisco(C - 1).Top
                LineaError = "000-0297"
                TapaCD(C).Left = lblDisco(C).Left
                LineaError = "000-0298"
                TapaCD(C).Top = TapaCD(C - 1).Top
                LineaError = "000-0299"
                TapaCD(C).Visible = True
            Else
                LineaError = "000-0300"
                TapaCD(C).Left = TapaCD(C - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                LineaError = "000-0301"
                TapaCD(C).Top = TapaCD(C - 1).Top
                LineaError = "000-0302"
                lblDisco(C).Left = TapaCD(C).Left
                LineaError = "000-0303"
                lblDisco(C).Top = lblDisco(C - 1).Top
                LineaError = "000-0304"
                TapaCD(C).Visible = True
            End If
            LineaError = "000-0305"
            If MostrarRotulos Then lblDisco(C).Visible = True
        End If
        
    Loop
    LineaError = "000-0306"
    OnOffCAPS vbKeyScrollLock, True
    LineaError = "000-0307"
    lblV = "versi�n " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    LineaError = "000-0308"
    lblTiempoRestante = "FALTA: " + "00:00"
    'ocultar las etiquetas
    LineaError = "000-0309"
    Me.AutoRedraw = AutoReDibuj
    LineaError = "000-0310"
    Me.Left = Screen.Width / 2 - Me.Width / 2
    LineaError = "000-0311"
    Me.Top = Screen.Height / 2 - Me.Height / 2
    'ver cuantos creditos hay
    LineaError = "000-0312"
    CREDITOS = Val(LeerArch1Linea(AP + "creditos.tbr"))
    LineaError = "000-0313"
    
    ShowCredits
    
    'dejar cargado el mostrados de procesos
    'Load frmini
    'cargar las variables globales
    LineaError = "000-0316"
    TEMA_REPRODUCIENDO = "Sin reproducci�n actual"
    LineaError = "000-0317"
    TEMA_SIGUIENTE = "No hay proximo tema"
    LineaError = "000-0318"
    TEMAS_EN_LISTA = 0
    
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
    LineaError = "000-0319"
    MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
    LineaError = "000-0320"
    Dim CarpActual As String
    Dim pathTema As String, DuracionTema As String, nombreTEMA As String
    'mostrar proceso
    LineaError = "000-0321"
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    
    'ret devuelve la cantidadd de discos cargados
    LineaError = "000-0322"
    DiscosEnPagina = CargarDiscos(0, True, 1)
    'inicializar la matriz_lista (lista de reproduccion
    LineaError = "000-0323"
    ReDim MATRIZ_LISTA(0)
    LineaError = "000-0324"
    lblTOTdiscos = "Discos: " + Trim(Str(UBound(MATRIZ_DISCOS)))
    
    'si quedaron temas pendientes cargarlos
    LineaError = "000-0325"
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            LineaError = "000-0326"
            If FSO.FileExists(AP + "reini.tbr") Then
                LineaError = "000-0327"
                Set TE = FSO.OpenTextFile(AP + "reini.tbr", ForReading, False)
                Dim TT As String 'cada tema
                Dim z As Integer 'contador de temas en lista anterior
                z = 1
                LineaError = "000-0328"
                Do While Not TE.AtEndOfStream
                    LineaError = "000-0329"
                    TT = TE.ReadLine
                    LineaError = "000-0330"
                    ReDim Preserve MATRIZ_LISTA(z)
                    LineaError = "000-0331"
                    MATRIZ_LISTA(z) = TT
                    LineaError = "000-0332"
                    z = z + 1
                Loop
                LineaError = "000-0333"
                TE.Close
            End If
            LineaError = "000-0334"
            EMPEZAR_SIGUIENTE
        Case "NADA"
            'no hacer nada
            'borrar la lista
            LineaError = "000-0335"
            'borrra los temas 'y los creditos?
            If FSO.FileExists(AP + "reini.tbr") Then FSO.DeleteFile AP + "reini.tbr", True
            
            LineaError = "000-0336"
            Timer1.Interval = 10000
    End Select
    LineaError = "000-0337"
    Unload frmINI
    
    'ver si hay validacion por creditos
    LineaError = "000-0338"
    Validar = LeerConfig("Validar", "0")
    If Validar Then
        'ver si existe el archivo Creditos Validar
        LineaError = "000-0339"
        If FSO.FileExists(SYSfolder + "\radilav.cfg") Then
            'leer el archivo de creditos vaildados
            LineaError = "000-0340"
            CreditosValidar = CLng(LeerArch1Linea(SYSfolder + "\radilav.cfg"))
            'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Else
            LineaError = "000-0341"
            EscribirArch1Linea SYSfolder + "\radilav.cfg", "0"
            LineaError = "000-0342"
            CreditosValidar = 0
            LineaError = "000-0343"
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el m�ximo y si hay que avisar
        LineaError = "000-0344"
        ValidarCada = LeerConfig("ValidarCada", "500")
        LineaError = "000-0345"
        AvisarAntes = LeerConfig("AvisarAntes", "50")
        LineaError = "000-0346"
        If CreditosValidar > ValidarCada - AvisarAntes Then
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            
            'uso el frmClave que tiene la variable publica ClaveIngresada
            LineaError = "000-0347"
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = ClaveParaValidar(CodigoParaClaveActual)
            LineaError = "000-0348"
            Dim QuedanC As Long
            QuedanC = ValidarCada - CreditosValidar
            If QuedanC > 0 Then
                'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
                LineaError = "000-0349"
                MsgBox "Ingrese a continuacion su clave para continuar utilizando 3PM. " + vbCrLf + _
                    "Debe enviar la administrador el codigo: " + vbCrLf + _
                    CodigoParaClaveActual + vbCrLf + _
                    "Puede todavia omitir esta clave. Solo le quedan " + CStr(QuedanC) + " creditos hasta que 3PM se inhabilite"
            Else
                LineaError = "000-0350"
                MsgBox "De no ingresar la clave correspondiente 3PM no podra continuar. Ha llegado al limite de creditos posibles"
            End If
            LineaError = "000-0351"
            frmCLAVE.Show 1
            LineaError = "000-0352"
            If UCase(ClaveIngresada) <> UCase(ClaveCorrespondiente) Then
                LineaError = "000-0353"
                If QuedanC > 0 Then
                    LineaError = "000-0354"
                    MsgBox "La clave es erronea!" + vbCrLf + _
                        "Le quedan " + CStr(QuedanC) + " creditos por cargar antes que se inhabilite 3PM"
                Else
                    LineaError = "000-0355"
                    MsgBox "No podra seguir utilizando 3PM hasta que valide con la clave correspondiente"
                    End
                End If
            Else
                LineaError = "000-0356"
                'todo OK. Cargo bien la clave
                CreditosValidar = 0
                LineaError = "000-0357"
                EscribirArch1Linea SYSfolder + "\radilav.cfg", "0"
                'empezar un nuevo periodo
                LineaError = "000-0358"
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        End If
        LineaError = "000-0359"
        lblValidar = "Val=" + CStr(ValidarCada) + "-Qued=" + CStr(ValidarCada - CreditosValidar) + "Actual=" + CStr(CreditosValidar) + " Codigo: " + CodigoParaClaveActual
    
    End If
    
    'caso especial Eduardo rodirguez
    If ClaveAdmin = "ERO77701192FF" Then frmIndex.lblTBR.Visible = False
    If ClaveAdmin = "MARC777" Then frmIndex.lblTBR.Visible = False
    
    'ver que onda con la publicidad de imagenes
    tbrPassImg1.ActivarPUBS = MostrarPUBIMG
    tbrPassImg1.IntervalBetwenIMGs = PubliIMGCada
    tbrPassImg1.ClearList
    'empiezan en 1 ambos!!!
    Dim AA As Long
    For AA = 1 To PUBs.TotalPUBsIMG
        tbrPassImg1.AddArchivoIMG (PUBs.ArchsPubsIMG(AA))
    Next
    tbrPassImg1.IniciarPASS
    
    Exit Sub
NoLoadIndex:
ErrMP3:
    MsgBox Err.Description + " N�: " + Str(Err.Number) + vbCrLf + "LINEA: " + LineaError
    WriteTBRLog "LINEA: " + LineaError + vbCrLf + Err.Description + " N�: " + Str(Err.Number), True
    Resume Next
End Sub

Public Sub SelDisco(nDisco As Long)
    
    LineaError = "000-0360"
    lblSel.Visible = False
    LineaError = "000-0361"
    lblDisco(nDisco).ForeColor = vbBlack
    LineaError = "000-0362"
    'lblDISCO(nDisco).Font.Bold = True
    lblDisco(nDisco).Font.Underline = True
    LineaError = "000-0363"
    lblDisco(nDisco).BackColor = vbYellow
    LineaError = "000-0364"
    nDiscoSEL = nDisco
    LineaError = "000-0365"
    lblSel.Top = TapaCD(nDiscoSEL).Top - lblSel.BorderWidth * 10
    LineaError = "000-0366"
    lblSel.Left = TapaCD(nDiscoSEL).Left - lblSel.BorderWidth * 10
    LineaError = "000-0367"
    lblSel.Height = TapaCD(nDiscoSEL).Height + lblSel.BorderWidth * 20
    LineaError = "000-0368"
    lblSel.Width = TapaCD(nDiscoSEL).Width + lblSel.BorderWidth * 20
    LineaError = "000-0369"
    lblSel.Visible = True
    LineaError = "000-0370"
    lblSel.ZOrder
    LineaError = "000-0371"
    lblDisco(nDisco).ZOrder
    
    'seleccionar de la lista de solo video
    LineaError = "000-0372"
    L(nDiscoGral).ForeColor = vbWhite
    LineaError = "000-0373"
    L(nDiscoGral).BackColor = vbBlack
    LineaError = "000-0374"
    LastDiscoSel = nDiscoGral 'para saber cual desactivar en unsel
    LineaError = "000-0375"
    If CargarIMGinicio Then
        TapaCD(nDiscoGral).BorderStyle = 1
    Else
        TapaCD(nDisco).BorderStyle = 1
    End If
    'imgARROW.Top = TapaCD(nDiscoGral).Top + TapaCD(nDiscoGral).Height - imgARROW.Height
    'imgArrow2.Top = TapaCD(nDiscoGral).Top + TapaCD(nDiscoGral).Height - imgArrow2.Height
    'imgARROW.Left = TapaCD(nDiscoGral).Left
    'imgArrow2.Left = TapaCD(nDiscoGral).Left + TapaCD(nDiscoGral).Width - imgARROW.Width
    'imgARROW.ZOrder
    'imgArrow2.ZOrder
    LineaError = "000-0376"
    If EsVideo Then OrdenarListaModoVideo
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    LineaError = "000-0377"
    lblDisco(nDisco).ForeColor = vbWhite
    'lblDISCO(nDisco).Font.Bold = False
    LineaError = "000-0378"
    lblDisco(nDisco).Font.Underline = False
    LineaError = "000-0379"
    lblDisco(nDisco).BackColor = vbBlack
    'seleccionar de la lista de solo video
    LineaError = "000-0380"
    L(LastDiscoSel).ForeColor = vbBlack
    LineaError = "000-0381"
    L(LastDiscoSel).BackColor = vbWhite
    LineaError = "000-0382"
    If CargarIMGinicio Then
        TapaCD(LastDiscoSel).BorderStyle = 0
    Else
        TapaCD(nDisco).BorderStyle = 0
    End If
    LineaError = "000-0383"
    If EsVideo Then OrdenarListaModoVideo
End Sub

Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean, DeQueFila As Long) As Long
    'indicando en que disco se inicia carga ese y los seis (o lo que corresponde) que le sigen
    'DeQueFial dice si es primero o �ltimo de cual fila!!!
    'devuelve el n�mero de discos cargados
    LineaError = "000-0384"
    Dim mCargarDiscos As Long
    mCargarDiscos = 0
    LineaError = "000-0385"
    Dim TotPags As Long
    TotPags = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
    LineaError = "000-0386"
    lblPag = "Pagina " + CStr(Round(numDiscoIniciar / (TapasMostradasH * TapasMostradasV) + 1, 0)) + " de " + CStr(TotPags + 1)
    'tomar el disco que va a quedar seleccionado
    'como numero de disco en el indice general
    If SelPrimero Then
        LineaError = "000-0387"
        'si la fila es uno (la primera) entonces el calculo es facil
        nDiscoGral = numDiscoIniciar + ((DeQueFila - 1) * TapasMostradasH)
    Else
        LineaError = "000-0388"
        nDiscoGral = numDiscoIniciar + ((TapasMostradasH * DeQueFila) - 1)
        'nDiscoGral = numDiscoIniciar + ((TapasMostradasH * TapasMostradasV) - 1)
        'era un 5, o sea total tapas-1
    End If
    'esconder todos los discos
    Dim NDR As Long 'numero de tapa de disco real del 0 al 5 (total de discos-1)
    
    'no hacer esto al pedo si ya estan cargadas
    Dim NDI As Long '=numdiscoiniciar de la pagina
    Dim C As Integer
    C = 1
    LineaError = "000-0389"
    NDI = numDiscoIniciar
    LineaError = "000-0390"
    If CargarIMGinicio Then
        If SelPrimero Then
            'si voy para adelante ocultar los que ya pase
            C = 1
            LineaError = "000-0391"
            Do While C <= (TapasMostradasH * TapasMostradasV)
                LineaError = "000-0392"
                'si no es la primera hoja!!
                If NDI >= (TapasMostradasH * TapasMostradasV) Then
                    LineaError = "000-0393"
                    TapaCD(NDI - C).Visible = False
                    'no se cargan lbldisco, usan solo del 0 al 5
                    LineaError = "000-0394"
                    lblDisco(C - 1).Visible = False
                End If
                C = C + 1
            Loop
            LineaError = "000-0395"
            Me.Refresh
        Else
            'sino ocultar los de adelante
            C = 1
            LineaError = "000-0396"
            Do While C <= (TapasMostradasH * TapasMostradasV)
                LineaError = "000-0397"
                If NDI + ((TapasMostradasH * TapasMostradasV) - 1) + C < UBound(MATRIZ_DISCOS) Then TapaCD(NDI + ((TapasMostradasH * TapasMostradasV) - 1) + C).Visible = False
                LineaError = "000-0398"
                lblDisco(C - 1).Visible = False
                C = C + 1
            Loop
            'Me.Refresh
        End If
    Else
        'si no se cargaron al inicio!!
        LineaError = "000-0399"
        Do While NDR < ((TapasMostradasH * TapasMostradasV))
            LineaError = "000-0400"
            TapaCD(NDR).Visible = False
            LineaError = "000-0401"
            lblDisco(NDR).Visible = False
            LineaError = "000-0402"
            NDR = NDR + 1
        Loop
        Dim ArchTapa As String
    End If
    NDR = 0
    LineaError = "000-0403"
    Do While NDI < numDiscoIniciar + ((TapasMostradasH * TapasMostradasV))
        'ver si existe si hay disco con este n�
        LineaError = "000-0404"
        If NDI < UBound(MATRIZ_DISCOS) Then
            LineaError = "000-0405"
            mCargarDiscos = mCargarDiscos + 1
            'ver si ya estan cargadas o se deben cargar
            LineaError = "000-0406"
            If CargarIMGinicio Then
                LineaError = "000-0407"
                TapaCD(NDI).Visible = True
                LineaError = "000-0408"
                TapaCD(NDI).ZOrder
            Else
                'ver si hay tapa
                LineaError = "000-0409"
                ArchTapa = txtInLista(MATRIZ_DISCOS(NDI + 1), 0, ",")
                LineaError = "000-0410"
                If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
                LineaError = "000-0411"
                ArchTapa = ArchTapa + "tapa.jpg"
                LineaError = "000-0412"
                If FSO.FileExists(ArchTapa) Then
                    LineaError = "000-0413"
                    TapaCD(NDR).Picture = LoadPicture(ArchTapa)
                Else
                    LineaError = "000-0414"
                    TapaCD(NDR).Picture = LoadPicture(SYSfolder + "f8ya.nam")
                End If
                LineaError = "000-0415"
                TapaCD(NDR).Visible = True
            End If
            'poner nombre al disco
            LineaError = "000-0416"
            lblDisco(NDR) = txtInLista(MATRIZ_DISCOS(NDI + 1), 1, ",")
            LineaError = "000-0417"
            If MostrarRotulos Then lblDisco(NDR).Visible = True
        End If
        LineaError = "000-0418"
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    CargarDiscos = mCargarDiscos
    If SelPrimero Then
        LineaError = "000-0419"
        'si es modo 46 no me importa la fila!!!!
        If IsMod46Teclas = 46 Then
            UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
        Else
            'si viene de una fila que no es la �ltima!!!!!!
            UnSelDisco (TapasMostradasH - 1) + ((DeQueFila - 1) * TapasMostradasH)
        End If
        
        LineaError = "000-0420"
        'si va a la primera fila queda en cero. JOIA
        'pero si existe la hoja y no el disco en esa fila
        'o sea la hoja tiene solo el primer disco y yo vengo
        'de la segunda fila !!!!!!!!
        ' y si esta despues del ultimo!!!!!!!!!
        If nDiscoGral >= TOTAL_DISCOS Then
            'tener en cuenta el nDiscoGral!!!!!!!!
            nDiscoGral = TOTAL_DISCOS - 1
            'elegir el ultimo que haya!!!
            'no el ultimo de la pagina bestia!!!!!
            'SelDisco (TapasMostradasV * TapasMostradasH) - 1
            'JOIA'JOIA'JOIA'JOIA'JOIA'JOIA
            SelDisco mCargarDiscos - 1
        Else
            SelDisco (DeQueFila - 1) * TapasMostradasH
        End If
        
    Else
        'si viene de una pagina de adelante para atras....
        LineaError = "000-0421"
        'si es modo 46 no me importa la fila!!!!
        If IsMod46Teclas = 46 Then
            UnSelDisco 0
            'SelDisco ((TapasMostradasH * TapasMostradasV) - 1)
        Else
            LineaError = "000-0422"
            'tiene que desseleccionar el que ven�a !!
            UnSelDisco (DeQueFila - 1) * TapasMostradasH
            
        End If
        SelDisco ((TapasMostradasH * DeQueFila) - 1)
        
    End If
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Constante Valor Descripci�n
'vbFormControlMenu 0 El usuario eligi� el comando Cerrar del men� Control del formulario.
'vbFormCode 1 Se invoc� la instrucci�n Unload desde el c�digo.
'vbAppWindows 2 La sesi�n actual del entorno operativo Microsoft Windows est� finalizando.
'vbAppTaskManager 3 El Administrador de tareas de Microsoft Windows est� cerrando la aplicaci�n.
'vbFormMDIForm 4 Un formulario MDI secundario se est� cerrando porque el formulario MDI tambi�n se est� cerrando.
'vbFormOwner 5 Un formulario se est� cerrando por que su formulario propietario se est� cerrando

    'Select Case UnloadMode
    '    Case 0
    '        MsgBox "El usuario eligi� el comando Cerrar del men� Control " + _
    '            "del formulario."
    '    Case 1
    '        MsgBox "Se invoc� la instrucci�n Unload desde el c�digo."
    '    Case 2
    '        MsgBox "La sesi�n actual del entorno operativo Microsoft Windows " + _
    '            "est� finalizando."
    '    Case 3
    '        MsgBox "El Administrador de tareas de Windows est� cerrando la " + _
    '           "aplicaci�n."
    '    Case 4
    '        MsgBox "Un formulario MDI secundario se est� cerrando porque " + _
    '            "el formulario MDI tambi�n se est� cerrando."
    '    Case 5
    '        MsgBox "Un formulario se est� cerrando por que su formulario " + _
    '            "propietario se est� cerrando"
    'End Select
    
    LineaError = "000-0423"
    MostrarCursor True
    LineaError = "000-0425"
    'MP3.DoStop EL DOsTOP GENERA EL EVENTO ENDPLAY QUE EJECUTA EL QUE SIGUE!!!
    LineaError = "000-0426"
    MP3.DoClose
    LineaError = "000-0427"
    If Is3pmExclusivo Then
        VU21.DoStop
    Else
        VU1.DoStop
    End If
    'esta es para rigoberto!!!!
    End
End Sub

Private Sub MP3_BeginPlay()
    Dim Tapa As String
    Tapa = FSO.GetParentFolderName(MP3.FileName) + "\tapa.jpg"
    If FSO.FileExists(Tapa) Then
        TapaEjecutando.Picture = LoadPicture(Tapa)
    Else
        TapaEjecutando.Picture = LoadPicture(SYSfolder + "f8ya.nam")
    End If
    LineaError = "000-0428"
    TotalTema = MP3.LengthInSec
    LineaError = "000-0429"
    Ancho = lblTemaSonando.Width
    'EVITAR DIVISIONES POR CERO
    LineaError = "000-0430"
    If TotalTema > 0 And MP3.IsPlaying Then
        LineaError = "000-0431"
        Variacion = Ancho / TotalTema
        LineaError = "000-0432"
        lblTiempoRestante = "TOTAL: " + MP3.Falta
    Else
        LineaError = "000-0433"
        lblTiempoRestante = "FALTA: " + "00:00"
    End If
    LineaError = "000-0434"
    VolBajando = MP3.Volumen
    
    Prog.Clear
    Prog.MAX = MP3.LengthInSec
    
End Sub

Private Sub MP3_EndPlay()
    EstoyEnModoVideoMiniSelDisco = False
    frmIndex.TapaEjecutando.Picture = LoadPicture(SYSfolder + "f8ya.nam")
    'volver a PasarHoja a su estado original3
    LineaError = "000-0435"
    PasarHoja = LeerConfig("PasarHoja", "1")
    LineaError = "000-0436"
    VU1.Width = Screen.Width
    LineaError = "000-0437"
    
    'ver si es fullscreen o no!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If vidFullScreen Then
        'frDISCOS.Height = picFondo.Top
        VU1.Height = picFondo.Top
    End If
    'reacomodo si vengo de video minimo
    'tener el cuenta el exclusivo!!!!!!!!!!!!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        LineaError = "000-0438"
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width - VU1.Width
        LineaError = "000-0439"
        frDISCOS.Left = VU1.AnchoBarra + 25 ' VU1.Width
        'vu no se mueve si termina un video        'VU1.Top = 0        'VU1.Height = Me.Height
    Else
        LineaError = "000-0440"
        frDISCOS.Width = VU1.Width ' Screen.Width
        LineaError = "000-0441"
        frDISCOS.Left = 0
    End If
    
    picFondoDisco.Height = frDISCOS.Height
    picFondoDisco.Width = frDISCOS.Width
    frModoVideo.Visible = False
    lblModoVideo.Visible = False
    frTEMAS.Visible = False
    lblTEMAS.Visible = False
    ModoVideoSelTema = False
    LBLpORCtEMA.Width = Ancho
    'termino una cancion
    If EsVideo Then MP3.DoClose
    'lo destapo al terminar de acomodar todos los controles en otro lado
    'picVideo.Visible = False
    EMPEZAR_SIGUIENTE
    lblREP.BackStyle = 0
    lblREP.ForeColor = vbWhite
    lblREP = ""
    
End Sub

Private Sub MP3_Played(SecondsPlayed As Long)
    
    lblREP.Caption = "Reproduciendo:"
    If SecondsPlayed / 2 = SecondsPlayed \ 2 Then
        lblREP.BackStyle = 1
        lblREP.BackColor = vbYellow
        lblREP.ForeColor = vbBlack
    Else
        lblREP.BackStyle = 0
        lblREP.ForeColor = vbWhite
    End If
    
    'esto pasa cada un segundo (si o si una vez por segundo)
    Dim sRest As Long
    LineaError = "000-0455"
    sRest = MP3.FaltaInSec
    LineaError = "000-0456"
    PorcEjecutado = MP3.PercentPlay
    LineaError = "000-0457"
    If PorcEjecutado > PorcentajeTEMA And CORTAR_TEMA Then
        LineaError = "000-0458"
        VolBajando = VolBajando - 5 'baja 1 por segundo
        LineaError = "000-0459"
        lblTemaSonando = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        lblTemaSonando2 = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        LineaError = "000-0460"
        If VolBajando > 0 Then
            LineaError = "000-0461"
            MP3.Volumen = VolBajando
        Else
            LineaError = "000-0462"
            MP3.DoStop
            'EL DOSTOP DESENCADENA UN END PLAY QUE REALIZA UN EMPEZAR SIGUINETE
            'EMPEZAR_SIGUIENTE
        End If
    End If
    LineaError = "000-0463"
    lblTiempoRestante = "FALTA: " + MP3.Falta
    Prog.DibujarCirculo CDbl(SecondsPlayed)
    
    LineaError = "000-0464"
    wi = Ancho - Variacion * (SecondsPlayed - 2)
    LineaError = "000-0465"
    If wi > 0 Then LBLpORCtEMA.Width = wi
    '=====================================
    LineaError = "000-0466"
    If K.LICENCIA = aSinCargar And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        LineaError = "000-0467"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        LineaError = "000-0468"
        MP3.DoStop
    End If
    'cotar tambin en el gratuito
    LineaError = "000-0469"
    If K.LICENCIA = CGratuita And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        LineaError = "000-0470"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        LineaError = "000-0471"
        MP3.DoStop
    End If
    '=====================================
End Sub

Private Sub TapaCD_Click(Index As Integer)
    'nunca hay que pasar hojas
    'nDiscoGral = nDiscoGral + (Index - nDiscoSEL)
    LineaError = "000-0473"
    nDiscoGral = Index 'si se cargan todas las im�genes al inicio index=nDiscoGral
    LineaError = "000-0474"
    If nDiscoGral + 1 > TOTAL_DISCOS Then
        LineaError = "000-0475"
        MsgBox "No existe el disco elegido!!. " + vbCrLf + _
            "Carge discos desde el ADMINISTRADOR DE DISCOS en la " + vbCrLf + _
            "p�gina de configuracion (presionando la tecla 'C')"
        LineaError = "000-0476"
        Exit Sub
    End If
    LineaError = "000-0477"
    UnSelDisco nDiscoSEL
    LineaError = "000-0478"
    Dim PagNum As Long
    PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
    LineaError = "000-0479"
    nDiscoSEL = Index - (PagNum * (TapasMostradasH * TapasMostradasV))
    LineaError = "000-0480"
    SelDisco nDiscoSEL
    LineaError = "000-0481"
    lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
    LineaError = "000-0482"
    'totar la tecla de enar a disco
    Form_KeyDown TeclaOK, 0
End Sub

Private Sub tbrPassImg1_ChangeImg()
    'si se esta pasando un video no dar bola!!!
    If MP3.IsPlaying And EsVideo Then
        frmVIDEO.picBigImg.Visible = False
    Else
        frmVIDEO.picBigImg.Visible = False
        
        'cambiar tambien las im�genes grandes de la salida de video
        PUBs.UltimaReproducidaBigIMG = PUBs.UltimaReproducidaBigIMG + 1
        'si me paso se va al primero ya
        If PUBs.UltimaReproducidaBigIMG > PUBs.TotalPUBsBigIMG Then PUBs.UltimaReproducidaBigIMG = 1
        '...
        '...
        'aca debe ir algun efecto. Ponete las pilas ANDRES
        '...
        '...
        With frmVIDEO.picBigImg
            .Picture = LoadPicture(PUBs.ArchsPubsBigIMG(PUBs.UltimaReproducidaBigIMG))
            .Top = Me.Height / 2 - .Height / 2
            .Left = Me.Width / 2 - .Width / 2
            .Visible = True
        End With
    End If
End Sub

Private Sub Timer1_Timer()
    'controla el tiempo sin uso (sin ejecucion de temas)
    LineaError = "000-0483"
    If MP3.IsPlaying Then Exit Sub
    'controla el tiempo sin uso (sin ejecucion de temas)
    LineaError = "000-0484"
    SecSinUso = SecSinUso + (Timer1.Interval / 1000)
    LineaError = "000-0485"
    lblNoUSO = Trim(Str(SecSinUso))
    LineaError = "000-0486"
    If SecSinUso >= EsperaMinutos Then 'esperaminutos esta en segundos
        LineaError = "000-0487"
        SecSinUso = 0
        LineaError = "000-0488"
        Dim TemasDisponibles As Long
        If TemasEnRank(1) > 50 Then
            LineaError = "000-0489"
            TemasDisponibles = TemasEnRank(1) 'todos los que se escucharon
        Else
            LineaError = "000-0490"
            TemasDisponibles = TemasEnRank(0) 'todos los que se escucharon
        End If
        LineaError = "000-0491"
        Randomize Timer
        LineaError = "000-0492"
        z = Int(Rnd * TemasDisponibles)
        z = z + 1
        CC = 0
        LineaError = "000-0493"
        If FSO.FileExists(AP + "ranking.tbr") = False Then
            LineaError = "000-0494"
            FSO.CreateTextFile AP + "ranking.tbr", True
            'me voy al azar ya que no hay para elegirdel rank
            LineaError = "000-0495"
            GoTo AZAR
        End If
        LineaError = "000-0496"
        Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
        LineaError = "000-0497"
        Dim TT As String
        'antes de entra ver si el archivo no tiene nada
        LineaError = "000-0498"
        If TE.AtEndOfStream Then GoTo AZAR
        LineaError = "000-0499"
        Do While Not TE.AtEndOfStream
            LineaError = "000-0500"
            CC = CC + 1
            LineaError = "000-0501"
            TT = TE.ReadLine
            LineaError = "000-0502"
            If CC = z Then
                LineaError = "000-0503"
                Dim TemaAzar As String
                LineaError = "000-0504"
                TemaAzar = txtInLista(TT, 1, ",")
                'si tuve los discos cargados en una unidad o una ubicaci�n distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                LineaError = "000-0505"
                If FSO.FileExists(TemaAzar) Then
                    LineaError = "000-0506"
                    CORTAR_TEMA = True 'este tema se eligio al azar no va entero
                    LineaError = "000-0507"
                    SecSinUso = 0
                    LineaError = "000-0508"
                    TE.Close
                    LineaError = "000-0509"
                    EjecutarTema TemaAzar, False
                    LineaError = "000-0510"
                    Exit Sub
                Else
AZAR:
                    'ejecutar algun tema de cualquier disco
                    LineaError = "000-0511"
                    Dim MTX10() As String: zz = 0
                    LineaError = "000-0512"
                    ruta = AP + "discos\"
                    LineaError = "000-0513"
                    Dim NombreDir As String
                    LineaError = "000-0514"
                    NombreDir = Dir$(ruta & "*.*", vbDirectory)
                    LineaError = "000-0515"
                    Do While Len(NombreDir)
                        LineaError = "000-0516"
                        If NombreDir = "." Or NombreDir = ".." Then
                            ' excluir las entradas "." y ".."
                            LineaError = "000-0517"
                        ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                            ' este es un archivo normal
                            LineaError = "000-0518"
                        Else
                            LineaError = "000-0519"
                            'ver los primeros diez discos. En alguno tiene que haber temas
                            'yo se que el primero no tiene temas por que es
                            '01 - los mas escuchados
                            LineaError = "000-0520"
                            ReDim Preserve MTX10(zz) As String
                            LineaError = "000-0521"
                            MTX10(zz) = ruta & NombreDir
                            LineaError = "000-0522"
                            zz = zz + 1
                        End If
                        LineaError = "000-0523"
                        NombreDir = Dir$
                    Loop
BuscaMP3:
                    LineaError = "000-0524"
                    'siempre cae en el primer tema del primer directorio habilitado
                    Randomize Timer
                    Dim A As Integer, ContA As Integer
                    LineaError = "000-0525"
                    A = Int(Rnd * 1000) + 1
                    LineaError = "000-0526"
                    Dim NombreMP3 As String: zz = 0
                    LineaError = "000-0527"
                    Dim temaMP As String
                    LineaError = "000-0528"
                    Do While zz < UBound(MTX10)
                        LineaError = "000-0529"
                        NombreMP3 = Dir$(MTX10(zz) & "\*.mp3")
                        'si no hay ningun tema se va a la prox carpeta
                        LineaError = "000-0530"
                        If NombreMP3 = "" Then GoTo NextFolder
                        'da vueltas hasta encontrar un tema valido
                        LineaError = "000-0531"
                        Do While Len(NombreMP3)
                            LineaError = "000-0532"
                            temaMP = MTX10(zz) & "\" & NombreMP3
                            LineaError = "000-0533"
                            If FSO.FileExists(temaMP) Then
                                LineaError = "000-0534"
                                ContA = ContA + 1
                                LineaError = "000-0535"
                                If ContA >= A Then
                                    LineaError = "000-0536"
                                    CORTAR_TEMA = True 'este tema va cortado ya que es de 3PM para que haga ruido
                                    LineaError = "000-0537"
                                    EjecutarTema temaMP, False
                                    'solo sale cueando encuentra un tema valido
                                    LineaError = "000-0538"
                                    SecSinUso = 0
                                    Exit Sub
                                End If
                            End If
                            LineaError = "000-0539"
                            NombreMP3 = Dir$
                        Loop
NextFolder:
                        zz = zz + 1
                    Loop
                End If
                Exit Do
            End If
         Loop
         LineaError = "000-0540"
         TE.Close
        'si llego aca es por que no encontro el numero sorteado al azar en la lista
        'de los mejores. Entonces elige un tema al azar
        LineaError = "000-0541"
        GoTo AZAR
    End If
    
End Sub

Private Sub Timer3_Timer()
    LineaError = "000-0542"
    If Protector = 0 Then Timer3.Interval = 0        'para el reloj del protector. Lo ha inhabilitado
    LineaError = "000-0543"
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + 10
    lblNoTecla = Trim(Str(SecSinTecla))
    'no protector en video
    LineaError = "000-0544"
    If EsVideo Then SecSinTecla = 0
    LineaError = "000-0545"
    If SecSinTecla > EsperaTecla And EsVideo = False Then
        LineaError = "000-0546"
        frmProtect.Show 1
    End If
End Sub

Public Function TemasEnRank(MasDeXVotos) As Long
    'indica cuantos temas hay en el ranking
    LineaError = "000-0547"
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        LineaError = "000-0548"
        FSO.CreateTextFile AP + "ranking.tbr", True
        LineaError = "000-0549"
        TemasEnRankMasDeUnVoto = 0
        Exit Function
    End If
    LineaError = "000-0550"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    LineaError = "000-0551"
    Dim TT As String
    'antes de entra ver si el archivo no tiene nada
    LineaError = "000-0552"
    If TE.AtEndOfStream Then
        LineaError = "000-0553"
        TemasEnRankMasDeUnVoto = 0
        LineaError = "000-0554"
        TE.Close
        LineaError = "000-0555"
        Exit Function
    End If
    Dim CA As Long
    CA = 0
    Dim PuntosEste  As Long
    LineaError = "000-0556"
    Do While Not TE.AtEndOfStream
        LineaError = "000-0557"
        TT = TE.ReadLine
        LineaError = "000-0558"
        PuntosEste = Val(txtInLista(TT, 0, ","))
        LineaError = "000-0559"
        If PuntosEste > MasDeXVotos Then
            LineaError = "000-0560"
            CA = CA + 1
        Else
            'todos los que siguen tienen uno (1)
            LineaError = "000-0561"
            Exit Do
        End If
    Loop
    LineaError = "000-0562"
    TE.Close
    LineaError = "000-0563"
    TemasEnRank = CA
End Function

Public Sub OrdenarListaModoVideo()
    'asegurarme que el disco elegido se ve en la lista
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    LineaError = "000-0564"
    If L(nDiscoGral).Top > frModoVideo.Height - (L(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        LineaError = "000-0565"
        HayQueCorrerse = L(nDiscoGral).Top - (frModoVideo.Height - (L(0).Height + 25))
        LineaError = "000-0566"
        CL = 0
        Do While CL < TOTAL_DISCOS
            LineaError = "000-0567"
            L(CL).Top = L(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    LineaError = "000-0568"
    If L(nDiscoGral).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        LineaError = "000-0569"
        HayQueCorrerse = -L(nDiscoGral).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        LineaError = "000-0570"
        Do While CL < TOTAL_DISCOS
            LineaError = "000-0571"
            L(CL).Top = L(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
End Sub

Public Sub SelTema(n As Integer)
    LineaError = "000-0571"
    T(n).BackColor = &H0&
    LineaError = "000-0572"
    T(n).ForeColor = &H80FFFF
End Sub

Public Sub UnSelTema(n As Integer)
    LineaError = "000-0573"
    T(n).BackColor = &H80FFFF
    LineaError = "000-0574"
    T(n).ForeColor = &H0&
End Sub

Public Sub OrdenarListaTemaVideo()
    'asegurarme que el disco elegido se ve en la lista
    LineaError = "000-0575"
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    LineaError = "000-0576"
    If T(TemaElegidoModoVideo).Top > frTEMAS.Height - (T(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que correrse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        LineaError = "000-0577"
        HayQueCorrerse = T(TemaElegidoModoVideo).Top - (frTEMAS.Height - (T(0).Height + 25))
        LineaError = "000-0578"
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            LineaError = "000-0579"
            T(CL).Top = T(CL).Top - HayQueCorrerse
            LineaError = "000-0580"
            CL = CL + 1
        Loop
    End If
    LineaError = "000-0581"
    If T(TemaElegidoModoVideo).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        LineaError = "000-0581"
        HayQueCorrerse = -T(TemaElegidoModoVideo).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        LineaError = "000-0582"
        Do While CL <= UBound(MATRIZ_TEMAS)
            LineaError = "000-0583"
            T(CL).Top = T(CL).Top + HayQueCorrerse
            LineaError = "000-0584"
            CL = CL + 1
        Loop
    End If
    
End Sub
