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
      TabIndex        =   35
      Top             =   4590
      Visible         =   0   'False
      Width           =   4440
   End
   Begin tbr3pm.VUMeter2 VU21 
      Height          =   1275
      Left            =   60
      TabIndex        =   36
      Top             =   5550
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   2249
      Begin tbr3pm.tbrProgressCircle Prog 
         Height          =   465
         Left            =   765
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   45
         Width           =   2175
      End
   End
   Begin VB.Frame frDISCOS 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   4950
      Left            =   3510
      TabIndex        =   11
      Top             =   90
      Width           =   5295
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
            Left            =   690
            TabIndex        =   32
            Top             =   2820
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
      Left            =   60
      TabIndex        =   7
      Top             =   210
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   7805
   End
   Begin VB.PictureBox picFondo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4230
      ScaleWidth      =   15360
      TabIndex        =   13
      Top             =   6930
      Width           =   15420
      Begin tbr3pm.tbrPassImg tbrPassImg1 
         Height          =   1965
         Left            =   60
         TabIndex        =   33
         Top             =   30
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   3466
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "VERSION DEMOSTRATIVA. tbrSoft Argentina. www.tbrsoft.com"
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
            Height          =   885
            Left            =   330
            TabIndex        =   34
            Top             =   510
            Visible         =   0   'False
            Width           =   1845
         End
      End
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   2100
         Left            =   10140
         TabIndex        =   14
         Top             =   -60
         Width           =   1875
         Begin VB.CommandButton cmdPagAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":0442
            Height          =   715
            Left            =   960
            Picture         =   "frmINDEX.frx":1403
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1365
            Width           =   900
         End
         Begin VB.CommandButton cmdDiscoAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":1E10
            Height          =   715
            Left            =   960
            Picture         =   "frmINDEX.frx":2B0D
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   630
            Width           =   900
         End
         Begin VB.CommandButton cmdDiscoAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":33E5
            Height          =   715
            Left            =   60
            Picture         =   "frmINDEX.frx":4157
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   630
            Width           =   900
         End
         Begin VB.CommandButton cmdPagAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmINDEX.frx":4A9A
            Height          =   715
            Left            =   60
            Picture         =   "frmINDEX.frx":5AF9
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1365
            Width           =   900
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "OK"
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
            Left            =   60
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   90
            Width           =   1875
         End
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
         Left            =   2730
         TabIndex        =   24
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   7395
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
         Left            =   2730
         TabIndex        =   30
         Top             =   60
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblTBR 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         Caption         =   "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2775
         TabIndex        =   29
         Top             =   1470
         Width           =   7350
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
         Left            =   2790
         TabIndex        =   28
         Top             =   1755
         UseMnemonic     =   0   'False
         Width           =   7350
      End
      Begin VB.Label lblTOTdiscos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   4770
         TabIndex        =   27
         Top             =   360
         Width           =   2145
      End
      Begin VB.Label lblTiempoRestante 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   6870
         TabIndex        =   26
         Top             =   360
         Width           =   1845
      End
      Begin VB.Label lblPag 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   10020
         TabIndex        =   25
         Top             =   360
         Width           =   1980
      End
      Begin VB.Label lblCreditos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Creditos: 00"
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
         Height          =   375
         Left            =   2730
         TabIndex        =   23
         Top             =   360
         Width           =   2025
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
         Left            =   2730
         TabIndex        =   22
         Top             =   810
         Width           =   2040
      End
      Begin VB.Label lblPuesto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
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
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   8790
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblValidar 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Validar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   6090
         TabIndex        =   20
         Top             =   90
         Visible         =   0   'False
         Width           =   3255
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
         Height          =   795
         Left            =   4800
         TabIndex        =   31
         Top             =   630
         UseMnemonic     =   0   'False
         Width           =   5325
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
            Name            =   "Arial"
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
            Name            =   "Arial"
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
         Name            =   "Arial"
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
        CaminoError "000-0001"
        Form_KeyDown TeclaDER, 0
        CaminoError "000-0002"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    CaminoError "000-0003"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdDiscoAt_Click()
    If MostrarTouch Then
        CaminoError "000-0004"
        Form_KeyDown TeclaIZQ, 0
        CaminoError "000-0005"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
    CaminoError "000-0006"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAd_Click()
    If MostrarTouch Then
        CaminoError "000-0007"
        Form_KeyDown TeclaPagAd, 0
        CaminoError "000-0008"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
    CaminoError "000-0009"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub cmdPagAt_Click()
    If MostrarTouch Then
        CaminoError "000-0010"
        Form_KeyDown TeclaPagAt, 0
        CaminoError "000-0011"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
    CaminoError "000-0012"
    Form_KeyDown KeyCode, Shift
End Sub

Private Sub Command1_Click()
    If MostrarTouch Then
        CaminoError "000-0023"
        Form_KeyDown TeclaOK, 0
    End If
End Sub

Private Sub Form_Activate()
    CaminoError "000-0024"
    MostrarCursor False
    'actualizar los precios
    CaminoError "000-0025"
    '---------------------
    'si es gratis no usar!
    If CreditosCuestaTema = 0 And CreditosCuestaTemaVIDEO = 0 Then
        lblPrecios = "Modo Gratuito"
        lblPrecios2 = "Modo Gratuito"
    Else
        If TemasPorCredito = 1 Then
            CaminoError "000-0026"
            lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " credito"
            lblPrecios2 = "1 coin = " + CStr(TemasPorCredito) + " credito"
        Else
            CaminoError "000-0027"
            lblPrecios = "1 coin = " + CStr(TemasPorCredito) + " creditos"
            lblPrecios2 = "1 coin = " + CStr(TemasPorCredito) + " creditos"
        End If
    End If
    '-------------------------
    CaminoError "000-0028"
    If CreditosCuestaTema = 1 Then
        CaminoError "000-0029"
        lblPrecios = lblPrecios + vbCrLf + "1 credito = 1 tema"
        lblPrecios2 = lblPrecios2 + " / " + "1 credito = 1 tema"
    Else
        If CreditosCuestaTema = 0 Then
            lblPrecios = lblPrecios + vbCrLf + "1 tema = GRATIS!"
            lblPrecios2 = lblPrecios2 + " / " + " 1 tema = GRATIS!"
        Else
            CaminoError "000-0030"
            lblPrecios = lblPrecios + vbCrLf + CStr(CreditosCuestaTema) + " creditos = 1 tema"
            lblPrecios2 = lblPrecios2 + " / " + CStr(CreditosCuestaTema) + " creditos = 1 tema"
        End If
    End If
    'agreagr el precio de los videos!!!
    If CreditosCuestaTemaVIDEO = 1 Then
        CaminoError "000-0029"
        lblPrecios = lblPrecios + vbCrLf + "1 credito = 1 VIDEO"
        lblPrecios2 = lblPrecios2 + " / " + "1 credito = 1 VIDEO"
    Else
        If CreditosCuestaTemaVIDEO = 0 Then
            lblPrecios = lblPrecios + vbCrLf + "1 VIDEO = GRATIS!"
            lblPrecios2 = lblPrecios2 + " / " + " 1 VIDEO = GRATIS!"
        Else
            CaminoError "000-0030"
            lblPrecios = lblPrecios + vbCrLf + CStr(CreditosCuestaTemaVIDEO) + " creditos = 1 VIDEO"
            lblPrecios2 = lblPrecios2 + " / " + CStr(CreditosCuestaTemaVIDEO) + " creditos = 1 VIDEO"
        End If
    End If
    
    'total ser�a
    '1 coin = 8 creditos /// " + "8 creditos = 1 tema /// 8 creditos = 1 VIDEO
        
    CaminoError "000-0031"
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
    
    On Local Error GoTo FallaKD
    
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    CaminoError "000-0033"
    Dim PagNum As Long
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    
    'de manera predeterminada son el mismo
    'salvo los casos que se especifican
    RealKeyCode = KeyCode
    
    If IsKeyPad(Me) Then
        'la falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        If KeyCode = 13 Then RealKeyCode = 108
        'ademas si esta apretado el BLOQ NUM
    End If
           
    '----------------------------------------
    'esta tecla es IZQ en el modo 46 pasandpo de arriba aa abjo y _
        siguiendo a la pag ant en el modo 5
    'para el modo video y en modo46=5 se pasan como p�ginas!
    '----------------------------------------
    
    EsModo5PeroLabura46 = (EsVideo And _
        Salida2 = False And _
        IsMod46Teclas = 5)
    '----------------------------------------
           
    'Select Case KeyCode
    Select Case RealKeyCode
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
                End
            End If
        Case TeclaShowContador
            frmOnlyContador.Show 1
        Case TeclaPutCeroContador
            SumarContadorCreditos -CONTADOR 'esto lo deja en cero
            frmOnlyContador.Show 1
        Case TeclaFF 'avanzar 10 segundos
            Dim ToSec As Long
            ToSec = (MP3.PositionInSec * 1000) + 10000
            MP3.SeekTo CStr(ToSec)
        'subir o bajar volumen
        Case TeclaBajaVolumen
            CaminoError "000-0034"
            If frmIndex.MP3.IsPlaying Then
                CaminoError "000-0035"
                If VolumenIni <= 5 Then
                    CaminoError "000-0036"
                    frmIndex.MP3.Volumen = 0
                Else
                    CaminoError "000-0037"
                    frmIndex.MP3.Volumen = VolumenIni - 5
                End If
                CaminoError "000-0038"
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case TeclaSubeVolumen
            CaminoError "000-0039"
            If frmIndex.MP3.IsPlaying Then
                CaminoError "000-0039"
                If VolumenIni >= 95 Then
                    CaminoError "000-0040"
                    frmIndex.MP3.Volumen = 100
                Else
                    CaminoError "000-0041"
                    frmIndex.MP3.Volumen = VolumenIni + 5
                End If
                CaminoError "000-0042"
                VolumenIni = frmIndex.MP3.Volumen
            End If
        Case TeclaNextMusic
            'si es video ocultar la pantalla de video
            CaminoError "000-0043"
            'If EsVideo Then
            '    picVideo.Visible = False
            'End If
            CaminoError "000-0044"
            EMPEZAR_SIGUIENTE
        Case TeclaPagAd
            'pase lo que pase registrar
            CaminoError "000-0054"
            TECLAS_PRES = TECLAS_PRES + "5"
            CaminoError "000-0055"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0056"
            lblTECLAS = TECLAS_PRES
            
            'es para abajo en el modo 5 y pagina adelante de el modo 46
            
            If EsModo5PeroLabura46 Then
                'esto confirma que es modo 5
                Form_KeyDown TeclaDER, 0
            End If
            If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                'esta tecla es pagina adelante en el modo 46 y abajo en el modo 5
                CaminoError "000-0045"
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                CaminoError "000-0046"
                Dim PrimeroDeLaPaginaQueSigue As Long
                CaminoError "000-0047"
                PrimeroDeLaPaginaQueSigue = (PagNum + 1) * (TapasMostradasH * TapasMostradasV)
                
                'NUEVO DE 6.5, pasa a la primer p�gina
                If PrimeroDeLaPaginaQueSigue > TOTAL_DISCOS Then
                    PrimeroDeLaPaginaQueSigue = 0
                End If
                CaminoError "000-0048"
                'supongo que lo puse para que no desseleccione el mismo _
                    que va a seleccionar???
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                CaminoError "000-0050"
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueSigue, True, 1)
                CaminoError "000-0051"
                lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueSigue + 1) + " de " + CStr(TOTAL_DISCOS)
                CaminoError "000-0052"
                nDiscoSEL = 0
            End If
            'si esta eligiendo discos en modo video min es
            'totalmente desitinto, solo va al que sigue
            'no importann p�ginas ni nada
            'If EstoyEnModoVideoMiniSelDisco = False Then
            '    'xxxx
            '    Exit Sub
            'End If
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo As Long
                DiskToGo = nDiscoSEL + TapasMostradasH
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo < DiscosEnPagina Then
                    nDiscoGral = nDiscoGral + TapasMostradasH
                    CaminoError "000-0083"
                    UnSelDisco nDiscoSEL
                    CaminoError "000-0084"
                    SelDisco nDiscoSEL + TapasMostradasH
                End If
            End If
            
        Case TeclaPagAt
            If EsModo5PeroLabura46 Then
                'esto confirma que es modo 5
                Form_KeyDown TeclaIZQ, 0
            End If
            If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                'esta tecla es pagina atras en el modo 46 y arriba en el modo 5
                CaminoError "000-0056"
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                CaminoError "000-0057"
                
                CaminoError "000-0058"
                Dim PrimeroDeLaPaginaQueAnterior As Long
                'NUEVO DE 6.5, se va a la ultima pagina
                If PagNum > 0 Then
                    PrimeroDeLaPaginaQueAnterior = (PagNum - 1) * (TapasMostradasH * TapasMostradasV)
                Else
                    Dim tmpUbic2 As Long
                    'primero ver cuantas pags ENTERAS hay!
                    tmpUbic2 = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
                    'despues saber cuantos discos sobran la ultima pagina
                    tmpUbic2 = TOTAL_DISCOS - ((TapasMostradasH * TapasMostradasV) * tmpUbic2)
                    'ahora saber que posicion ocupa el primero de los que sobran el ultima p�gina
                    tmpUbic2 = TOTAL_DISCOS - tmpUbic2
                    PrimeroDeLaPaginaQueAnterior = tmpUbic2
                End If
                CaminoError "000-0059"
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                CaminoError "000-0060"
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueAnterior, False, TapasMostradasV)
                CaminoError "000-0061"
                lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            End If
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo2 As Long
                DiskToGo2 = nDiscoSEL - TapasMostradasH
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo2 >= 0 Then
                    nDiscoGral = nDiscoGral - TapasMostradasH
                    CaminoError "000-0083"
                    UnSelDisco nDiscoSEL
                    CaminoError "000-0084"
                    SelDisco nDiscoSEL - TapasMostradasH
                End If
            End If
            CaminoError "000-0062"
            TECLAS_PRES = TECLAS_PRES + "6"
            CaminoError "000-0063"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0064"
            lblTECLAS = TECLAS_PRES
        Case TeclaConfig
            CaminoError "000-0065"
             frmConfig.Show 1
        Case TeclaIZQ
            CaminoError "000-0066"
            If ModoVideoSelTema Then
                CaminoError "000-0067"
                If TemaElegidoModoVideo > 0 Then
                    CaminoError "000-0068"
                    UnSelTema TemaElegidoModoVideo
                    CaminoError "000-0069"
                    TemaElegidoModoVideo = TemaElegidoModoVideo - 1
                    CaminoError "000-0070"
                    SelTema TemaElegidoModoVideo
                    CaminoError "000-0071"
                    OrdenarListaTemaVideo
                End If
                GoTo FinTeclaZ
            End If
            'no ir a -1
            CaminoError "000-0072"
            'ver si es el primero
            If nDiscoSEL = 0 Then
                'ver si hay que pasar hoja o no
                CaminoError "000-0073"
                If PasarHoja Then
                    CaminoError "000-0074"
                    'ver si hay p�ginas antes
                    'si el gral es mayor que cero entonces si hay
                    'en la primera p�gina gral y discosel son iguales
                    If nDiscoGral > 0 Then
                        'como si viene eligiendo desde la ultima fila
                        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                            CaminoError "112CRG" + CStr(nDiscoGral - (TapasMostradasH * TapasMostradasV)) + "." + CStr(TapasMostradasV)
                            DiscosEnPagina = CargarDiscos(nDiscoGral - _
                            ((TapasMostradasH * TapasMostradasV)), False, TapasMostradasV)
                        End If
                        
                        'busca solo la fila!!
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            CaminoError "113CRG" + CStr(nDiscoGral - (TapasMostradasH * TapasMostradasV)) + "." + CStr(EnQueFilaEstoy)
                            DiscosEnPagina = CargarDiscos(nDiscoGral - _
                            ((TapasMostradasH * TapasMostradasV)), False, EnQueFilaEstoy)
                        End If
                    End If
                    
                    'NUEVO 6.5 si esta en el disco cero se va a la ultima hoja
                    'o sea se hace ciclico como mprock
                    If nDiscoGral = 0 Then
                        Dim tmpUbic As Long
                        'primero ver cuantas pags ENTERAS hay!
                        tmpUbic = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
                        'despues saber cuantos discos sobran la ultima pagina
                        tmpUbic = TOTAL_DISCOS - ((TapasMostradasH * TapasMostradasV) * tmpUbic)
                        'ahora saber que posicion ocupa el primero de los que sobran el ultima p�gina
                        tmpUbic = TOTAL_DISCOS - tmpUbic
                        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                            CaminoError "111CRG" + CStr(tmpUbic) + ".1"
                            DiscosEnPagina = CargarDiscos(tmpUbic, False, 1)
                        End If
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            CaminoError "110CRG" + CStr(tmpUbic) + "." + CStr(EnQueFilaEstoy)
                            DiscosEnPagina = CargarDiscos(tmpUbic, False, EnQueFilaEstoy)
                        End If
                    End If
                    
                Else
                    'NO NO NO!!!! nDiscoGral = (TapasMostradasH * TapasMostradasV) - 1
                    'estoy en una hoja al principio y debo elegir el disco del final
                    'sel y unsel trabajan con referencias de o al total de discos por pag
                    'nDiscoGral es el numero absoluto del disco
                    'ver si existe el disco al que voy
                    CaminoError "000-0075"
                    If TOTAL_DISCOS > nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1 Then
                        CaminoError "000-0076"
                        nDiscoGral = nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1
                        CaminoError "000-0077"
                        UnSelDisco nDiscoSEL
                        CaminoError "000-0078"
                        SelDisco (TapasMostradasH * TapasMostradasV) - 1
                    Else
                        CaminoError "000-0079"
                        nDiscoGral = TOTAL_DISCOS - 1
                        CaminoError "000-0080"
                        UnSelDisco nDiscoSEL
                        CaminoError "000-0081"
                        SelDisco DiscosEnPagina - 1
                    End If
                End If
            Else
                'si no es el primero ver si es
                'el primero de una fila y esta en modo 5 el teclado
                If nDiscoSEL = TapasMostradasH * (EnQueFilaEstoy - 1) Then
                    'si esta en el modo 5 me fijo si esta al final de una l�nea
                    If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                        'el disco a iniciar ya no es nDiscoGral-(tapash*tapasv)!!!!!!
                        'hay que restar tambien el nOrden de esta pagina
                        Dim DiscoToIni As Long
                        'el primero de esta mas el total de esta!
                        DiscoToIni = nDiscoGral - nDiscoSEL - (TapasMostradasH * TapasMostradasV)
                        'ver que no se vaya a la mierda!!
                        If DiscoToIni >= 0 Then
                            CaminoError "101CRG" + CStr(DiscoToIni) + "." + CStr(EnQueFilaEstoy)
                            DiscosEnPagina = CargarDiscos(DiscoToIni, False, EnQueFilaEstoy)
                        Else
                            Dim tmpUbic3 As Long
                            'primero ver cuantas pags ENTERAS hay!
                            tmpUbic3 = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
                            'despues saber cuantos discos sobran la ultima pagina
                            tmpUbic3 = TOTAL_DISCOS - ((TapasMostradasH * TapasMostradasV) * tmpUbic3)
                            'ahora saber que posicion ocupa el primero de los que sobran el ultima p�gina
                            tmpUbic3 = TOTAL_DISCOS - tmpUbic3
                            'no tengo tiempo de hacerlo ir a la mejor fila
                            'este es el caso de la primera p�gina hacia atras
                            'osea que le digo que se vaya a la fila 1
                            CaminoError "100CRG" + CStr(tmpUbic3) + "." + CStr(EnQueFilaEstoy)
                            DiscosEnPagina = CargarDiscos(tmpUbic3, False, EnQueFilaEstoy)
                        End If
                    Else
                        'tratarlo normalmente como el 46
                        GoTo Mod46IZQ
                    End If
                Else
Mod46IZQ:
                    CaminoError "000-0082"
                    nDiscoGral = nDiscoGral - 1
                    CaminoError "000-0083"
                    UnSelDisco nDiscoSEL
                    CaminoError "000-0084"
                    SelDisco nDiscoSEL - 1
                End If
            End If
            CaminoError "000-0085"
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
FinTeclaZ:
            CaminoError "000-0086"
            TECLAS_PRES = TECLAS_PRES + "1"
            CaminoError "000-0087"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0088"
            lblTECLAS = TECLAS_PRES
            
        Case TeclaDER
            'esta tecla es DER en el modo 46 pasandpo de abajo a arriba
            'y siguiendo a la atras �? sig en el modo 5
            CaminoError "000-0089"
            If ModoVideoSelTema Then
                CaminoError "000-0090"
                If TemaElegidoModoVideo < UBound(MATRIZ_TEMAS) Then
                    CaminoError "000-0091"
                    UnSelTema TemaElegidoModoVideo
                    CaminoError "000-0092"
                    TemaElegidoModoVideo = TemaElegidoModoVideo + 1
                    CaminoError "000-0093"
                    SelTema TemaElegidoModoVideo
                    CaminoError "000-0094"
                    OrdenarListaTemaVideo
                End If
            Else
                'esta eligiendo discos ya sea en las portadas o en el modo video!!
                CaminoError "000-0095"
                If nDiscoSEL = DiscosEnPagina - 1 Then
                    'ver si hay que pasar hojas (segun config)
                    If PasarHoja Then
                        CaminoError "000-0096"
                        'ver que no se vaya a la mierda!!
                        If nDiscoGral + 1 < TOTAL_DISCOS Then
                            CaminoError "000-0097"
                            'si esta en el modtec 46 pasa al primero
                            'pero si esta en el modo 5 pasa a su mismo nivel
                            'vertical en la hoja que sigue
                            If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                                CaminoError "109CRG" + CStr(nDiscoGral + 1) + ".1"
                                'va a la primera fila!!
                                DiscosEnPagina = CargarDiscos(nDiscoGral + 1, True, 1)
                            End If
                            'busca solo la fila!!
                            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                                CaminoError "102CRG" + CStr(nDiscoGral + 1) + "." + CStr(EnQueFilaEstoy)
                                DiscosEnPagina = CargarDiscos(nDiscoGral + 1, True, EnQueFilaEstoy)
                            End If
                        Else
                            'es el ultimo disco y debe empezar de cero!!!
                            If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                                'es el ultimo disco y debe empezar de cero!!!
                                CaminoError "108CRG0.1"
                                DiscosEnPagina = CargarDiscos(0, True, 1)
                            End If
                            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                                CaminoError "103CRG0." + CStr(EnQueFilaEstoy)
                                DiscosEnPagina = CargarDiscos(0, True, EnQueFilaEstoy)
                                'va a la primera fila!!
                            End If
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
                        CaminoError "000-0098"
                        nDiscoGral = nDiscoGral - DiscosEnPagina + 1
                        CaminoError "000-0099"
                        UnSelDisco nDiscoSEL
                        CaminoError "000-0100"
                        SelDisco 0
                    End If
                Else
                    'ver si llego al final de una linea horizontal para pasar a la hoja
                    'que sigue si esta en el modTeclado5
                    
                    CaminoError "000-0101"
                    'ver si el disco existe !!! o llegamos al final de todo !!!!
                    If nDiscoGral + 1 < TOTAL_DISCOS Then
                        'si esta en el modo 5 me fijo si esta al final de una l�nea
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            'ver ahora si es el �ltimo de una l�nea!!!
                            If nDiscoSEL = (TapasMostradasH * EnQueFilaEstoy) - 1 Then
                                'el disco a iniciar ya no es nDiscoGral + 1  !!!!!!
                                Dim DiscoToIni2 As Long
                                'el primero de esta mas el total de esta!
                                DiscoToIni2 = nDiscoGral - nDiscoSEL + (TapasMostradasH * TapasMostradasV)
                                'ver que no se vaya a la mierda!!
                                If DiscoToIni2 < TOTAL_DISCOS Then
                                    CaminoError "104CRG" + CStr(DiscoToIni2) + "." + CStr(EnQueFilaEstoy)
                                    DiscosEnPagina = CargarDiscos(DiscoToIni2, True, EnQueFilaEstoy)
                                Else
                                    'se termino, ir a la pag1!!
                                    DiscoToIni2 = 0
                                    CaminoError "105CRG" + CStr(DiscoToIni2) + "." + CStr(EnQueFilaEstoy)
                                    DiscosEnPagina = CargarDiscos(DiscoToIni2, True, EnQueFilaEstoy)
                                End If
                            Else
                                'tratarlo como el modo 46
                                GoTo Mod46
                            End If
                        Else
Mod46:
                            CaminoError "000-0102"
                            nDiscoGral = nDiscoGral + 1
                            CaminoError "000-0103"
                            UnSelDisco nDiscoSEL
                            CaminoError "000-0104"
                            SelDisco nDiscoSEL + 1
                        End If
                    End If
                End If
            End If
            CaminoError "000-0105"
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            CaminoError "000-0106"
            TECLAS_PRES = TECLAS_PRES + "2"
            CaminoError "000-0107"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0108"
            lblTECLAS = TECLAS_PRES
        Case TeclaOK
            CaminoError "000-0109"
            TECLAS_PRES = TECLAS_PRES + "3"
            CaminoError "000-0110"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0111"
            lblTECLAS = TECLAS_PRES
            'si estoy en video
            'saber si estoy eligiendo tema. Si no estoy en disco
            
            If ModoVideoSelTema Then
                'si esta en fullscreen NO EJECUTAR!!!
                'solo si no sale por la segunda salida!!!
                If EsVideo And vidFullScreen And Salida2 = False Then GoTo FinKD 'fin keydown
                'si no dice salir cargar tema
                CaminoError "000-0112"
                If T(TemaElegidoModoVideo) = "SALIR" Or T(TemaElegidoModoVideo) = "No hay temas" Then
                    'volver a elegir discos
                    CaminoError "000-0113"
                    frTEMAS.Visible = False
                    CaminoError "000-0114"
                    lblTEMAS.Visible = False
                    CaminoError "000-0115"
                    frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                    CaminoError "000-0116"
                    UnSelTema 0
                    CaminoError "000-0117"
                    ModoVideoSelTema = False
                Else
                    'ejecutar el tema
                    CaminoError "000-0118"
                    
                    'ANTES DE VER CUANTOS CREDITOS NECESITA TENGO QUE SABER SI QUIERE EJECUTAR
                    'MP3 O VIDEO!!!!!!
                    CaminoError "000-0126"
                    Dim temaElegido As String
                    'lstext es una lista oculta  con datos completos
                    temaElegido = txtInLista(MATRIZ_TEMAS(TemaElegidoModoVideo), 0, ",")
                    
                    If LCase(Right(temaElegido, 3)) = "mp3" Or LCase(Right(temaElegido, 3)) = "wma" Then
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
                        CaminoError "000-0119"
                        'restar lo que corresponde!!!
                        If PideVideo Then
                            CREDITOS = CREDITOS - CreditosCuestaTemaVIDEO
                        Else
                            CREDITOS = CREDITOS - CreditosCuestaTema
                        End If
                        'siempre que se ejecute un credito estaremos por debajo de maximo
                        CaminoError "000-0120"
                        OnOffCAPS vbKeyScrollLock, True
                        'grabar cant de creditos
                        CaminoError "000-0121"
                        EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                        CaminoError "000-0122"
                        
                        ShowCredits
                        
                        CaminoError "000-0124"
                        'grabar credito para validar
                        'creditosValidar ya se cargo en load de frmindex
                        CreditosValidar = CreditosValidar + TemasPorCredito
                        CaminoError "000-0125"
                        EscribirArch1Linea SYSfolder + "radilav.cfg", CStr(CreditosValidar)
                        
                        CaminoError "000-0127"
                        'si esta ejecutando pasa a la lista de reproducci�n
                        If MP3.IsPlaying Then
                            'pasar a la lista de reproducci�n
                            Dim NewIndLista As Long
                            CaminoError "000-0128"
                            NewIndLista = UBound(MATRIZ_LISTA)
                            CaminoError "000-0129"
                            ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                            CaminoError "000-0130"
                            'se graba en Matriz_Listas como path, nombre(sin .mp3)
                            MATRIZ_LISTA(NewIndLista + 1) = _
                                temaElegido + "," + _
                                FSO.GetBaseName(T(TemaElegidoModoVideo)) + _
                                " / " + FSO.GetBaseName(UbicDiscoActual)
                            CaminoError "000-0131"
                            CargarProximosTemas
                            'graba en reini.tbr los datos que correspondan por si se corta la luz
                            CaminoError "000-0132"
                            CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                            'volver a elegir discos
                            CaminoError "000-0133"
                            frTEMAS.Visible = False
                            CaminoError "000-0134"
                            lblTEMAS.Visible = False
                            CaminoError "000-0135"
                            frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                            CaminoError "000-0136"
                            UnSelTema 0
                            CaminoError "000-0137"
                            ModoVideoSelTema = False
                        Else
                            'NUNCA ENTRARA AQUI, siempre esta rep video
                            'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                            'paciencia
                            CaminoError "000-0138"
                            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                            CaminoError "000-0139"
                            EjecutarTema temaElegido, True
                        End If
                        
                        VerSiTocaPUB
                        
                    End If
                End If
            Else
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                CaminoError "000-0140"
                'ver si es video deber�a desplegar los temas del disco elegido
                'en modo de texto
                'pero si estoy viendo el video en salida2 es video sera verdadero
                'pero de todas formas no veo als lista de texto y sigo igual
                'solo si esvideo y necesito el modo texto del video!!!!
                If EsVideo And Salida2 = False Then
                    frModoVideo.Height = frDISCOS.Height / 4
                    CaminoError "000-0141"
                    OrdenarListaModoVideo
                    CaminoError "000-0142"
                    lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
                    CaminoError "000-0143"
                    lblTEMAS.Left = lblModoVideo.Left
                    CaminoError "000-0144"
                    frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
                    CaminoError "000-0145"
                    frTEMAS.Height = frDISCOS.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
                    CaminoError "000-0146"
                    lblTEMAS.Visible = True
                    CaminoError "000-0147"
                    frTEMAS.Visible = True
                    CaminoError "000-0148"
                    'cargar los temas multimedia en t()
                    ReDim MATRIZ_TEMAS(0) 'matriz en blanco
                    'es una matriz global
                    CaminoError "000-0149"
                    'en la 6.3 era nDiscoGral+1!!!
                    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
                    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
                    CaminoError "000-0150"
                    ReDim Preserve MATRIZ_TEMAS(0)
                    CaminoError "000-0151"
                    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
                    CaminoError "000-0152"
                    If UBound(MATRIZ_TEMAS) = 0 Then
                        CaminoError "000-0153"
                        T(0) = "No hay temas"
                        CaminoError "000-0154"
                        SelTema 0
                        CaminoError "000-0155"
                        ModoVideoSelTema = True
                        CaminoError "000-0156"
                        Exit Sub
                    End If
                    CaminoError "000-0157"
                    T(0) = "SALIR"
                    '----------------------------
                    'a daniel cruz le da un error como si se volviera a cargar algo que esta cargado
                    'por lo tanto tengo que poner un manejador de error aqui, unico lugar en que se carga esto
                    CaminoError "000-0158"
                    For Each LLL In frmIndex.T
                        CaminoError "000-0159"
                        If LLL.Index > 0 Then Unload LLL
                    Next
                    '----------------------------
                    CaminoError "000-0160"
                    For AA = 1 To UBound(MATRIZ_TEMAS)
                        CaminoError "000-0161"
                        Load T(AA)
                        CaminoError "000-0162"
                        T(AA) = FSO.GetBaseName(txtInLista(MATRIZ_TEMAS(AA), 1, ","))
                        CaminoError "000-0163"
                        T(AA).Top = T(AA - 1).Top + T(AA - 1).Height
                        CaminoError "000-0164"
                        T(AA).Left = T(AA - 1).Left
                        CaminoError "000-0165"
                        T(AA).Visible = True
                        CaminoError "000-0166"
                    Next
                    TemaElegidoModoVideo = 0
                    CaminoError "000-0167"
                    SelTema 0
                    CaminoError "000-0168"
                    ModoVideoSelTema = True
                Else
                    CaminoError "000-0169"
                    If lblDisco(nDiscoSEL) = "01- Los mas escuchados" Then GoTo TOP10Show
                    CaminoError "000-0170"
                    frmTemasDeDisco.Show 1
                End If
            End If
        Case TeclaCerrarSistema
            CaminoError "000-0171"
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            CaminoError "000-0172"
            MostrarCursor True
            CaminoError "000-0173"
            MP3.DoClose
            CaminoError "000-0174"
            If ApagarAlCierre Then APAGAR_PC
            CaminoError "000-0175"
            Unload Me
            End
        Case TeclaESC
            CaminoError "000-0176"
            TECLAS_PRES = TECLAS_PRES + "4"
            CaminoError "000-0177"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            CaminoError "000-0178"
            lblTECLAS = TECLAS_PRES
            CaminoError "000-0179"
            If ModoVideoSelTema Then
                CaminoError "000-0180"
                'volver a elegir discos
                frTEMAS.Visible = False
                CaminoError "000-0181"
                lblTEMAS.Visible = False
                CaminoError "000-0182"
                frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                CaminoError "000-0183"
                UnSelTema 0
                CaminoError "000-0184"
                ModoVideoSelTema = False
            End If
    End Select
FinKD:
    CaminoError "000-0185"
    VerClaves TECLAS_PRES
    CaminoError "000-0186"
    SecSinTecla = 0
    CaminoError "000-0187"
    lblNoTecla = 0
    Exit Sub
TOP10Show:
    CaminoError "000-0188"
    FRMTOP10.Show 1
    
    Exit Sub
    
FallaKD:
    WriteTBRLog "LINEA: " + LineaError + vbCrLf + Err.Description + " N�: " + Str(Err.Number), True
    Resume Next

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Local Error GoTo FallaKD
    
    CaminoError "000-0189"
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    
    If IsKeyPad(Me) Then
        'lasa falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        If KeyCode = 13 Then RealKeyCode = 108
        'ademas si esta apretado el BLOQ NUM
    Else
        'de manera predeterminada son el mismo
        'salvo los casos que se especifican
        RealKeyCode = KeyCode
    End If
    
      
    If RealKeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        CaminoError "000-0190"
        If CREDITOS <= MaximoFichas Then
            CaminoError "000-0191"
            OnOffCAPS vbKeyScrollLock, True
            CaminoError "000-0192"
            CREDITOS = CREDITOS + TemasPorCredito
            CaminoError "000-0193"
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            CaminoError "000-0194"
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            CaminoError "000-0195"
            
            ShowCredits
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CaminoError "000-0198"
            CreditosValidar = CreditosValidar + TemasPorCredito
            CaminoError "000-0199"
            EscribirArch1Linea SYSfolder + "radilav.cfg", CStr(CreditosValidar)
        Else
            'apagar el fichero electronico
            CaminoError "000-0200"
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
    
Exit Sub
    
FallaKD:
    WriteTBRLog "LINEA: " + LineaError + vbCrLf + Err.Description + " N�: " + Str(Err.Number), True
    Resume Next

End Sub

Private Sub Form_Load()
    
    'imagenes no cargadas, vewr si hay algo configurado para el fondo
    Dim ImgFondo As String
    ImgFondo = Trim(LeerConfig("ImgFondo", "NO"))
    If ImgFondo = "NO" Then
        picFondoDisco.Picture = LoadPicture(SYSfolder + "f53.dlw")
    Else
        If FSO.FileExists(ImgFondo) Then
            picFondoDisco.Picture = LoadPicture(ImgFondo)
        Else
            picFondoDisco.Picture = LoadPicture(SYSfolder + "f53.dlw")
        End If
    End If
    'imagen detras de los indicadores
    Dim ImgFondo2 As String
    ImgFondo2 = Trim(LeerConfig("ImgFondo2", "NO"))
    If ImgFondo2 = "NO" Then
        picFondo.Picture = LoadPicture(SYSfolder + "f55.dlw")
        VU21.Picture SYSfolder + "f55.dlw"
    Else
        If FSO.FileExists(ImgFondo) Then
            picFondoDisco.Picture = LoadPicture(ImgFondo)
        Else
            picFondo.Picture = LoadPicture(SYSfolder + "f55.dlw")
            VU21.Picture SYSfolder + "f55.dlw"
        End If
    End If
        
    tbrPassImg1.Picture SYSfolder + "f61.dlw"
    TapaEjecutando.Picture = LoadPicture(SYSfolder + "f61.dlw")
    'la imagen chiquita del exclusivo es la misma!!
    Image1.Picture = LoadPicture(SYSfolder + "f61.dlw")
    
    
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
    
    CaminoError "000-0201"
    RegistroDiario 'anota la fecha, hora y numero del contador
    '--------
    CaminoError "000-0202"
    If K.LICENCIA = HSuperLicencia Then
        CaminoError "000-0203"
        If FSO.FileExists(WINfolder + "SL\indexchi.tbr") Then
            tbrPassImg1.Picture WINfolder + "SL\indexchi.tbr"
            'la imagen chiquita del exclusivo es la misma!!
            Image1.Picture = LoadPicture(WINfolder + "SL\indexchi.tbr")
        End If
    End If
    '--------
    CaminoError "000-0204"
    AjustarFRM Me, 12000
    CaminoError "000-0205"
    If K.LICENCIA = aSinCargar Then
        CaminoError "000-0206"
        lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        CaminoError "000-0207"
        lblDEMO = textoUsuario
    End If
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    
    CaminoError "000-0208"
    '-----------------
    If K.LICENCIA = HSuperLicencia Then
        CaminoError "000-0209"
        If FSO.FileExists(WINfolder + "SL\txtIDX.tbr") Then
            CaminoError "000-0210"
            Set TE = FSO.OpenTextFile(WINfolder + "SL\txtIDX.tbr", ForReading, False)
            CaminoError "000-0211"
            Dim NewT As String
            CaminoError "000-0212"
            NewT = TE.ReadAll
            CaminoError "000-0213"
            lblTBR = NewT
            CaminoError "000-0214"
            TE.Close
        Else
            CaminoError "000-0215"
            lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
        End If
    Else
        CaminoError "000-0216"
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
    End If
    '-----------------
    CaminoError "000-0217"
    VU1.Width = Screen.Width
    CaminoError "000-0218"
    VU1.Left = 0: VU1.Top = 0
    CaminoError "000-0219"
    VU1.Height = picFondo.Top - 25
    CaminoError "000-0220"
    
    'si es exclusivo inhabilito el vumetro GRANDE !!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        'que entre en el control
        CaminoError "000-0221"
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width
        CaminoError "000-0222"
        frDISCOS.Left = VU1.AnchoBarra + 25 '0
    Else
        CaminoError "000-0223"
        frDISCOS.Left = 0 ' tapa a las barras que no se usan 'VU1.Left + VU1.Width
        CaminoError "000-0224"
        frDISCOS.Width = VU1.Width ' Screen.Width - VU1.Width
    End If
    CaminoError "000-0225"
    frDISCOS.Top = 0
    CaminoError "000-0226"
    frDISCOS.Height = picFondo.Top
    CaminoError "000-0227"
    picFondoDisco.Height = frDISCOS.Height
    CaminoError "000-0228"
    picFondoDisco.Width = frDISCOS.Width
    
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    CaminoError "000-0231"
    If MostrarTouch = False Then
        CaminoError "000-0232"
        Frame1.Visible = False 'frame del touch
        lblTemaSonando.Width = Screen.Width - lblTemaSonando.Left - 250
        lstProximos.Width = Screen.Width - lstProximos.Left - 250
        lblTBR.Width = Screen.Width - lblTBR.Left - 250
        lblDEMO.Width = Screen.Width - lblDEMO.Left - 250
    End If
    'frDISCOS contiene los discos a mostrar
    'se debera calcualr el tama�o de cada discos asi como cantidad horizontal y vertical
    CaminoError "000-0241"
    
    Dim EspacioEntreDiscosH As Long
    Dim EspacioEntreDiscosV As Long
    Dim AnchoTapaDisco As Long
    Dim AltoTapaDisco As Long
    
    CaminoError "000-0245"
    If DistorcionarTapas Then
        EspacioEntreDiscosV = 0
        EspacioEntreDiscosH = 0
        CaminoError "000-0242b"
        AnchoTapaDisco = (frDISCOS.Width / TapasMostradasH)
        CaminoError "000-0243b"
        AltoTapaDisco = (frDISCOS.Height / TapasMostradasV)
    Else
        'el alto de estos incluye tambien el lbldisco
        CaminoError "000-0242"
        AnchoTapaDisco = (frDISCOS.Width * 0.8 / TapasMostradasH)
        CaminoError "000-0243"
        AltoTapaDisco = (frDISCOS.Height * 0.8 / TapasMostradasV)
        'ver cual es mayor para no permitir mucha distorsion
        'lo que se ajuste se agranda del espacio entrediscos
        CaminoError "000-0244"
        EspacioEntreDiscosV = (frDISCOS.Height * 0.2 / (TapasMostradasV + 1))
        EspacioEntreDiscosH = (frDISCOS.Width * 0.2 / (TapasMostradasH + 1))
    End If
    CaminoError "000-0246"
    
'    If DistorcionarTapas = False Then
'        CaminoError "000-0247"
'        Dim DIFF As Double
'        CaminoError "000-0248"
'        DIFF = AnchoTapaDisco - AltoTapaDisco
'        CaminoError "000-0249"
'        If DIFF > 0 Then
'            CaminoError "000-0250"
'            'el ancho es mas que el alto
'            AnchoTapaDisco = AltoTapaDisco
'            CaminoError "000-0251"
'            'EspacioEntreDiscosH = DIFF
'        Else
'            CaminoError "000-0252"
'            'el alto es mas que el ancho
'            AltoTapaDisco = AnchoTapaDisco
'            CaminoError "000-0253"
'            'EspacioEntreDiscosV = -DIFF
'        End If
'    End If
'
    CaminoError "000-0254"
    If MostrarRotulos Then
        CaminoError "000-0255"
        TapaCD(0).Width = AnchoTapaDisco
        CaminoError "000-0256"
        TapaCD(0).Height = AltoTapaDisco * 0.79 '80%disco, 20% lbldisco
        CaminoError "000-0257"
        lblDisco(0).Height = AltoTapaDisco * 0.19 '80%disco, 20% lbldisco
        CaminoError "000-0258"
        lblDisco(0).Width = AnchoTapaDisco
    Else
        CaminoError "000-0259"
        TapaCD(0).Width = AnchoTapaDisco
        CaminoError "000-0260"
        TapaCD(0).Height = AltoTapaDisco
        CaminoError "000-0261"
        lblDisco(0).Visible = False
    End If
    'centrar!!
    Dim IniCentrarH As Long
    IniCentrarH = EspacioEntreDiscosH
    Dim IniCentrarV As Long
    IniCentrarV = EspacioEntreDiscosV
    CaminoError "000-0262"
    lblDisco(0).Left = IniCentrarH
    CaminoError "000-0268"
    TapaCD(0).Left = IniCentrarH
    'ver si los rotulos van arriba o abajo
    If RotulosArriba Then
        CaminoError "000-0263"
        lblDisco(0).Top = IniCentrarV
        CaminoError "000-0265"
        If MostrarRotulos Then
            CaminoError "000-0266"
            TapaCD(0).Top = lblDisco(0).Top + lblDisco(0).Height + 50
        Else
            CaminoError "000-0267"
            TapaCD(0).Top = IniCentrarV
        End If
    Else
        CaminoError "000-0269"
        TapaCD(0).Top = IniCentrarV
        CaminoError "000-0271"
        lblDisco(0).Top = TapaCD(0).Top + TapaCD(0).Height + 50
    End If
    CaminoError "000-0272"
    Dim CantDiscos As Long
    CaminoError "000-0273"
    CantDiscos = TapasMostradasH * TapasMostradasV
    'cargar la cantidad de tapas correspondientes
    CaminoError "000-0274"
    c = 0
    CaminoError "000-0275"
    Do While c < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        CaminoError "000-0276"
        c = c + 1
        CaminoError "000-0277"
        Load TapaCD(c)
        CaminoError "000-0278"
        Load lblDisco(c)
        'ya toman el tama�o del original
        CaminoError "000-0279"
        If c / TapasMostradasH = c \ TapasMostradasH Then
            'es una tapa al principio de linea
            CaminoError "000-0280"
            If RotulosArriba Then
                CaminoError "000-0281"
                lblDisco(c).Left = IniCentrarH
                CaminoError "000-0282"
                lblDisco(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                CaminoError "000-0283"
                TapaCD(c).Left = IniCentrarH
                If MostrarRotulos Then
                    CaminoError "000-0284"
                    TapaCD(c).Top = lblDisco(c).Top + lblDisco(c).Height + 50
                Else
                    CaminoError "000-0285"
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + 50
                End If
                CaminoError "000-0286"
                TapaCD(c).Visible = True
                CaminoError "000-0287"
                If MostrarRotulos Then lblDisco(c).Visible = True
            Else
                CaminoError "000-0288"
                TapaCD(c).Left = IniCentrarH
                If MostrarRotulos Then
                    CaminoError "000-0289"
                    TapaCD(c).Top = lblDisco(c - TapasMostradasH).Top + lblDisco(c - TapasMostradasH).Height + EspacioEntreDiscosV
                Else
                    CaminoError "000-0290"
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                End If
                CaminoError "000-0291"
                lblDisco(c).Left = IniCentrarH
                CaminoError "000-0292"
                lblDisco(c).Top = TapaCD(c).Top + TapaCD(c).Height + 50
                CaminoError "000-0293"
                TapaCD(c).Visible = True
                CaminoError "000-0294"
                If MostrarRotulos Then lblDisco(c).Visible = True
            End If
        Else
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                CaminoError "000-0295"
                lblDisco(c).Left = lblDisco(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                CaminoError "000-0296"
                lblDisco(c).Top = lblDisco(c - 1).Top
                CaminoError "000-0297"
                TapaCD(c).Left = lblDisco(c).Left
                CaminoError "000-0298"
                TapaCD(c).Top = TapaCD(c - 1).Top
                CaminoError "000-0299"
                TapaCD(c).Visible = True
            Else
                CaminoError "000-0300"
                TapaCD(c).Left = TapaCD(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                CaminoError "000-0301"
                TapaCD(c).Top = TapaCD(c - 1).Top
                CaminoError "000-0302"
                lblDisco(c).Left = TapaCD(c).Left
                CaminoError "000-0303"
                lblDisco(c).Top = lblDisco(c - 1).Top
                CaminoError "000-0304"
                TapaCD(c).Visible = True
            End If
            CaminoError "000-0305"
            If MostrarRotulos Then lblDisco(c).Visible = True
        End If
        
    Loop
    CaminoError "000-0306"
    OnOffCAPS vbKeyScrollLock, True
    CaminoError "000-0307"
    lblV = "versi�n " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    CaminoError "000-0308"
    lblTiempoRestante = "Falta: " + "00:00"
    'ocultar las etiquetas
    CaminoError "000-0309"
    Me.AutoRedraw = AutoReDibuj
    CaminoError "000-0310"
    Me.Left = Screen.Width / 2 - Me.Width / 2
    CaminoError "000-0311"
    Me.Top = Screen.Height / 2 - Me.Height / 2
    'ver cuantos creditos hay
    CaminoError "000-0312"
    CREDITOS = Val(LeerArch1Linea(AP + "creditos.tbr"))
    CaminoError "000-0313"
    
    ShowCredits
    
    'dejar cargado el mostrados de procesos
    'Load frmini
    'cargar las variables globales
    CaminoError "000-0316"
    TEMA_REPRODUCIENDO = "Sin reproducci�n actual"
    CaminoError "000-0317"
    TEMA_SIGUIENTE = "No hay proximo tema"
    CaminoError "000-0318"
    TEMAS_EN_LISTA = 0
    
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
    CaminoError "000-0319"
    
    'usar el que lee los discos con matrices temporales y _
    sumar todas esas matrics a Matriz_Discos _
    fijarse que el orden no sea alfabetico, solo alfabetico _
    dentro de cada origen de discos
    
    'obtenerDir ya los Ordena JOIA!
    
    
    Dim MtxTmpOrigenes() As String
    Dim Origenes As String
    Origenes = LeerArch1Linea(SYSfolder + "oddtb.jut")
    
    Dim PartOrigenes() As String
    PartOrigenes = Split(Origenes, "*")
    
    Dim AAA As Long
    For AAA = 0 To UBound(PartOrigenes)
        'ver los discos del origene elegido
        MtxTmpOrigenes() = ObtenerDir(PartOrigenes(AAA))
        'acumular a la matriz general
        SumarMatriz MATRIZ_DISCOS, MtxTmpOrigenes
    Next AAA
    
    'ya se sumop y esta listo para cargarse ordenados los discos dentro de cada origen
    MostrarDiscosMTX
    'MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
        
    CaminoError "000-0320"
    Dim CarpActual As String
    Dim pathTema As String, DuracionTema As String, nombreTEMA As String
    'mostrar proceso
    CaminoError "000-0321"
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    
    'ret devuelve la cantidadd de discos cargados
    CaminoError "000-0322"
    DiscosEnPagina = CargarDiscos(0, True, 1)
    'inicializar la matriz_lista (lista de reproduccion
    CaminoError "000-0323"
    ReDim MATRIZ_LISTA(0)
    CaminoError "000-0324"
    lblTOTdiscos = "Discos: " + Trim(Str(UBound(MATRIZ_DISCOS)))
    
    'si quedaron temas pendientes cargarlos
    CaminoError "000-0325"
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            CaminoError "000-0326"
            If FSO.FileExists(AP + "reini.tbr") Then
                CaminoError "000-0327"
                Set TE = FSO.OpenTextFile(AP + "reini.tbr", ForReading, False)
                Dim TT As String 'cada tema
                Dim z As Integer 'contador de temas en lista anterior
                z = 1
                CaminoError "000-0328"
                Do While Not TE.AtEndOfStream
                    CaminoError "000-0329"
                    TT = TE.ReadLine
                    CaminoError "000-0330"
                    ReDim Preserve MATRIZ_LISTA(z)
                    CaminoError "000-0331"
                    MATRIZ_LISTA(z) = TT
                    CaminoError "000-0332"
                    z = z + 1
                Loop
                CaminoError "000-0333"
                TE.Close
            End If
            CaminoError "000-0334"
            EMPEZAR_SIGUIENTE
        Case "NADA"
            'no hacer nada
            'borrar la lista
            CaminoError "000-0335"
            'borrra los temas 'y los creditos?
            If FSO.FileExists(AP + "reini.tbr") Then FSO.DeleteFile AP + "reini.tbr", True
            
            CaminoError "000-0336"
            Timer1.Interval = 10000
    End Select
    CaminoError "000-0337"
    Unload frmINI
    
    'ver si hay validacion por creditos
    CaminoError "000-0338"
    Validar = LeerConfig("Validar", "0")
    If Validar Then
        'ver si existe el archivo Creditos Validar
        CaminoError "000-0339"
        If FSO.FileExists(SYSfolder + "radilav.cfg") Then
            'leer el archivo de creditos vaildados
            CaminoError "000-0340"
            CreditosValidar = CLng(LeerArch1Linea(SYSfolder + "radilav.cfg"))
            'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Else
            CaminoError "000-0341"
            EscribirArch1Linea SYSfolder + "radilav.cfg", "0"
            CaminoError "000-0342"
            CreditosValidar = 0
            CaminoError "000-0343"
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el m�ximo y si hay que avisar
        CaminoError "000-0344"
        ValidarCada = LeerConfig("ValidarCada", "500")
        CaminoError "000-0345"
        AvisarAntes = LeerConfig("AvisarAntes", "50")
        CaminoError "000-0346"
        If CreditosValidar > ValidarCada - AvisarAntes Then
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            
            'uso el frmClave que tiene la variable publica ClaveIngresada
            CaminoError "000-0347"
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = ClaveParaValidar(CodigoParaClaveActual)
            CaminoError "000-0348"
            Dim QuedanC As Long
            QuedanC = ValidarCada - CreditosValidar
            If QuedanC > 0 Then
                'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
                CaminoError "000-0349"
                MsgBox "Ingrese a continuacion su clave para continuar utilizando 3PM. " + vbCrLf + _
                    "Debe enviar la administrador el codigo: " + vbCrLf + _
                    CodigoParaClaveActual + vbCrLf + _
                    "Puede todavia omitir esta clave. Solo le quedan " + CStr(QuedanC) + " creditos hasta que 3PM se inhabilite"
            Else
                CaminoError "000-0350"
                MsgBox "De no ingresar la clave correspondiente 3PM no podra continuar. Ha llegado al limite de creditos posibles"
            End If
            CaminoError "000-0351"
            frmCLAVE.Show 1
            CaminoError "000-0352"
            If UCase(ClaveIngresada) <> UCase(ClaveCorrespondiente) Then
                CaminoError "000-0353"
                If QuedanC > 0 Then
                    CaminoError "000-0354"
                    MsgBox "La clave es erronea!" + vbCrLf + _
                        "Le quedan " + CStr(QuedanC) + " creditos por cargar antes que se inhabilite 3PM"
                Else
                    CaminoError "000-0355"
                    MsgBox "No podra seguir utilizando 3PM hasta que valide con la clave correspondiente"
                    End
                End If
            Else
                CaminoError "000-0356"
                'todo OK. Cargo bien la clave
                CreditosValidar = 0
                CaminoError "000-0357"
                EscribirArch1Linea SYSfolder + "radilav.cfg", "0"
                'empezar un nuevo periodo
                CaminoError "000-0358"
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        End If
        CaminoError "000-0359"
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
    
    CaminoError "000-0360"
    lblSel.Visible = False
    CaminoError "000-0361"
    lblDisco(nDisco).ForeColor = vbBlack
    CaminoError "000-0362"
    'lblDISCO(nDisco).Font.Bold = True
    lblDisco(nDisco).Font.Underline = True
    CaminoError "000-0363"
    lblDisco(nDisco).BackColor = vbYellow
    CaminoError "000-0364"
    nDiscoSEL = nDisco
    CaminoError "000-0365"
    lblSel.Top = TapaCD(nDiscoSEL).Top - lblSel.BorderWidth * 10
    CaminoError "000-0366"
    lblSel.Left = TapaCD(nDiscoSEL).Left - lblSel.BorderWidth * 10
    CaminoError "000-0367"
    lblSel.Height = TapaCD(nDiscoSEL).Height + lblSel.BorderWidth * 20
    CaminoError "000-0368"
    lblSel.Width = TapaCD(nDiscoSEL).Width + lblSel.BorderWidth * 20
    CaminoError "000-0369"
    lblSel.Visible = True
    CaminoError "000-0370"
    lblSel.ZOrder
    CaminoError "000-0371"
    lblDisco(nDisco).ZOrder
    
    'seleccionar de la lista de solo video
    CaminoError "000-0372:" + CStr(nDisco) + ":" + CStr(nDiscoGral)
    L(nDiscoGral).ForeColor = vbWhite
    CaminoError "000-0373"
    L(nDiscoGral).BackColor = vbBlack
    CaminoError "000-0374"
    LastDiscoSel = nDiscoGral 'para saber cual desactivar en unsel
    CaminoError "000-0375"
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
    CaminoError "000-0376"
    If EsVideo Then OrdenarListaModoVideo
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    CaminoError "000-0377"
    lblDisco(nDisco).ForeColor = vbWhite
    'lblDISCO(nDisco).Font.Bold = False
    CaminoError "000-0378"
    lblDisco(nDisco).Font.Underline = False
    CaminoError "000-0379"
    lblDisco(nDisco).BackColor = vbBlack
    'seleccionar de la lista de solo video
    CaminoError "000-0380"
    L(LastDiscoSel).ForeColor = vbBlack
    CaminoError "000-0381"
    L(LastDiscoSel).BackColor = vbWhite
    CaminoError "000-0382"
    If CargarIMGinicio Then
        TapaCD(LastDiscoSel).BorderStyle = 0
    Else
        TapaCD(nDisco).BorderStyle = 0
    End If
    CaminoError "000-0383"
    If EsVideo Then OrdenarListaModoVideo
End Sub

Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean, DeQueFila As Long) As Long
        
    On Local Error GoTo NoCRG
    
    'indicando en que disco se inicia carga ese y los seis (o lo que corresponde) _
        que le sigen
    'DeQueFial dice si es primero o �ltimo de cual fila!!!
    'devuelve el n�mero de discos cargados
    CaminoError "000-0384"
    Dim mCargarDiscos As Long
    mCargarDiscos = 0
    CaminoError "000-0385"
    Dim TotPags As Long
    TotPags = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
    CaminoError "000-0386"
    lblPag = "Pagina " + CStr(Round(numDiscoIniciar / (TapasMostradasH * TapasMostradasV) + 1, 0)) + " de " + CStr(TotPags + 1)
    'tomar el disco que va a quedar seleccionado
    'como numero de disco en el indice general
    If SelPrimero Then
        CaminoError "000-0387"
        'si la fila es uno (la primera) entonces el calculo es facil
        nDiscoGral = numDiscoIniciar + ((DeQueFila - 1) * TapasMostradasH)
    Else
        CaminoError "000-0388"
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            nDiscoGral = numDiscoIniciar + ((TapasMostradasH * TapasMostradasV) - 1)
        End If
        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
            nDiscoGral = numDiscoIniciar + ((TapasMostradasH * DeQueFila) - 1)
        End If
                
        'si no va a seleccionar el primero es el ultimo
        'y si no hay p�gina completa!!!!!!!!!!
        If nDiscoGral >= TOTAL_DISCOS Then nDiscoGral = TOTAL_DISCOS - 1
        
    End If
    'esconder todos los discos
    Dim NDR As Long 'numero de tapa de disco real del 0 al 5 (total de discos-1)
    
    'no hacer esto al pedo si ya estan cargadas
    Dim NDI As Long '=numdiscoiniciar de la pagina
    Dim c As Integer
    c = 1
    CaminoError "000-0389"
    NDI = numDiscoIniciar
    CaminoError "000-0390"
    If CargarIMGinicio Then
        If SelPrimero Then
            'si voy para adelante ocultar los que ya pase
            c = 1
            CaminoError "000-0391"
            Do While c <= (TapasMostradasH * TapasMostradasV)
                CaminoError "000-0392"
                'si no es la primera hoja!!
                If NDI >= (TapasMostradasH * TapasMostradasV) Then
                    CaminoError "000-0393"
                    TapaCD(NDI - c).Visible = False
                    'no se cargan lbldisco, usan solo del 0 al 5
                    CaminoError "000-0394"
                    lblDisco(c - 1).Visible = False
                End If
                c = c + 1
            Loop
            CaminoError "000-0395"
            Me.Refresh
        Else
            'sino ocultar los de adelante
            c = 1
            CaminoError "000-0396"
            Do While c <= (TapasMostradasH * TapasMostradasV)
                CaminoError "000-0397"
                'ocultar solo si estaba visible (si existe)
                Dim DiscoAOcultar As Long
                'InicioDePag + DiscosEnPag + Contador
                DiscoAOcultar = NDI + ((TapasMostradasH * TapasMostradasV) - 1) + c
                If DiscoAOcultar < UBound(MATRIZ_DISCOS) Then
                    TapaCD(DiscoAOcultar).Visible = False
                End If
                'ADEMAS VER SI ESTOY LLENDO DESDE LA PRIMERA PAGINA_
                    'HACIA ATRAS!!!!
                Dim UltimoDeEstaPagina As Long
                UltimoDeEstaPagina = NDI + (TapasMostradasH * TapasMostradasV) - c
                If UltimoDeEstaPagina > UBound(MATRIZ_DISCOS) Then
                    'si entra aca es por que la pagina elegida es la ultima
                    'y vengo volviendo desde la primera
                    'lo que hay que ocultar entonces son los discos de
                    'la primera p�gina!
                    Dim DiscoPag1Borrar As Long
                    DiscoPag1Borrar = (TapasMostradasH * TapasMostradasV) - c
                    TapaCD(DiscoPag1Borrar).Visible = False
                End If
                CaminoError "000-0398"
                lblDisco(c - 1).Visible = False
                c = c + 1
            Loop
            'Me.Refresh
        End If
    Else
        'si no se cargaron al inicio!!
        CaminoError "000-0399"
        Do While NDR < ((TapasMostradasH * TapasMostradasV))
            CaminoError "000-0400"
            TapaCD(NDR).Visible = False
            CaminoError "000-0401"
            lblDisco(NDR).Visible = False
            CaminoError "000-0402"
            NDR = NDR + 1
        Loop
        Dim ArchTapa As String
    End If
    NDR = 0
    CaminoError "000-0403"
    
    Do While NDI < numDiscoIniciar + ((TapasMostradasH * TapasMostradasV))
        'ver si existe si hay disco con este n�
        CaminoError "000-0404"
        'el = es de la 6.5
        If NDI <= UBound(MATRIZ_DISCOS) Then
            CaminoError "000-0405"
            mCargarDiscos = mCargarDiscos + 1
            'ver si ya estan cargadas o se deben cargar
            CaminoError "000-0406"
            If CargarIMGinicio Then
                CaminoError "000-0407"
                TapaCD(NDI).Visible = True
                CaminoError "000-0408"
                TapaCD(NDI).ZOrder
            Else
                'ver si hay tapa
                CaminoError "000-0409"
                ArchTapa = txtInLista(MATRIZ_DISCOS(NDI), 0, ",")
                CaminoError "000-0410"
                If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
                CaminoError "000-0411"
                ArchTapa = ArchTapa + "tapa.jpg"
                CaminoError "000-0412"
                If FSO.FileExists(ArchTapa) Then
                    CaminoError "000-0413"
                    TapaCD(NDR).Picture = LoadPicture(ArchTapa)
                Else
                    CaminoError "000-0414"
                    TapaCD(NDR).Picture = LoadPicture(SYSfolder + "f61.dlw")
                End If
                CaminoError "000-0415"
                TapaCD(NDR).Visible = True
            End If
            'poner nombre al disco
            CaminoError "000-0416"
            'antes en la 6.3 era NDI+1 !!
            lblDisco(NDR) = txtInLista(MATRIZ_DISCOS(NDI), 1, ",")
            CaminoError "000-0417"
            If MostrarRotulos Then lblDisco(NDR).Visible = True
        End If
        CaminoError "000-0418"
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    CargarDiscos = mCargarDiscos
    If SelPrimero Then
        CaminoError "000-0419"
        'si es modo 46 no me importa la fila!!!!
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            'y si voy de la ultima pagina incompleta hasta la primera???
            'UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
            'discos en pagina es la cantidad actual en la pagina
            'si es la ultima y esta incompleta debe saber cuantos se cargaron!!!
            
            'Y SI ES LA PRIMEERA VEZ!!!
            'UFFFFFFFFFFFFFFFFFFFF
            If DiscosEnPagina > 0 Then
                UnSelDisco DiscosEnPagina - 1
            Else
                UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
            End If
        Else
            'supone que es de la ultima columna siempre
            'pero en la 6.5 ya puede pasar al inicio de nuevo desde
            'una columna que no sea necesariamnete la ultima
            'si viene de una fila que no es la �ltima!!!!!!
            Dim DesSelModo5 As Long
            'el (TapasMostradasH - 1) inicial supone la ultima columna
            'DesSelModo5 = (TapasMostradasH - 1) + ((DeQueFila - 1) * TapasMostradasH)
            'pero ya no es asi!!!!
            Dim ColumnaSel As Long
            'nDisco-(fila*Tapash)
            ColumnaSel = nDiscoSEL - (nDiscoSEL \ TapasMostradasH) * TapasMostradasH
            DesSelModo5 = ColumnaSel + ((DeQueFila - 1) * TapasMostradasH)
            UnSelDisco DesSelModo5
        End If
        
        CaminoError "000-0420"
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
        CaminoError "000-0421"
        'si es modo 46 no me importa la fila!!!!
        'SI IMPORTA AHORA QUE SE PUEDE VENIR DESDEW LA PRIMEWRA PAGINA HACIA ATRAS!
        'HAY QUE ELEGIR EL ULTMODE LA ULTIA PAGINA!!!!
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            UnSelDisco 0
            'si o si la ultima!?????
            SelDisco mCargarDiscos - 1
        Else
            CaminoError "000-0422"
            'tiene que desseleccionar el que ven�a !!
            UnSelDisco (DeQueFila - 1) * TapasMostradasH
            
            
            
            Dim DiscoSelModo5TT As Long
            DiscoSelModo5TT = ((TapasMostradasH * DeQueFila) - 1)
            'ver si esta volviendo a la ultima p�gina desde la primera!!!
            If DiscoSelModo5TT + numDiscoIniciar >= TOTAL_DISCOS Then
                DiscoSelModo5TT = (TOTAL_DISCOS - 1) - numDiscoIniciar
            End If
            SelDisco DiscoSelModo5TT
        End If
        
        
    End If
    
    Exit Function
    
NoCRG:
    WriteTBRLog Err.Description + " N�: " + Str(Err.Number), True
    Resume Next

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
    
    CaminoError "000-0423"
    MostrarCursor True
    CaminoError "000-0425"
    'MP3.DoStop EL DOsTOP GENERA EL EVENTO ENDPLAY QUE EJECUTA EL QUE SIGUE!!!
    CaminoError "000-0426"
    MP3.DoClose
    CaminoError "000-0427"
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
        TapaEjecutando.Picture = LoadPicture(SYSfolder + "f61.dlw")
    End If
    CaminoError "000-0428"
    TotalTema = MP3.LengthInSec
    CaminoError "000-0429"
    Ancho = lblTemaSonando.Width
    'EVITAR DIVISIONES POR CERO
    CaminoError "000-0430"
    If TotalTema > 0 And MP3.IsPlaying Then
        CaminoError "000-0431"
        Variacion = Ancho / TotalTema
        CaminoError "000-0432"
        lblTiempoRestante = "TOTAL: " + MP3.Falta
    Else
        CaminoError "000-0433"
        lblTiempoRestante = "Falta: " + "00:00"
    End If
    CaminoError "000-0434"
    VolBajando = MP3.Volumen
    
    Prog.Clear
    Prog.MAX = MP3.LengthInSec
    
End Sub

Private Sub MP3_EndPlay()
    EstoyEnModoVideoMiniSelDisco = False
    frmIndex.TapaEjecutando.Picture = LoadPicture(SYSfolder + "f61.dlw")
    'volver a PasarHoja a su estado original3
    CaminoError "000-0435"
    PasarHoja = LeerConfig("PasarHoja", "1")
    CaminoError "000-0436"
    VU1.Width = Screen.Width
    CaminoError "000-0437"
    
    'ver si es fullscreen o no!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If vidFullScreen Then
        'frDISCOS.Height = picFondo.Top
        VU1.Height = picFondo.Top
    End If
    'reacomodo si vengo de video minimo
    'tener el cuenta el exclusivo!!!!!!!!!!!!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        CaminoError "000-0438"
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width - VU1.Width
        CaminoError "000-0439"
        frDISCOS.Left = VU1.AnchoBarra + 25 ' VU1.Width
        'vu no se mueve si termina un video        'VU1.Top = 0        'VU1.Height = Me.Height
    Else
        CaminoError "000-0440"
        frDISCOS.Width = VU1.Width ' Screen.Width
        CaminoError "000-0441"
        frDISCOS.Left = 0
    End If
    
    picFondoDisco.Height = frDISCOS.Height
    picFondoDisco.Width = frDISCOS.Width
    frModoVideo.Visible = False
    lblModoVideo.Visible = False
    frTEMAS.Visible = False
    lblTEMAS.Visible = False
    ModoVideoSelTema = False
    'termino una cancion
    If EsVideo Then MP3.DoClose
    'lo destapo al terminar de acomodar todos los controles en otro lado
    'picVideo.Visible = False
    lblREP.BackStyle = 0
    lblREP.ForeColor = vbWhite
    lblREP = ""
    EMPEZAR_SIGUIENTE
    
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
    CaminoError "000-0455"
    sRest = MP3.FaltaInSec
    CaminoError "000-0456"
    PorcEjecutado = MP3.PercentPlay
    CaminoError "000-0457"
    If PorcEjecutado > PorcentajeTEMA And CORTAR_TEMA Then
        CaminoError "000-0458"
        VolBajando = VolBajando - 5 'baja 1 por segundo
        CaminoError "000-0459"
        lblTemaSonando = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        lblTemaSonando2 = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        CaminoError "000-0460"
        If VolBajando > 0 Then
            CaminoError "000-0461"
            MP3.Volumen = VolBajando
        Else
            CaminoError "000-0462"
            MP3.DoStop
            'EL DOSTOP DESENCADENA UN END PLAY QUE REALIZA UN EMPEZAR SIGUINETE
            'EMPEZAR_SIGUIENTE
        End If
    End If
    CaminoError "000-0463"
    lblTiempoRestante = "Falta: " + MP3.Falta
    Prog.DibujarCirculo CDbl(SecondsPlayed)
    
    CaminoError "000-0464"
    wi = Ancho - Variacion * (SecondsPlayed - 2)
    CaminoError "000-0465"
    
    '=====================================
    CaminoError "000-0466"
    If K.LICENCIA = aSinCargar And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        CaminoError "000-0467"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        CaminoError "000-0468"
        MP3.DoStop
    End If
    'cotar tambin en el gratuito
    CaminoError "000-0469"
    If K.LICENCIA = CGratuita And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        CaminoError "000-0470"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        CaminoError "000-0471"
        MP3.DoStop
    End If
    '=====================================
End Sub

Private Sub TapaCD_Click(Index As Integer)
    'nunca hay que pasar hojas
    'nDiscoGral = nDiscoGral + (Index - nDiscoSEL)
    CaminoError "000-0473"
    nDiscoGral = Index 'si se cargan todas las im�genes al inicio index=nDiscoGral
    CaminoError "000-0474"
    If nDiscoGral + 1 > TOTAL_DISCOS Then
        CaminoError "000-0475"
        MsgBox "No existe el disco elegido!!. " + vbCrLf + _
            "Carge discos desde el ADMINISTRADOR DE DISCOS en la " + vbCrLf + _
            "p�gina de configuracion (presionando la tecla 'C')"
        CaminoError "000-0476"
        Exit Sub
    End If
    CaminoError "000-0477"
    UnSelDisco nDiscoSEL
    CaminoError "000-0478"
    Dim PagNum As Long
    PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
    CaminoError "000-0479"
    nDiscoSEL = Index - (PagNum * (TapasMostradasH * TapasMostradasV))
    CaminoError "000-0480"
    SelDisco nDiscoSEL
    CaminoError "000-0481"
    lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
    CaminoError "000-0482"
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
    CaminoError "000-0483"
    If MP3.IsPlaying Then Exit Sub
    'controla el tiempo sin uso (sin ejecucion de temas)
    CaminoError "000-0484"
    SecSinUso = SecSinUso + (Timer1.Interval / 1000)
    CaminoError "000-0485"
    lblNoUSO = Trim(Str(SecSinUso))
    CaminoError "000-0486"
    If SecSinUso >= EsperaMinutos Then 'esperaminutos esta en segundos
        CaminoError "000-0487"
        SecSinUso = 0
        CaminoError "000-0488"
        Dim TemasDisponibles As Long
        If TemasEnRank(1) > 50 Then
            CaminoError "000-0489"
            TemasDisponibles = TemasEnRank(1) 'todos los que se escucharon
        Else
            CaminoError "000-0490"
            TemasDisponibles = TemasEnRank(0) 'todos los que se escucharon
        End If
        CaminoError "000-0491"
        Randomize Timer
        CaminoError "000-0492"
        z = Int(Rnd * TemasDisponibles)
        z = z + 1
        CC = 0
        CaminoError "000-0493"
        If FSO.FileExists(AP + "ranking.tbr") = False Then
            CaminoError "000-0494"
            FSO.CreateTextFile AP + "ranking.tbr", True
            'me voy al azar ya que no hay para elegirdel rank
            CaminoError "000-0495"
            GoTo AZAR
        End If
        CaminoError "000-0496"
        Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
        CaminoError "000-0497"
        Dim TT As String
        'antes de entra ver si el archivo no tiene nada
        CaminoError "000-0498"
        If TE.AtEndOfStream Then GoTo AZAR
        CaminoError "000-0499"
        Do While Not TE.AtEndOfStream
            CaminoError "000-0500"
            CC = CC + 1
            CaminoError "000-0501"
            TT = TE.ReadLine
            CaminoError "000-0502"
            If CC = z Then
                CaminoError "000-0503"
                Dim TemaAzar As String
                CaminoError "000-0504"
                TemaAzar = txtInLista(TT, 1, ",")
                'si tuve los discos cargados en una unidad o una ubicaci�n distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                CaminoError "000-0505"
                If FSO.FileExists(TemaAzar) Then
                    CaminoError "000-0506"
                    CORTAR_TEMA = True 'este tema se eligio al azar no va entero
                    CaminoError "000-0507"
                    SecSinUso = 0
                    CaminoError "000-0508"
                    TE.Close
                    CaminoError "000-0509"
                    EjecutarTema TemaAzar, False
                    CaminoError "000-0510"
                    Exit Sub
                Else
AZAR:
                    'ejecutar algun tema de cualquier disco
                    CaminoError "000-0511"
                    Dim MTX10() As String: zz = 0
                    CaminoError "000-0512"
                    ruta = AP + "discos\"
                    CaminoError "000-0513"
                    Dim NombreDir As String
                    CaminoError "000-0514"
                    NombreDir = Dir$(ruta & "*.*", vbDirectory)
                    CaminoError "000-0515"
                    Do While Len(NombreDir)
                        CaminoError "000-0516"
                        If NombreDir = "." Or NombreDir = ".." Then
                            ' excluir las entradas "." y ".."
                            CaminoError "000-0517"
                        ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                            ' este es un archivo normal
                            CaminoError "000-0518"
                        Else
                            CaminoError "000-0519"
                            'ver los primeros diez discos. En alguno tiene que haber temas
                            'yo se que el primero no tiene temas por que es
                            '01 - los mas escuchados
                            CaminoError "000-0520"
                            ReDim Preserve MTX10(zz) As String
                            CaminoError "000-0521"
                            MTX10(zz) = ruta & NombreDir
                            CaminoError "000-0522"
                            zz = zz + 1
                        End If
                        CaminoError "000-0523"
                        NombreDir = Dir$
                    Loop
BuscaMP3:
                    CaminoError "000-0524"
                    'siempre cae en el primer tema del primer directorio habilitado
                    Randomize Timer
                    Dim A As Integer, ContA As Integer
                    CaminoError "000-0525"
                    A = Int(Rnd * 1000) + 1
                    CaminoError "000-0526"
                    Dim NombreMP3 As String: zz = 0
                    CaminoError "000-0527"
                    Dim temaMP As String
                    CaminoError "000-0528"
                    Do While zz < UBound(MTX10)
                        CaminoError "000-0529"
                        NombreMP3 = Dir$(MTX10(zz) & "\*.mp3")
                        'si no hay ningun tema se va a la prox carpeta
                        CaminoError "000-0530"
                        If NombreMP3 = "" Then GoTo NextFolder
                        'da vueltas hasta encontrar un tema valido
                        CaminoError "000-0531"
                        Do While Len(NombreMP3)
                            CaminoError "000-0532"
                            temaMP = MTX10(zz) & "\" & NombreMP3
                            CaminoError "000-0533"
                            If FSO.FileExists(temaMP) Then
                                CaminoError "000-0534"
                                ContA = ContA + 1
                                CaminoError "000-0535"
                                If ContA >= A Then
                                    CaminoError "000-0536"
                                    CORTAR_TEMA = True 'este tema va cortado ya que es de 3PM para que haga ruido
                                    CaminoError "000-0537"
                                    EjecutarTema temaMP, False
                                    'solo sale cueando encuentra un tema valido
                                    CaminoError "000-0538"
                                    SecSinUso = 0
                                    Exit Sub
                                End If
                            End If
                            CaminoError "000-0539"
                            NombreMP3 = Dir$
                        Loop
NextFolder:
                        zz = zz + 1
                    Loop
                End If
                Exit Do
            End If
         Loop
         CaminoError "000-0540"
         'xxxxx
         On Local Error Resume Next
         TE.Close
        'si llego aca es por que no encontro el numero sorteado al azar en la lista
        'de los mejores. Entonces elige un tema al azar
        CaminoError "000-0541"
        GoTo AZAR
    End If
    
End Sub

Private Sub Timer3_Timer()
    CaminoError "000-0542"
    If Protector = 0 Then Timer3.Interval = 0        'para el reloj del protector. Lo ha inhabilitado
    CaminoError "000-0543"
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + 10
    lblNoTecla = Trim(Str(SecSinTecla))
    'no protector en video
    CaminoError "000-0544"
    If EsVideo Then SecSinTecla = 0
    CaminoError "000-0545"
    If SecSinTecla > EsperaTecla And EsVideo = False Then
        CaminoError "000-0546"
        frmProtect.Show 1
    End If
End Sub

Public Function TemasEnRank(MasDeXVotos) As Long
    'indica cuantos temas hay en el ranking
    CaminoError "000-0547"
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        CaminoError "000-0548"
        FSO.CreateTextFile AP + "ranking.tbr", True
        CaminoError "000-0549"
        TemasEnRankMasDeUnVoto = 0
        Exit Function
    End If
    CaminoError "000-0550"
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    CaminoError "000-0551"
    Dim TT As String
    'antes de entra ver si el archivo no tiene nada
    CaminoError "000-0552"
    If TE.AtEndOfStream Then
        CaminoError "000-0553"
        TemasEnRankMasDeUnVoto = 0
        CaminoError "000-0554"
        TE.Close
        CaminoError "000-0555"
        Exit Function
    End If
    Dim CA As Long
    CA = 0
    Dim PuntosEste  As Long
    CaminoError "000-0556"
    Do While Not TE.AtEndOfStream
        CaminoError "000-0557"
        TT = TE.ReadLine
        CaminoError "000-0558"
        PuntosEste = Val(txtInLista(TT, 0, ","))
        CaminoError "000-0559"
        If PuntosEste > MasDeXVotos Then
            CaminoError "000-0560"
            CA = CA + 1
        Else
            'todos los que siguen tienen uno (1)
            CaminoError "000-0561"
            Exit Do
        End If
    Loop
    CaminoError "000-0562"
    TE.Close
    CaminoError "000-0563"
    TemasEnRank = CA
End Function

Public Sub OrdenarListaModoVideo()
    'asegurarme que el disco elegido se ve en la lista
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    CaminoError "000-0564:" + CStr(nDiscoGral)
    If L(nDiscoGral).Top > frModoVideo.Height - (L(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        CaminoError "000-0565"
        HayQueCorrerse = L(nDiscoGral).Top - (frModoVideo.Height - (L(0).Height + 25))
        CaminoError "000-0566"
        CL = 0
        Do While CL < TOTAL_DISCOS
            CaminoError "000-0567"
            L(CL).Top = L(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    CaminoError "000-0568"
    If L(nDiscoGral).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        CaminoError "000-0569"
        HayQueCorrerse = -L(nDiscoGral).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        CaminoError "000-0570"
        Do While CL < TOTAL_DISCOS
            CaminoError "000-0571"
            L(CL).Top = L(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
End Sub

Public Sub SelTema(n As Integer)
    CaminoError "000-0571"
    T(n).BackColor = &H0&
    CaminoError "000-0572"
    T(n).ForeColor = &H80FFFF
End Sub

Public Sub UnSelTema(n As Integer)
    CaminoError "000-0573"
    T(n).BackColor = &H80FFFF
    CaminoError "000-0574"
    T(n).ForeColor = &H0&
End Sub

Public Sub OrdenarListaTemaVideo()
    'asegurarme que el disco elegido se ve en la lista
    CaminoError "000-0575"
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    CaminoError "000-0576"
    If T(TemaElegidoModoVideo).Top > frTEMAS.Height - (T(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que correrse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        CaminoError "000-0577"
        HayQueCorrerse = T(TemaElegidoModoVideo).Top - (frTEMAS.Height - (T(0).Height + 25))
        CaminoError "000-0578"
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            CaminoError "000-0579"
            T(CL).Top = T(CL).Top - HayQueCorrerse
            CaminoError "000-0580"
            CL = CL + 1
        Loop
    End If
    CaminoError "000-0581"
    If T(TemaElegidoModoVideo).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        CaminoError "000-0581"
        HayQueCorrerse = -T(TemaElegidoModoVideo).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        CaminoError "000-0582"
        Do While CL <= UBound(MATRIZ_TEMAS)
            CaminoError "000-0583"
            T(CL).Top = T(CL).Top + HayQueCorrerse
            CaminoError "000-0584"
            CL = CL + 1
        Loop
    End If
    
End Sub
