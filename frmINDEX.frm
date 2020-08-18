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
         Left            =   600
         TabIndex        =   42
         Top             =   780
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1200
         TabIndex        =   41
         Top             =   900
         Width           =   10125
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
         Caption         =   "Sin Reproducción actual Sin Reproducción actual Sin Reproducción actual Sin Reproducción actual"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   705
         Left            =   8400
         TabIndex        =   37
         Top             =   90
         Width           =   2265
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
      BackColor       =   &H00404040&
      Height          =   4290
      Left            =   0
      ScaleHeight     =   4230
      ScaleWidth      =   15360
      TabIndex        =   13
      Top             =   6930
      Width           =   15420
      Begin VB.PictureBox p1 
         Height          =   285
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   2685
         TabIndex        =   43
         Top             =   1710
         Visible         =   0   'False
         Width           =   2745
      End
      Begin tbr3pm.tbrPassImg tbrPassImg1 
         Height          =   1575
         Left            =   60
         TabIndex        =   33
         Top             =   420
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   2778
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
      Begin VB.Label lblCreditos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         Caption         =   "Credito $ 15000,00"
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
         Height          =   345
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   2565
      End
      Begin VB.Label lblTemaSonando 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sin Reproducción actual"
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
         Left            =   2670
         TabIndex        =   24
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   7455
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   2790
         TabIndex        =   29
         Top             =   1560
         Width           =   7320
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
         Height          =   240
         Left            =   2790
         TabIndex        =   28
         Top             =   1800
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
      Begin VB.Label lblPrecios 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "1 cancion $1500,00 1 video   $1700,00 1 cancion $1500,00 1 video   $1700,001 cancion $1500,00 1 video   $1700,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1275
         Left            =   2730
         TabIndex        =   22
         Top             =   330
         Width           =   2175
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
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Sin Reproducción actual"
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
         Left            =   4920
         TabIndex        =   31
         Top             =   630
         UseMnemonic     =   0   'False
         Width           =   5205
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
    'la barra invertida devuelve solo la parte entera!!!
    EnQueFilaEstoy = (nDiscoSEL \ TapasMostradasH) + 1
    tERR.Anotar "acaa", nDiscoSEL, TapasMostradasH
End Function

Private Sub cmdDiscoAd_Click()
    If MostrarTouch Then
        Form_KeyDown TeclaDER, 0
        tERR.Anotar "acag"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    tERR.Anotar "acab", KeyCode, Shift
End Sub

Private Sub cmdDiscoAt_Click()
    If MostrarTouch Then
        Form_KeyDown TeclaIZQ, 0
        tERR.Anotar "acai"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    tERR.Anotar "acac", KeyCode, Shift
End Sub

Private Sub cmdPagAd_Click()
    If MostrarTouch Then
        Form_KeyDown TeclaPagAd, 0
        tERR.Anotar "acam"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    tERR.Anotar "acad", KeyCode, Shift
End Sub

Private Sub cmdPagAt_Click()
    If MostrarTouch Then
        Form_KeyDown TeclaPagAt, 0
        tERR.Anotar "acak"
        Command1.SetFocus
    End If
End Sub

Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
    Form_KeyDown KeyCode, Shift
    tERR.Anotar "acae", KeyCode, Shift
End Sub

Private Sub Command1_Click()
    If MostrarTouch Then
        tERR.Anotar "acal"
        Form_KeyDown TeclaOK, 0
    End If
End Sub

Private Sub Form_Activate()
    On Error GoTo regERR
    tERR.Anotar "acan"
    MostrarCursor False
    
    tERR.Anotar "acaq2.STFCS"
    frmIndex.SetFocus
    
    'actualizar los precios
    '---------------------
    'si es gratis no usar!
    If CreditosCuestaTema(0) = 0 Then
        lblPrecios = "Musica Gratis"
        lblPrecios2 = "Musica Gratis"
    Else
        lblPrecios = "1 cancion   = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(0), , , , vbFalse))
        lblPrecios2 = "1 cancion=" + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(0), , , , vbFalse))
        
        If CreditosCuestaTema(1) > 0 Then
            lblPrecios = lblPrecios + vbCrLf + "2 canciones = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(1), , , , vbFalse))
            lblPrecios2 = lblPrecios2 + " 2 canciones=" + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(1), , , , vbFalse))
        End If
        
        If CreditosCuestaTema(2) > 0 Then
            lblPrecios = lblPrecios + vbCrLf + "3 canciones = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(2), , , , vbFalse))
            lblPrecios2 = lblPrecios2 + " 3 canciones=" + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTema(2), , , , vbFalse))
        End If
    End If
    
    'si es gratis no usar!
    If CreditosCuestaTemaVIDEO(0) = 0 Then
        lblPrecios = lblPrecios + vbCrLf + "Videos Gratis"
        lblPrecios2 = lblPrecios2 + " / Videos Gratis"
    Else
        lblPrecios = lblPrecios + vbCrLf + "1 video     = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(0) * (PrecioBase / TemasPorCredito), , , , vbFalse))
        lblPrecios2 = lblPrecios2 + " 1 video=" + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(0) * (PrecioBase / TemasPorCredito), , , , vbFalse))
        
        If CreditosCuestaTemaVIDEO(1) > 0 Then
            lblPrecios = lblPrecios + vbCrLf + "2 videos    = " + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(1) * (PrecioBase / TemasPorCredito), , , , vbFalse))
            lblPrecios2 = lblPrecios2 + " 2 videos=" + CStr(FormatCurrency(CreditosCuestaTemaVIDEO(1) * (PrecioBase / TemasPorCredito), , , , vbFalse))
        End If
        
        If CreditosCuestaTemaVIDEO(2) > 0 Then
            lblPrecios = lblPrecios + vbCrLf + "3 videos    = " + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTemaVIDEO(2), , , , vbFalse))
            lblPrecios2 = lblPrecios2 + " 3 videos=" + CStr(FormatCurrency((PrecioBase / TemasPorCredito) * CreditosCuestaTemaVIDEO(2), , , , vbFalse))
        End If
    End If
    
    If HabilitarVUMetro Then
        If Is3pmExclusivo Then
            tERR.Anotar "acaq"
            If VU21.inHabilitado = False And VU21.IsPlaying = False Then
                VU21.DoStart
            End If
        Else
            tERR.Anotar "acar"
            If VU1.inHabilitado = False And VU1.IsPlaying = False Then
                VU1.DoStart
            End If
        End If
    End If
    
    Exit Sub
regERR:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acap"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    On Error GoTo FallaKD
    
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    Dim PagNum As Long
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    'ver si es o no numpad
    If IsKeyPad(Me) Then
        'la falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        RealKeyCode = KeyCode
        If KeyCode = 13 Then RealKeyCode = 108
        'ademas si esta apretado el BLOQ NUM
    Else
        'de manera predeterminada son el mismo
        'salvo los casos que se especifican
        RealKeyCode = KeyCode
    End If
    tERR.Anotar "acat", KeyCode, RealKeyCode, Shift
    '----------------------------------------
    'esta tecla es IZQ en el modo 46 pasandpo de arriba aa abjo y _
        siguiendo a la pag ant en el modo 5
    'para el modo video y en modo46=5 se pasan como páginas!
    '----------------------------------------
    
    
    EsModo5PeroLabura46 = (EsVideo And _
        Salida2 = False And _
        IsMod46Teclas = 5)
    
    tERR.Anotar "acau", EsModo5PeroLabura46, EsVideo, Salida2, IsMod46Teclas
    '----------------------------------------
    Select Case RealKeyCode
        Case vbKeyF1
            frmERRORES.Show 1
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
            tERR.Anotar "acav", ToSec
            MP3.SeekTo CStr(ToSec)
        'subir o bajar volumen
        Case TeclaBajaVolumen
            If frmIndex.MP3.IsPlaying Then
                If CORTAR_TEMA = False Then 'TEMA PAGO
                    If VolumenIni <= 5 Then
                        frmIndex.MP3.Volumen = 0
                    Else
                        frmIndex.MP3.Volumen = VolumenIni - 5
                    End If
                    VolumenIni = frmIndex.MP3.Volumen
                Else 'TEMA GRATUITO VARIA VOLUMEN 2
                    If VolumenIni2 <= 5 Then
                        frmIndex.MP3.Volumen = 0
                    Else
                        frmIndex.MP3.Volumen = VolumenIni2 - 5
                    End If
                    VolumenIni2 = frmIndex.MP3.Volumen
                End If
            End If
        Case TeclaSubeVolumen
            If frmIndex.MP3.IsPlaying Then
                If CORTAR_TEMA = False Then 'TEMA PAGO
                    If VolumenIni >= 95 Then
                        frmIndex.MP3.Volumen = 100
                    Else
                        frmIndex.MP3.Volumen = VolumenIni + 5
                    End If
                    VolumenIni = frmIndex.MP3.Volumen
                Else 'TEMA GRATUITO
                    If VolumenIni2 >= 95 Then
                        frmIndex.MP3.Volumen = 100
                    Else
                        frmIndex.MP3.Volumen = VolumenIni2 + 5
                    End If
                    VolumenIni2 = frmIndex.MP3.Volumen
                End If
            End If
        Case TeclaNextMusic
            tERR.Anotar "acaw"
            EMPEZAR_SIGUIENTE
        Case TeclaPagAd
            'pase lo que pase registrar
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            
            'es para abajo en el modo 5 y pagina adelante de el modo 46
            
            If EsModo5PeroLabura46 Then
                'esto confirma que es modo 5
                tERR.Anotar "acax"
                Form_KeyDown TeclaDER, 0
            End If
            If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                'esta tecla es pagina adelante en el modo 46 y abajo en el modo 5
                tERR.Anotar "acay", nDiscoGral, TapasMostradasH, TapasMostradasV
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                
                Dim PrimeroDeLaPaginaQueSigue As Long
                PrimeroDeLaPaginaQueSigue = (PagNum + 1) * (TapasMostradasH * TapasMostradasV)
                tERR.Anotar "acaz", PrimeroDeLaPaginaQueSigue, TOTAL_DISCOS
                'NUEVO DE 6.5, pasa a la primer página
                If PrimeroDeLaPaginaQueSigue > TOTAL_DISCOS Then
                    PrimeroDeLaPaginaQueSigue = 0
                End If
                'supongo que lo puse para que no desseleccione el mismo _
                    que va a seleccionar???
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                tERR.Anotar "acba", nDiscoSEL
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueSigue, True, 1)
                lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueSigue + 1) + " de " + CStr(TOTAL_DISCOS)
                nDiscoSEL = 0
            End If
            'si esta eligiendo discos en modo video min es
            'totalmente desitinto, solo va al que sigue
            'no importann páginas ni nada
            'If EstoyEnModoVideoMiniSelDisco = False Then
            '    'xxxx
            '    Exit Sub
            'End If
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                tERR.Anotar "acbb"
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo As Long
                DiskToGo = nDiscoSEL + TapasMostradasH
                tERR.Anotar "acbb", DiskToGo, DiscosEnPagina, nDiscoGral, nDiscoSEL
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo < DiscosEnPagina Then
                    nDiscoGral = nDiscoGral + TapasMostradasH
                    UnSelDisco nDiscoSEL
                    SelDisco nDiscoSEL + TapasMostradasH
                End If
                tERR.Anotar "acbc", DiskToGo, DiscosEnPagina, nDiscoGral, nDiscoSEL
            End If
            
        Case TeclaPagAt
            If EsModo5PeroLabura46 Then
                tERR.Anotar "acbd"
                'esto confirma que es modo 5
                Form_KeyDown TeclaIZQ, 0
            End If
            If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                'esta tecla es pagina atras en el modo 46 y arriba en el modo 5
                PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
                tERR.Anotar "acbe", nDiscoGral, TapasMostradasH, TapasMostradasV, PagNum
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
                    'ahora saber que posicion ocupa el primero de los que sobran el ultima pàgina
                    tmpUbic2 = TOTAL_DISCOS - tmpUbic2
                    PrimeroDeLaPaginaQueAnterior = tmpUbic2
                End If
                tERR.Anotar "acbf", PrimeroDeLaPaginaQueAnterior, nDiscoSEL
                If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueAnterior, False, TapasMostradasV)
                tERR.Anotar "acbg", PrimeroDeLaPaginaQueAnterior, nDiscoSEL
                lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            End If
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                
                'ver que no se vaya a la mierda!!!
                Dim DiskToGo2 As Long
                DiskToGo2 = nDiscoSEL - TapasMostradasH
                tERR.Anotar "acbh", DiskToGo2
                'discos en pagina me dice cuantos hay la ultima vez que se cargo
                If DiskToGo2 >= 0 Then
                    nDiscoGral = nDiscoGral - TapasMostradasH
                    UnSelDisco nDiscoSEL
                    SelDisco nDiscoSEL - TapasMostradasH
                End If
                tERR.Anotar "acbh", nDiscoSEL
            End If
            TECLAS_PRES = TECLAS_PRES + "6"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
        Case TeclaConfig
             frmConfig.Show 1
        Case TeclaIZQ
            If ModoVideoSelTema Then
                tERR.Anotar "acbi", TemaElegidoModoVideo
                If TemaElegidoModoVideo > 0 Then
                    UnSelTema TemaElegidoModoVideo
                    TemaElegidoModoVideo = TemaElegidoModoVideo - 1
                    SelTema TemaElegidoModoVideo
                    tERR.Anotar "acbj", TemaElegidoModoVideo
                    OrdenarListaTemaVideo
                End If
                GoTo FinTeclaZ
            End If
            'no ir a -1
            'ver si es el primero
            If nDiscoSEL = 0 Then
                tERR.Anotar "acbk", nDiscoSEL
                'ver si hay que pasar hoja o no
                If PasarHoja Then
                    tERR.Anotar "acbl", nDiscoGral
                    'ver si hay páginas antes
                    'si el gral es mayor que cero entonces si hay
                    'en la primera página gral y discosel son iguales
                    If nDiscoGral > 0 Then
                        'como si viene eligiendo desde la ultima fila
                        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                            tERR.Anotar "acbm", nDiscoGral, TapasMostradasH, TapasMostradasV
                            DiscosEnPagina = CargarDiscos(nDiscoGral - _
                            ((TapasMostradasH * TapasMostradasV)), False, TapasMostradasV)
                            tERR.Anotar "acbn", nDiscoGral, nDiscoSEL
                        End If
                        
                        'busca solo la fila!!
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            tERR.Anotar "acbo", nDiscoGral, TapasMostradasH, TapasMostradasV, EnQueFilaEstoy
                            DiscosEnPagina = CargarDiscos(nDiscoGral - _
                            ((TapasMostradasH * TapasMostradasV)), False, EnQueFilaEstoy)
                            tERR.Anotar "acbp", nDiscoGral, nDiscoSEL
                        End If
                    End If
                    
                    'NUEVO 6.5 si esta en el disco cero se va a la ultima hoja
                    'o sea se hace ciclico como mprock
                    If nDiscoGral = 0 Then
                        tERR.Anotar "acbq"
                        Dim tmpUbic As Long
                        'primero ver cuantas pags ENTERAS hay!
                        tmpUbic = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
                        'despues saber cuantos discos sobran la ultima pagina
                        tmpUbic = TOTAL_DISCOS - ((TapasMostradasH * TapasMostradasV) * tmpUbic)
                        'ahora saber que posicion ocupa el primero de los que sobran el ultima pàgina
                        tmpUbic = TOTAL_DISCOS - tmpUbic
                        tERR.Anotar "acbr", tmpUbic
                        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                            tERR.Anotar "acbs", tmpUbic
                            DiscosEnPagina = CargarDiscos(tmpUbic, False, 1)
                            tERR.Anotar "acbt", nDiscoGral, nDiscoSEL
                        End If
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            tERR.Anotar "acbu", tmpUbic, EnQueFilaEstoy
                            DiscosEnPagina = CargarDiscos(tmpUbic, False, EnQueFilaEstoy)
                            tERR.Anotar "acbu", nDiscoGral, nDiscoSEL
                        End If
                    End If
                    
                Else
                    'NO NO NO!!!! nDiscoGral = (TapasMostradasH * TapasMostradasV) - 1
                    'estoy en una hoja al principio y debo elegir el disco del final
                    'sel y unsel trabajan con referencias de o al total de discos por pag
                    'nDiscoGral es el numero absoluto del disco
                    'ver si existe el disco al que voy
                    tERR.Anotar "acbv", TOTAL_DISCOS, nDiscoGral, TapasMostradasH, TapasMostradasV
                    If TOTAL_DISCOS > nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1 Then
                        nDiscoGral = nDiscoGral + (TapasMostradasH * TapasMostradasV) - 1
                        UnSelDisco nDiscoSEL
                        SelDisco (TapasMostradasH * TapasMostradasV) - 1
                        tERR.Anotar "acbw", nDiscoGral, nDiscoSEL
                    Else
                        nDiscoGral = TOTAL_DISCOS - 1
                        UnSelDisco nDiscoSEL
                        SelDisco DiscosEnPagina - 1
                        tERR.Anotar "acbx", nDiscoGral, nDiscoSEL
                    End If
                End If
            Else
                'si no es el primero ver si es
                'el primero de una fila y esta en modo 5 el teclado
                tERR.Anotar "acby", nDiscoGral, nDiscoSEL, TapasMostradasH, EnQueFilaEstoy
                If nDiscoSEL = TapasMostradasH * (EnQueFilaEstoy - 1) Then
                    'si esta en el modo 5 me fijo si esta al final de una línea
                    If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                        'el disco a iniciar ya no es nDiscoGral-(tapash*tapasv)!!!!!!
                        'hay que restar tambien el nOrden de esta pagina
                        Dim DiscoToIni As Long
                        'el primero de esta mas el total de esta!
                        DiscoToIni = nDiscoGral - nDiscoSEL - (TapasMostradasH * TapasMostradasV)
                        'ver que no se vaya a la mierda!!
                        If DiscoToIni >= 0 Then
                            tERR.Anotar "acbz", DiscoToIni, EnQueFilaEstoy
                            DiscosEnPagina = CargarDiscos(DiscoToIni, False, EnQueFilaEstoy)
                            tERR.Anotar "acbz", DiscoToIni, EnQueFilaEstoy, nDiscoGral, nDiscoSEL
                        Else
                            Dim tmpUbic3 As Long
                            'primero ver cuantas pags ENTERAS hay!
                            tmpUbic3 = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
                            'despues saber cuantos discos sobran la ultima pagina
                            tmpUbic3 = TOTAL_DISCOS - ((TapasMostradasH * TapasMostradasV) * tmpUbic3)
                            'ahora saber que posicion ocupa el primero de los que sobran el ultima pàgina
                            tmpUbic3 = TOTAL_DISCOS - tmpUbic3
                            'no tengo tiempo de hacerlo ir a la mejor fila
                            'este es el caso de la primera página hacia atras
                            'osea que le digo que se vaya a la fila 1
                            tERR.Anotar "acca", tmpUbic3, EnQueFilaEstoy, nDiscoGral, nDiscoSEL
                            DiscosEnPagina = CargarDiscos(tmpUbic3, False, EnQueFilaEstoy)
                            tERR.Anotar "accb", tmpUbic3, EnQueFilaEstoy, nDiscoGral, nDiscoSEL
                        End If
                    Else
                        'tratarlo normalmente como el 46
                        GoTo Mod46IZQ
                    End If
                Else
Mod46IZQ:
                    nDiscoGral = nDiscoGral - 1
                    tERR.Anotar "accb", nDiscoGral, nDiscoSEL
                    UnSelDisco nDiscoSEL
                    tERR.Anotar "accc", nDiscoGral, nDiscoSEL
                    SelDisco nDiscoSEL - 1
                    tERR.Anotar "accc", nDiscoGral, nDiscoSEL
                End If
            End If
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
FinTeclaZ:
            TECLAS_PRES = TECLAS_PRES + "1"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            
        Case TeclaDER
            'esta tecla es DER en el modo 46 pasandpo de abajo a arriba
            'y siguiendo a la atras ¿? sig en el modo 5
            tERR.Anotar "accd", ModoVideoSelTema, TemaElegidoModoVideo
            If ModoVideoSelTema Then
                tERR.Anotar "accd2", UBound(MATRIZ_TEMAS)
                If TemaElegidoModoVideo < UBound(MATRIZ_TEMAS) Then
                    tERR.Anotar "acce", nDiscoGral, nDiscoSEL
                    UnSelTema TemaElegidoModoVideo
                    TemaElegidoModoVideo = TemaElegidoModoVideo + 1
                    SelTema TemaElegidoModoVideo
                    tERR.Anotar "accf", nDiscoGral, nDiscoSEL
                    OrdenarListaTemaVideo
                End If
            Else
                'esta eligiendo discos ya sea en las portadas o en el modo video!!
                tERR.Anotar "accg", nDiscoGral, DiscosEnPagina, PasarHoja, TOTAL_DISCOS
                If nDiscoSEL = DiscosEnPagina - 1 Then
                    'ver si hay que pasar hojas (segun config)
                    If PasarHoja Then
                        'ver que no se vaya a la mierda!!
                        If nDiscoGral + 1 < TOTAL_DISCOS Then
                            'si esta en el modtec 46 pasa al primero
                            'pero si esta en el modo 5 pasa a su mismo nivel
                            'vertical en la hoja que sigue
                            If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                                tERR.Anotar "acch", nDiscoGral
                                'va a la primera fila!!
                                DiscosEnPagina = CargarDiscos(nDiscoGral + 1, True, 1)
                            End If
                            'busca solo la fila!!
                            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                                tERR.Anotar "acci", nDiscoGral, EnQueFilaEstoy
                                DiscosEnPagina = CargarDiscos(nDiscoGral + 1, True, EnQueFilaEstoy)
                                tERR.Anotar "accj", nDiscoGral, nDiscoSEL
                            End If
                        Else
                            'es el ultimo disco y debe empezar de cero!!!
                            If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
                                'es el ultimo disco y debe empezar de cero!!!
                                tERR.Anotar "acck", nDiscoGral, nDiscoSEL
                                DiscosEnPagina = CargarDiscos(0, True, 1)
                                tERR.Anotar "accl", nDiscoGral, nDiscoSEL
                            End If
                            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                                tERR.Anotar "accm", EnQueFilaEstoy, nDiscoGral, nDiscoSEL
                                DiscosEnPagina = CargarDiscos(0, True, EnQueFilaEstoy)
                                tERR.Anotar "accn", EnQueFilaEstoy, nDiscoGral, nDiscoSEL
                                'va a la primera fila!!
                            End If
                        End If
                    Else
                        '------------------------------
                        'si no esta configurado para pasar hojas entonces debe _
                        estar en el modo 46
                        'en el modo 5 no hay salto de página...
                        '------------------------------
                        '!!!NO NO NO nDiscoGral = 0
                        'estoy en una hoja al final y debo elegir el disco del principio
                        'sel y unsel trabajan con referencias de o al total de discos por pag
                        'nDiscoGral es el numero absoluto del disco
                        nDiscoGral = nDiscoGral - DiscosEnPagina + 1
                        UnSelDisco nDiscoSEL
                        tERR.Anotar "accp", nDiscoGral, nDiscoSEL
                        SelDisco 0
                    End If
                Else
                    'ver si llego al final de una linea horizontal para pasar a la hoja
                    'que sigue si esta en el modTeclado5
                    
                    tERR.Anotar "accq", nDiscoGral, nDiscoSEL, TOTAL_DISCOS
                    'ver si el disco existe !!! o llegamos al final de todo !!!!
                    If nDiscoGral + 1 < TOTAL_DISCOS Then
                        'si esta en el modo 5 me fijo si esta al final de una línea
                        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                            'ver ahora si es el último de una línea!!!
                            If nDiscoSEL = (TapasMostradasH * EnQueFilaEstoy) - 1 Then
                                'el disco a iniciar ya no es nDiscoGral + 1  !!!!!!
                                Dim DiscoToIni2 As Long
                                'el primero de esta mas el total de esta!
                                DiscoToIni2 = nDiscoGral - nDiscoSEL + (TapasMostradasH * TapasMostradasV)
                                'ver que no se vaya a la mierda!!
                                tERR.Anotar "accq", nDiscoGral, nDiscoSEL, TOTAL_DISCOS, DiscoToIni2
                                If DiscoToIni2 < TOTAL_DISCOS Then
                                    tERR.Anotar "accr", DiscoToIni2, EnQueFilaEstoy
                                    DiscosEnPagina = CargarDiscos(DiscoToIni2, True, EnQueFilaEstoy)
                                Else
                                    'se termino, ir a la pag1!!
                                    DiscoToIni2 = 0
                                    tERR.Anotar "accs", DiscoToIni2, EnQueFilaEstoy
                                    DiscosEnPagina = CargarDiscos(DiscoToIni2, True, EnQueFilaEstoy)
                                End If
                            Else
                                'tratarlo como el modo 46
                                GoTo Mod46
                            End If
                        Else
Mod46:
                            tERR.Anotar "acct", nDiscoGral, nDiscoSEL
                            nDiscoGral = nDiscoGral + 1
                            UnSelDisco nDiscoSEL
                            SelDisco nDiscoSEL + 1
                            tERR.Anotar "accu", nDiscoGral, nDiscoSEL
                        End If
                    End If
                End If
            End If
            
            lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
            TECLAS_PRES = TECLAS_PRES + "2"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
        Case TeclaOK
            
            TECLAS_PRES = TECLAS_PRES + "3"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            'si estoy en video
            'saber si estoy eligiendo tema. Si no estoy en disco
            tERR.Anotar "accv", nDiscoGral, nDiscoSEL, ModoVideoSelTema
            If ModoVideoSelTema Then
                'si esta en fullscreen NO EJECUTAR!!!
                'solo si no sale por la segunda salida!!!
                If EsVideo And vidFullScreen And Salida2 = False Then GoTo FinKD 'fin keydown
                'si no dice salir cargar tema
                tERR.Anotar "accw", T(TemaElegidoModoVideo)
                If T(TemaElegidoModoVideo) = "SALIR" Or T(TemaElegidoModoVideo) = "No hay temas" Then
                    'volver a elegir discos
                    frTEMAS.Visible = False
                    lblTEMAS.Visible = False
                    frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                    UnSelTema 0
                    ModoVideoSelTema = False
                Else
                    'ejecutar el tema
                    'ANTES DE VER CUANTOS CREDITOS NECESITA TENGO QUE SABER SI QUIERE EJECUTAR
                    'MP3 O VIDEO!!!!!!
                    Dim temaElegido As String
                    'lstext es una lista oculta  con datos completos
                    temaElegido = txtInLista(MATRIZ_TEMAS(TemaElegidoModoVideo), 0, "#")
                    tERR.Anotar "accx", temaElegido, CREDITOS
                    If LCase(Right(temaElegido, 3)) = "mp3" Or LCase(Right(temaElegido, 3)) = "wma" Then '''Or LCase(Right(temaElegido, 3)) = "mp4" Then
                        PideVideo = False
                    Else
                        PideVideo = True
                    End If
                                      
                    'ver si puede pagar lo que pide!!!
                    'que joyita papa!!!. Parece que supieras programar
                    '--------------------------------------------------------------
                    If (PideVideo = False And CREDITOS >= PrecNowAudio) Or _
                        (PideVideo And CREDITOS >= PrecNowVideo) Then
                    '--------------------------------------------------------------
                        'siempre que se ejecute un credito estaremos por debajo de maximo
                        OnOffCAPS vbKeyScrollLock, True
                        
                        'restar lo que corresponde!!!
                        If PideVideo Then
                            VarCreditos -PrecNowVideo
                        Else
                            VarCreditos -PrecNowAudio
                        End If
                        
                        tERR.Anotar "accy"
                        'si esta ejecutando pasa a la lista de reproducción
                        If MP3.IsPlaying Then
                            'pasar a la lista de reproducción
                            Dim NewIndLista As Long
                            NewIndLista = UBound(MATRIZ_LISTA)
                            tERR.Anotar "accz", NewIndLista, UbicDiscoActual
                            ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
                            'se graba en Matriz_Listas como path, nombre(sin .mp3)
                            MATRIZ_LISTA(NewIndLista + 1) = _
                                temaElegido + "," + _
                                FSO.GetBaseName(T(TemaElegidoModoVideo)) + _
                                " / " + FSO.GetBaseName(UbicDiscoActual)
                            tERR.Anotar "acda"
                            CargarProximosTemas
                            'graba en reini.tbr los datos que correspondan por si se corta la luz
                            CargarArchReini UCase(ReINI) 'POR LAS DUDAS que no este en mayusculas
                            'volver a elegir discos
                            frTEMAS.Visible = False
                            lblTEMAS.Visible = False
                            frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                            UnSelTema 0
                            tERR.Anotar "acdb", nDiscoSEL, nDiscoGral
                            ModoVideoSelTema = False
                        Else
                            'NUNCA ENTRARA AQUI, siempre esta rep video
                            'TEMA_REPRODUCIENDO y mp3.isplayin se cargan en ejecutartema
                            'paciencia
                            tERR.Anotar "acdc", temaElegido
                            CORTAR_TEMA = False 'este tema va entero ya que lo eligio el usuario
                            EjecutarTema temaElegido, True
                        End If
                        
                        VerSiTocaPUB
                        
                    End If
                End If
            Else
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                tERR.Anotar "acdd"
                'ver si es video debería desplegar los temas del disco elegido
                'en modo de texto
                'pero si estoy viendo el video en salida2 es video sera verdadero
                'pero de todas formas no veo als lista de texto y sigo igual
                'solo si esvideo y necesito el modo texto del video!!!!
                If EsVideo And Salida2 = False Then
                    frModoVideo.Height = frDISCOS.Height / 4
                    OrdenarListaModoVideo
                    lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
                    lblTEMAS.Left = lblModoVideo.Left
                    frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
                    frTEMAS.Height = frDISCOS.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
                    lblTEMAS.Visible = True
                    frTEMAS.Visible = True
                    'cargar los temas multimedia en t()
                    'es una matriz global
                    'en la 6.3 era nDiscoGral+1!!!
                    tERR.Anotar "acde", MATRIZ_DISCOS(nDiscoGral)
                    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
                    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
                    tERR.Anotar "acdf", UbicDiscoActual
                    ReDim Preserve MATRIZ_TEMAS(0)
                    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
                    tERR.Anotar "acdg", UBound(MATRIZ_TEMAS)
                    If UBound(MATRIZ_TEMAS) = 0 Then
                        T(0) = "No hay temas"
                        SelTema 0
                        ModoVideoSelTema = True
                        tERR.Anotar "acdh", nDiscoSEL, nDiscoGral
                        Exit Sub
                    End If
                    tERR.Anotar "acdi"
                    T(0) = "SALIR"
                    '----------------------------
                    'a daniel cruz le da un error como si se volviera a cargar algo que esta cargado
                    'por lo tanto tengo que poner un manejador de error aqui, unico lugar en que se carga esto
                    For Each LLL In frmIndex.T
                        If LLL.Index > 0 Then Unload LLL
                    Next
                    '----------------------------
                    tERR.Anotar "acdj", UBound(MATRIZ_TEMAS)
                    For AA = 1 To UBound(MATRIZ_TEMAS)
                        tERR.Anotar "acdk", AA, MATRIZ_TEMAS(AA)
                        Load T(AA)
                        T(AA) = FSO.GetBaseName(txtInLista(MATRIZ_TEMAS(AA), 1, "#"))
                        T(AA).Top = T(AA - 1).Top + T(AA - 1).Height
                        T(AA).Left = T(AA - 1).Left
                        T(AA).Visible = True
                    Next
                    tERR.Anotar "acdl", nDiscoSEL, nDiscoGral
                    TemaElegidoModoVideo = 0
                    SelTema 0
                    ModoVideoSelTema = True
                Else
                    If lblDISCO(nDiscoSEL) = "01- Los mas escuchados" Then GoTo TOP10Show
                    tERR.Anotar "acdm", lblDISCO(nDiscoSEL), nDiscoSEL, nDiscoGral
                    frmTemasDeDisco.Show 1
                End If
            End If
        Case TeclaCerrarSistema
            tERR.Anotar "acdn"
            OnOffCAPS vbKeyCapital, False
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            MostrarCursor True
            MP3.DoClose
            If ApagarAlCierre Then APAGAR_PC
            Unload Me
            End
        Case TeclaESC
            tERR.Anotar "acdo"
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            If ModoVideoSelTema Then
                tERR.Anotar "acdp", nDiscoSEL, nDiscoGral
                'volver a elegir discos
                frTEMAS.Visible = False
                lblTEMAS.Visible = False
                frModoVideo.Height = frDISCOS.Height - lblModoVideo.Height
                UnSelTema 0
                ModoVideoSelTema = False
            End If
    End Select
FinKD:
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    lblNoTecla = 0
    Exit Sub
TOP10Show:
    tERR.Anotar "acdq"
    FRMTOP10.Show 1
    Exit Sub
    
FallaKD:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acas"
    Resume Next

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Local Error GoTo FallaKD
    
    tERR.Anotar "acds", KeyCode, Shift
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
    
    tERR.Anotar "acdt", KeyCode, RealKeyCode
    
    If RealKeyCode = TeclaNewFicha Then
        
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            VarCreditos CSng(TemasPorCredito)
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
    
Exit Sub
    
FallaKD:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdr"
    Resume Next

End Sub

Private Sub Form_Load()
    
    On Error GoTo MiErr
    'imagenes no cargadas, vewr si hay algo configurado para el fondo
    Dim ImgFondo As String
    ImgFondo = Trim(LeerConfig("ImgFondo", "NO"))
    tERR.Anotar "acek", ImgFondo
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
    tERR.Anotar "acel", ImgFondo2
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
    tERR.Anotar "acem", SYSfolder, Is3pmExclusivo
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
    
    tERR.Anotar "acen"
    Prog.MIN = 0 'barra de progreso circular
    picFondoDisco.Top = 0
    picFondoDisco.Left = 0
    
    RegistroDiario 'anota la fecha, hora y numero del contador
    '--------
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "SL\indexchi.tbr") Then
            tbrPassImg1.Picture WINfolder + "SL\indexchi.tbr"
            'la imagen chiquita del exclusivo es la misma!!
            Image1.Picture = LoadPicture(WINfolder + "SL\indexchi.tbr")
        End If
    End If
    '--------
    AjustarFRM Me, 12000
    tERR.Anotar "acep", K.LICENCIA
    If K.LICENCIA = aSinCargar Then
        lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        lblDEMO = textoUsuario
    End If
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    '-----------------
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "SL\txtIDX.tbr") Then
            tERR.Anotar "aceq"
            Set TE = FSO.OpenTextFile(WINfolder + "SL\txtIDX.tbr", ForReading, False)
            Dim NewT As String
            NewT = TE.ReadAll
            lblTBR = NewT
            TE.Close
        Else
            tERR.Anotar "acer"
            lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
        End If
    Else
        tERR.Anotar "aces"
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
    End If
    '-----------------
    VU1.Width = Screen.Width
    VU1.Left = 0: VU1.Top = 0
    VU1.Height = picFondo.Top - 25
    tERR.Anotar "acet", HabilitarVUMetro, Is3pmExclusivo
    'si es exclusivo inhabilito el vumetro GRANDE !!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        'que entre en el control
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width
        frDISCOS.Left = VU1.AnchoBarra + 25 '0
    Else
        frDISCOS.Left = 0 ' tapa a las barras que no se usan 'VU1.Left + VU1.Width
        frDISCOS.Width = VU1.Width ' Screen.Width - VU1.Width
    End If
    frDISCOS.Top = 0
    frDISCOS.Height = picFondo.Top
    picFondoDisco.Height = frDISCOS.Height
    picFondoDisco.Width = frDISCOS.Width
    
    'ver si hay que mostrar el touch
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    tERR.Anotar "aceu", MostrarTouch
    If MostrarTouch = False Then
        Frame1.Visible = False 'frame del touch
        lblTemaSonando.Width = Screen.Width - lblTemaSonando.Left - 250
        lstProximos.Width = Screen.Width - lstProximos.Left - 250
        lblTBR.Width = Screen.Width - lblTBR.Left - 250
        lblDEMO.Width = Screen.Width - lblDEMO.Left - 250
    End If
    'frDISCOS contiene los discos a mostrar
    'se debera calcualr el tamaño de cada discos asi como cantidad horizontal y vertical
    Dim EspacioEntreDiscosH As Long
    Dim EspacioEntreDiscosV As Long
    Dim AnchoTapaDisco As Long
    Dim AltoTapaDisco As Long
    tERR.Anotar "acev", DistorcionarTapas
    If DistorcionarTapas Then
        EspacioEntreDiscosV = 0
        EspacioEntreDiscosH = 0
        AnchoTapaDisco = (frDISCOS.Width / TapasMostradasH)
        AltoTapaDisco = (frDISCOS.Height / TapasMostradasV)
    Else
        'el alto de estos incluye tambien el lbldisco
        AnchoTapaDisco = (frDISCOS.Width * 0.8 / TapasMostradasH)
        AltoTapaDisco = (frDISCOS.Height * 0.8 / TapasMostradasV)
        'ver cual es mayor para no permitir mucha distorsion
        'lo que se ajuste se agranda del espacio entrediscos
        EspacioEntreDiscosV = (frDISCOS.Height * 0.2 / (TapasMostradasV + 1))
        EspacioEntreDiscosH = (frDISCOS.Width * 0.2 / (TapasMostradasH + 1))
    End If
    
    tERR.Anotar "acew", MostrarRotulos
    If MostrarRotulos Then
        TapaCD(0).Width = AnchoTapaDisco
        TapaCD(0).Height = AltoTapaDisco * 0.79 '80%disco, 20% lbldisco
        lblDISCO(0).Height = AltoTapaDisco * 0.19 '80%disco, 20% lbldisco
        lblDISCO(0).Width = AnchoTapaDisco
    Else
        TapaCD(0).Width = AnchoTapaDisco
        TapaCD(0).Height = AltoTapaDisco
        lblDISCO(0).Visible = False
    End If
    'centrar!!
    Dim IniCentrarH As Long
    IniCentrarH = EspacioEntreDiscosH
    Dim IniCentrarV As Long
    IniCentrarV = EspacioEntreDiscosV
    lblDISCO(0).Left = IniCentrarH
    TapaCD(0).Left = IniCentrarH
    'ver si los rotulos van arriba o abajo
    tERR.Anotar "acex", RotulosArriba
    If RotulosArriba Then
        lblDISCO(0).Top = IniCentrarV
        If MostrarRotulos Then
            TapaCD(0).Top = lblDISCO(0).Top + lblDISCO(0).Height + 50
        Else
            TapaCD(0).Top = IniCentrarV
        End If
    Else
        tERR.Anotar "000-0269"
        TapaCD(0).Top = IniCentrarV
        tERR.Anotar "000-0271"
        lblDISCO(0).Top = TapaCD(0).Top + TapaCD(0).Height + 50
    End If
    Dim CantDiscos As Long
    CantDiscos = TapasMostradasH * TapasMostradasV
    tERR.Anotar "acey", CantDiscos
    'cargar la cantidad de tapas correspondientes
    c = 0
    Do While c < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        tERR.Anotar "acez", c
        c = c + 1
        Load TapaCD(c)
        Load lblDISCO(c)
        'ya toman el tamaño del original
        
        If c / TapasMostradasH = c \ TapasMostradasH Then
            'es una tapa al principio de linea
            If RotulosArriba Then
                lblDISCO(c).Left = IniCentrarH
                lblDISCO(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                TapaCD(c).Left = IniCentrarH
                If MostrarRotulos Then
                    TapaCD(c).Top = lblDISCO(c).Top + lblDISCO(c).Height + 50
                Else
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + 50
                End If
                TapaCD(c).Visible = True
                If MostrarRotulos Then lblDISCO(c).Visible = True
            Else
                TapaCD(c).Left = IniCentrarH
                If MostrarRotulos Then
                    TapaCD(c).Top = lblDISCO(c - TapasMostradasH).Top + lblDISCO(c - TapasMostradasH).Height + EspacioEntreDiscosV
                Else
                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
                End If
                lblDISCO(c).Left = IniCentrarH
                lblDISCO(c).Top = TapaCD(c).Top + TapaCD(c).Height + 50
                TapaCD(c).Visible = True
                If MostrarRotulos Then lblDISCO(c).Visible = True
            End If
        Else
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                lblDISCO(c).Left = lblDISCO(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                lblDISCO(c).Top = lblDISCO(c - 1).Top
                TapaCD(c).Left = lblDISCO(c).Left
                TapaCD(c).Top = TapaCD(c - 1).Top
                TapaCD(c).Visible = True
            Else
                TapaCD(c).Left = TapaCD(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                TapaCD(c).Top = TapaCD(c - 1).Top
                lblDISCO(c).Left = TapaCD(c).Left
                lblDISCO(c).Top = lblDISCO(c - 1).Top
                TapaCD(c).Visible = True
            End If
            If MostrarRotulos Then lblDISCO(c).Visible = True
        End If
        
    Loop
    tERR.Anotar "acfa"
    OnOffCAPS vbKeyScrollLock, True
    lblV = "versión " + Trim(Str(App.Major)) + "." + Trim(Str(App.Minor)) + "." + Trim(Str(App.Revision))
    lblTiempoRestante = "Falta: " + "00:00"
    'ocultar las etiquetas
    tERR.Anotar "acfa2", lblV.Caption
    Me.AutoRedraw = AutoReDibuj
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    'ver cuantos creditos hay
    CREDITOS = Val(LeerArch1Linea(AP + "creditos.tbr"))
    tERR.Anotar "acfb", CREDITOS
    
    ShowCredits
    
    'dejar cargado el mostrados de procesos
    'Load frmini
    'cargar las variables globales
    
    TEMA_REPRODUCIENDO = "Sin reproducción actual"
    TEMA_SIGUIENTE = "No hay proximo tema"
    TEMAS_EN_LISTA = 0
    
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
        
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
        tERR.Anotar "acfc", PartOrigenes(AAA)
        'ver los discos del origene elegido
        MtxTmpOrigenes() = ObtenerDir(PartOrigenes(AAA))
        'acumular a la matriz general
        SumarMatriz MATRIZ_DISCOS, MtxTmpOrigenes
    Next AAA
    
    'ya se sumop y esta listo para cargarse ordenados los discos dentro de cada origen
    MostrarDiscosMTX
    'MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
    
    Dim CarpActual As String
    Dim pathTema As String, DuracionTema As String, nombreTEMA As String
    'mostrar proceso
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    
    'ret devuelve la cantidadd de discos cargados
    tERR.Anotar "acfd"
    DiscosEnPagina = CargarDiscos(0, True, 1)
    'inicializar la matriz_lista (lista de reproduccion
    tERR.Anotar "acfe", DiscosEnPagina
    ReDim MATRIZ_LISTA(0)
    
    lblTOTdiscos = "Discos: " + Trim(Str(UBound(MATRIZ_DISCOS)))
    tERR.Anotar "acff", ReINI
    'si quedaron temas pendientes cargarlos
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            If FSO.FileExists(AP + "reini.tbr") Then
                Set TE = FSO.OpenTextFile(AP + "reini.tbr", ForReading, False)
                Dim TT As String 'cada tema
                Dim z As Integer 'contador de temas en lista anterior
                z = 1
                tERR.Anotar "acfg"
                Do While Not TE.AtEndOfStream
                    TT = TE.ReadLine
                    ReDim Preserve MATRIZ_LISTA(z)
                    MATRIZ_LISTA(z) = TT
                    z = z + 1
                Loop
                TE.Close
            End If
            EMPEZAR_SIGUIENTE
        Case "NADA"
            'no hacer nada
            'borrar la lista
            'borrra los temas 'y los creditos?
            If FSO.FileExists(AP + "reini.tbr") Then FSO.DeleteFile AP + "reini.tbr", True
            Timer1.Interval = 10000
    End Select
    
    Unload frmINI
    
    'ver si hay validacion por creditos
    Validar = LeerConfig("Validar", "0")
    tERR.Anotar "acfh", Validar
    If Validar Then
        'ver si existe el archivo Creditos Validar
        
        If FSO.FileExists(SYSfolder + "radilav.cfg") Then
            'leer el archivo de creditos vaildados
            CreditosValidar = CLng(LeerArch1Linea(SYSfolder + "radilav.cfg"))
            tERR.Anotar "acfi", CreditosValidar
            'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Else
            tERR.Anotar "acfj"
            EscribirArch1Linea SYSfolder + "radilav.cfg", "0"
            CreditosValidar = 0
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el máximo y si hay que avisar
        
        ValidarCada = LeerConfig("ValidarCada", "500")
        AvisarAntes = LeerConfig("AvisarAntes", "50")
        tERR.Anotar "acfj", CreditosValidar, ValidarCada, AvisarAntes
        If (CreditosValidar > ValidarCada - AvisarAntes) Then
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            'uso el frmClave que tiene la variable publica ClaveIngresada
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = ClaveParaValidar(CodigoParaClaveActual)
            tERR.Anotar "acfk", ClaveCorrespondiente
            Dim QuedanC As Long
            QuedanC = ValidarCada - CreditosValidar
            If QuedanC > 0 Then
                'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
                MsgBox "Ingrese a continuacion su clave para continuar utilizando 3PM. " + vbCrLf + _
                    "Debe enviar la administrador el codigo: " + vbCrLf + _
                    CodigoParaClaveActual + vbCrLf + _
                    "Puede todavia omitir esta clave. Solo le quedan " + CStr(QuedanC) + " creditos hasta que 3PM se inhabilite"
            Else
                MsgBox "De no ingresar la clave correspondiente 3PM no podra continuar. Ha llegado al limite de creditos posibles"
            End If
            tERR.Anotar "acfl"
            frmCLAVE.Show 1
            tERR.Anotar "acfm", UCase(ClaveIngresada), UCase(ClaveCorrespondiente)
            If UCase(ClaveIngresada) <> UCase(ClaveCorrespondiente) Then
                If QuedanC > 0 Then
                    MsgBox "La clave es erronea!" + vbCrLf + _
                        "Le quedan " + CStr(QuedanC) + " creditos por cargar antes que se inhabilite 3PM"
                Else
                    If K.LICENCIA <= CGratuita Then
                        MsgBox "Si hubiera una licencia cargada esta máquina estaría bloqueada!!!" + vbCrLf + "MAS CUIDADO LA PROXIMA VEZ"
                    Else 'solo lo mato si no es ua PC de prueba
                        MsgBox "No podra seguir utilizando 3PM hasta que valide con la clave correspondiente"
                        End
                    End If
                End If
            Else
                tERR.Anotar "acfn"
                'todo OK. Cargo bien la clave
                CreditosValidar = 0
                EscribirArch1Linea SYSfolder + "radilav.cfg", "0"
                'empezar un nuevo periodo
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        End If
        tERR.Anotar "acfo", ValidarCada, CodigoParaClaveActual
        lblValidar = "Val=" + CStr(ValidarCada) + "-Qued=" + CStr(ValidarCada - CreditosValidar) + "Actual=" + CStr(CreditosValidar) + " Codigo: " + CodigoParaClaveActual
    End If
    tERR.Anotar "acfj2", MostrarPUBIMG, PubliIMGCada
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
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdu"
    Resume Next
End Sub

Public Sub SelDisco(nDisco As Long)
    
    On Error GoTo MiErr
    
    lblSel.Visible = False
    tERR.Anotar "acfp", nDisco, nDiscoSEL, nDiscoGral
    lblDISCO(nDisco).ForeColor = vbBlack
    'lblDISCO(nDisco).Font.Bold = True
    lblDISCO(nDisco).Font.Underline = True
    lblDISCO(nDisco).BackColor = vbYellow
    nDiscoSEL = nDisco
    lblSel.Top = TapaCD(nDiscoSEL).Top - lblSel.BorderWidth * 10
    lblSel.Left = TapaCD(nDiscoSEL).Left - lblSel.BorderWidth * 10
    lblSel.Height = TapaCD(nDiscoSEL).Height + lblSel.BorderWidth * 20
    lblSel.Width = TapaCD(nDiscoSEL).Width + lblSel.BorderWidth * 20
    lblSel.Visible = True
    lblSel.ZOrder
    lblDISCO(nDisco).ZOrder
    
    'seleccionar de la lista de solo video
    tERR.Anotar "acfq", nDisco, nDiscoSEL, nDiscoGral
    L(nDiscoGral).ForeColor = vbWhite
    L(nDiscoGral).BackColor = vbBlack
    LastDiscoSel = nDiscoGral 'para saber cual desactivar en unsel
    If CargarIMGinicio Then
        TapaCD(nDiscoGral).BorderStyle = 1
    Else
        TapaCD(nDisco).BorderStyle = 1
    End If
    tERR.Anotar "acfr", EsVideo
    If EsVideo Then OrdenarListaModoVideo
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdv"
    Resume Next
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    On Error GoTo MiErr
    tERR.Anotar "acfs", nDisco, nDiscoSEL, nDiscoGral, LastDiscoSel
    lblDISCO(nDisco).ForeColor = vbWhite
    'lblDISCO(nDisco).Font.Bold = False
    lblDISCO(nDisco).Font.Underline = False
    lblDISCO(nDisco).BackColor = vbBlack
    'seleccionar de la lista de solo video
    tERR.Anotar "acft", LastDiscoSel, CargarIMGinicio, EsVideo
    L(LastDiscoSel).ForeColor = vbBlack
    L(LastDiscoSel).BackColor = vbWhite
    If CargarIMGinicio Then
        TapaCD(LastDiscoSel).BorderStyle = 0
    Else
        TapaCD(nDisco).BorderStyle = 0
    End If
    If EsVideo Then OrdenarListaModoVideo
        
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdw"
    Resume Next

End Sub

Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean, DeQueFila As Long) As Long
        
    On Local Error GoTo NoCRG
    
    'indicando en que disco se inicia carga ese y los seis (o lo que corresponde) _
        que le sigen
    'DeQueFial dice si es primero o último de cual fila!!!
    'devuelve el número de discos cargados
    Dim mCargarDiscos As Long
    mCargarDiscos = 0
    Dim TotPags As Long
    TotPags = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
    tERR.Anotar "acfu", numDiscoIniciar, SelPrimero, DeQueFila, TotPags
    lblPag = "Pagina " + CStr(Round(numDiscoIniciar / (TapasMostradasH * TapasMostradasV) + 1, 0)) + " de " + CStr(TotPags + 1)
    'tomar el disco que va a quedar seleccionado
    'como numero de disco en el indice general
    If SelPrimero Then
        'si la fila es uno (la primera) entonces el calculo es facil
        nDiscoGral = numDiscoIniciar + ((DeQueFila - 1) * TapasMostradasH)
    Else
        
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            nDiscoGral = numDiscoIniciar + ((TapasMostradasH * TapasMostradasV) - 1)
        End If
        If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
            nDiscoGral = numDiscoIniciar + ((TapasMostradasH * DeQueFila) - 1)
        End If
                
        'si no va a seleccionar el primero es el ultimo
        'y si no hay pàgina completa!!!!!!!!!!
        If nDiscoGral >= TOTAL_DISCOS Then nDiscoGral = TOTAL_DISCOS - 1
        
    End If
    tERR.Anotar "acfv", nDiscoGral, nDiscoSEL, TOTAL_DISCOS
    'esconder todos los discos
    Dim NDR As Long 'numero de tapa de disco real del 0 al 5 (total de discos-1)
    
    'no hacer esto al pedo si ya estan cargadas
    Dim NDI As Long '=numdiscoiniciar de la pagina
    Dim c As Integer
    c = 1
    
    NDI = numDiscoIniciar
    tERR.Anotar "acfw", NDI, CargarIMGinicio, SelPrimero
    If CargarIMGinicio Then
        If SelPrimero Then
            'si voy para adelante ocultar los que ya pase
            c = 1
            Do While c <= (TapasMostradasH * TapasMostradasV)
                tERR.Anotar "acfx", c, NDI, TapasMostradasH, TapasMostradasV
                'si no es la primera hoja!!
                If NDI >= (TapasMostradasH * TapasMostradasV) Then
                    TapaCD(NDI - c).Visible = False
                    'no se cargan lbldisco, usan solo del 0 al 5
                    lblDISCO(c - 1).Visible = False
                End If
                c = c + 1
            Loop
            tERR.Anotar "acfy"
            Me.Refresh
        Else
            'sino ocultar los de adelante
            c = 1
            Do While c <= (TapasMostradasH * TapasMostradasV)
                tERR.Anotar "acfz", c, NDI, TapasMostradasH, TapasMostradasV
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
                tERR.Anotar "acga", UltimoDeEstaPagina, UBound(MATRIZ_DISCOS)
                If UltimoDeEstaPagina > UBound(MATRIZ_DISCOS) Then
                    'si entra aca es por que la pagina elegida es la ultima
                    'y vengo volviendo desde la primera
                    'lo que hay que ocultar entonces son los discos de
                    'la primera pàgina!
                    Dim DiscoPag1Borrar As Long
                    DiscoPag1Borrar = (TapasMostradasH * TapasMostradasV) - c
                    TapaCD(DiscoPag1Borrar).Visible = False
                End If
                lblDISCO(c - 1).Visible = False
                c = c + 1
            Loop
            'Me.Refresh
        End If
    Else
        'si no se cargaron al inicio!!
        tERR.Anotar "acgb", NDR, TapasMostradasH, TapasMostradasV
        Do While NDR < ((TapasMostradasH * TapasMostradasV))
            TapaCD(NDR).Visible = False
            lblDISCO(NDR).Visible = False
            NDR = NDR + 1
        Loop
        Dim ArchTapa As String
    End If
    NDR = 0
    tERR.Anotar "acgc", NDI, numDiscoIniciar, TapasMostradasH, TapasMostradasV
    
    Do While NDI < numDiscoIniciar + ((TapasMostradasH * TapasMostradasV))
        'ver si existe si hay disco con este n°
        'el = es de la 6.5
        If NDI <= UBound(MATRIZ_DISCOS) Then
            mCargarDiscos = mCargarDiscos + 1
            'ver si ya estan cargadas o se deben cargar
            tERR.Anotar "acgd", mCargarDiscos, CargarIMGinicio, NDI
            If CargarIMGinicio Then
                TapaCD(NDI).Visible = True
                TapaCD(NDI).ZOrder
            Else
                'ver si hay tapa
                ArchTapa = txtInLista(MATRIZ_DISCOS(NDI), 0, ",")
                tERR.Anotar "acge", ArchTapa
                If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
                ArchTapa = ArchTapa + "tapa.jpg"
                If FSO.FileExists(ArchTapa) Then
                    'si la tapa es demasiado grande
                    'no colocarla
                    'xxxx
                    If FileLen(ArchTapa) > 50000 Then
                        tERR.Anotar "acgf2", NDR, ArchTapa, CStr(FileLen(ArchTapa))
                        GoTo TAPADEF
                    End If
                    tERR.Anotar "acgf", NDR
                    TapaCD(NDR).Picture = LoadPicture(ArchTapa)
                Else
TAPADEF:
                    tERR.Anotar "acgg", NDR
                    TapaCD(NDR).Picture = LoadPicture(SYSfolder + "f61.dlw")
                End If
                TapaCD(NDR).Visible = True
            End If
            'poner nombre al disco
            'antes en la 6.3 era NDI+1 !!
            lblDISCO(NDR) = txtInLista(MATRIZ_DISCOS(NDI), 1, ",")
            If MostrarRotulos Then lblDISCO(NDR).Visible = True
        End If
        tERR.Anotar "acgh", NDI, NDR
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    CargarDiscos = mCargarDiscos
    If SelPrimero Then
        tERR.Anotar "acgi", IsMod46Teclas, EsModo5PeroLabura46
        'si es modo 46 no me importa la fila!!!!
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            'y si voy de la ultima pagina incompleta hasta la primera???
            'UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
            'discos en pagina es la cantidad actual en la pagina
            'si es la ultima y esta incompleta debe saber cuantos se cargaron!!!
            
            'Y SI ES LA PRIMEERA VEZ!!!
            'UFFFFFFFFFFFFFFFFFFFF
            tERR.Anotar "acgj", DiscosEnPagina
            If DiscosEnPagina > 0 Then
                UnSelDisco DiscosEnPagina - 1
            Else
                UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
            End If
        Else
            'supone que es de la ultima columna siempre
            'pero en la 6.5 ya puede pasar al inicio de nuevo desde
            'una columna que no sea necesariamnete la ultima
            'si viene de una fila que no es la última!!!!!!
            Dim DesSelModo5 As Long
            'el (TapasMostradasH - 1) inicial supone la ultima columna
            'DesSelModo5 = (TapasMostradasH - 1) + ((DeQueFila - 1) * TapasMostradasH)
            'pero ya no es asi!!!!
            Dim ColumnaSel As Long
            'nDisco-(fila*Tapash)
            ColumnaSel = nDiscoSEL - (nDiscoSEL \ TapasMostradasH) * TapasMostradasH
            DesSelModo5 = ColumnaSel + ((DeQueFila - 1) * TapasMostradasH)
            tERR.Anotar "acgk", ColumnaSel, DesSelModo5
            UnSelDisco DesSelModo5
        End If
        tERR.Anotar "acgl", nDiscoGral, TOTAL_DISCOS
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
        tERR.Anotar "acgm", IsMod46Teclas, EsModo5PeroLabura46
        'si es modo 46 no me importa la fila!!!!
        'SI IMPORTA AHORA QUE SE PUEDE VENIR DESDEW LA PRIMEWRA PAGINA HACIA ATRAS!
        'HAY QUE ELEGIR EL ULTMODE LA ULTIA PAGINA!!!!
        If IsMod46Teclas = 46 Or EsModo5PeroLabura46 Then
            UnSelDisco 0
            'si o si la ultima!?????
            SelDisco mCargarDiscos - 1
        Else
            'tiene que desseleccionar el que venía !!
            UnSelDisco (DeQueFila - 1) * TapasMostradasH
            
            Dim DiscoSelModo5TT As Long
            DiscoSelModo5TT = ((TapasMostradasH * DeQueFila) - 1)
            'ver si esta volviendo a la ultima página desde la primera!!!
            If DiscoSelModo5TT + numDiscoIniciar >= TOTAL_DISCOS Then
                DiscoSelModo5TT = (TOTAL_DISCOS - 1) - numDiscoIniciar
            End If
            tERR.Anotar "acgn", DiscoSelModo5TT, DeQueFila
            SelDisco DiscoSelModo5TT
        End If
    End If
    Exit Function
    
NoCRG:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdx"
    Resume Next

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Constante Valor Descripción
'vbFormControlMenu 0 El usuario eligió el comando Cerrar del menú Control del formulario.
'vbFormCode 1 Se invocó la instrucción Unload desde el código.
'vbAppWindows 2 La sesión actual del entorno operativo Microsoft Windows está finalizando.
'vbAppTaskManager 3 El Administrador de tareas de Microsoft Windows está cerrando la aplicación.
'vbFormMDIForm 4 Un formulario MDI secundario se está cerrando porque el formulario MDI también se está cerrando.
'vbFormOwner 5 Un formulario se está cerrando por que su formulario propietario se está cerrando

    'Select Case UnloadMode
    '    Case 0
    '        MsgBox "El usuario eligió el comando Cerrar del menú Control " + _
    '            "del formulario."
    '    Case 1
    '        MsgBox "Se invocó la instrucción Unload desde el código."
    '    Case 2
    '        MsgBox "La sesión actual del entorno operativo Microsoft Windows " + _
    '            "está finalizando."
    '    Case 3
    '        MsgBox "El Administrador de tareas de Windows está cerrando la " + _
    '           "aplicación."
    '    Case 4
    '        MsgBox "Un formulario MDI secundario se está cerrando porque " + _
    '            "el formulario MDI también se está cerrando."
    '    Case 5
    '        MsgBox "Un formulario se está cerrando por que su formulario " + _
    '            "propietario se está cerrando"
    'End Select
    
    tERR.Anotar "acgo"
    MostrarCursor True
    'MP3.DoStop EL DOsTOP GENERA EL EVENTO ENDPLAY QUE EJECUTA EL QUE SIGUE!!!
    MP3.DoClose
    If Is3pmExclusivo Then
        VU21.DoStop
    Else
        VU1.DoStop
    End If
    tERR.Anotar "acgp"
    If ActivarERR Then tERR.StopGrabaTodo 'cierra y borra el archivo ya que se grabo OK
    'esta es para rigoberto!!!!
    End
End Sub

Private Sub MP3_BeginPlay()
    On Error GoTo MiErr
    tERR.Anotar "acgq", MP3.FileName
    Dim Tapa As String
    Tapa = FSO.GetParentFolderName(MP3.FileName) + "\tapa.jpg"
    If FSO.FileExists(Tapa) Then
        TapaEjecutando.Picture = LoadPicture(Tapa)
    Else
        TapaEjecutando.Picture = LoadPicture(SYSfolder + "f61.dlw")
    End If
    TotalTema = MP3.LengthInSec
    Ancho = lblTemaSonando.Width
    'EVITAR DIVISIONES POR CERO
    tERR.Anotar "acgr", TotalTema
    If TotalTema > 0 And MP3.IsPlaying Then
        Variacion = Ancho / TotalTema
        lblTiempoRestante = "TOTAL: " + MP3.Falta
    Else
        lblTiempoRestante = "Falta: " + "00:00"
    End If
    
    VolBajando = MP3.Volumen
    tERR.Anotar "acgs", VolBajando
    Prog.Clear
    Prog.MAX = MP3.LengthInSec
    
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdy"
    Resume Next
    
End Sub

Private Sub MP3_EndPlay()
    On Error GoTo MiErr
    tERR.Anotar "acgt"
    EstoyEnModoVideoMiniSelDisco = False
    frmIndex.TapaEjecutando.Picture = LoadPicture(SYSfolder + "f61.dlw")
    'volver a PasarHoja a su estado original3
    PasarHoja = LeerConfig("PasarHoja", "1")
    tERR.Anotar "acgt", PasarHoja, vidFullScreen, HabilitarVUMetro, Is3pmExclusivo
    VU1.Width = Screen.Width
    'ver si es fullscreen o no!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    If vidFullScreen Then
        'frDISCOS.Height = picFondo.Top
        VU1.Height = picFondo.Top
    End If
    'reacomodo si vengo de video minimo
    'tener el cuenta el exclusivo!!!!!!!!!!!!!
    If HabilitarVUMetro And Is3pmExclusivo = False Then
        frDISCOS.Width = VU1.Width - (VU1.AnchoBarra * 2) - 50 'Screen.Width - VU1.Width
        frDISCOS.Left = VU1.AnchoBarra + 25 ' VU1.Width
        'vu no se mueve si termina un video        'VU1.Top = 0        'VU1.Height = Me.Height
    Else
        frDISCOS.Width = VU1.Width ' Screen.Width
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
    tERR.Anotar "acgu", EsVideo
    
    If EsVideo Then MP3.DoClose
    'lo destapo al terminar de acomodar todos los controles en otro lado
    'picVideo.Visible = False
    lblREP.BackStyle = 0
    lblREP.ForeColor = vbWhite
    lblREP = ""
    EMPEZAR_SIGUIENTE
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdz"
    Resume Next
    
End Sub

Private Sub MP3_Played(SecondsPlayed As Long)
    On Error GoTo MiErr
    
    tERR.Anotar "acgv", SecondsPlayed
    
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
    sRest = MP3.FaltaInSec
    PorcEjecutado = MP3.PercentPlay
    
    tERR.Anotar "acgw", SecondsPlayed, PorcEjecutado, PorcentajeTEMA
    
    If PorcEjecutado > PorcentajeTEMA And CORTAR_TEMA Then
        VolBajando = VolBajando - 5 'baja 1 por segundo
        lblTemaSonando = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        lblTemaSonando2 = "Cerrando " + QuitarNumeroDeTema(FSO.GetBaseName(TEMA_REPRODUCIENDO))
        If VolBajando > 0 Then
            MP3.Volumen = VolBajando
        Else
            tERR.Anotar "acgw2"
            MP3.DoStop
            'EL DOSTOP DESENCADENA UN END PLAY QUE REALIZA UN EMPEZAR SIGUINETE
            'EMPEZAR_SIGUIENTE
        End If
    End If
    
    tERR.Anotar "acgw3"
    
    lblTiempoRestante = "Falta: " + MP3.Falta
    Prog.DibujarCirculo CDbl(SecondsPlayed)
    wi = Ancho - Variacion * (SecondsPlayed - 2)
    '=====================================
    If K.LICENCIA = aSinCargar And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        tERR.Anotar "acgw4"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        MP3.DoStop
    End If
    'cotar tambin en el gratuito
    If K.LICENCIA = CGratuita And SecondsPlayed > 126 And SecondsPlayed < TotalTema - 5 Then
        tERR.Anotar "acgw5"
        lblTemaSonando = "Tema Truncado. Version DEMO"
        lblTemaSonando2 = "Tema Truncado. Version DEMO"
        MP3.DoStop
    End If
    '=====================================
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdz"
    Resume Next

End Sub

Private Sub TapaCD_Click(Index As Integer)
    On Error GoTo MiErr
    'nunca hay que pasar hojas
    'nDiscoGral = nDiscoGral + (Index - nDiscoSEL)
    tERR.Anotar "acgx", Index, nDiscoGral, TOTAL_DISCOS, nDiscoSEL
    nDiscoGral = Index 'si se cargan todas las imágenes al inicio index=nDiscoGral
    If nDiscoGral + 1 > TOTAL_DISCOS Then
        MsgBox "No existe el disco elegido!!. " + vbCrLf + _
            "Carge discos desde el ADMINISTRADOR DE DISCOS en la " + vbCrLf + _
            "página de configuracion (presionando la tecla 'C')"
        Exit Sub
    End If
    
    UnSelDisco nDiscoSEL
    Dim PagNum As Long
    PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
    tERR.Anotar "acgy", PagNum
    nDiscoSEL = Index - (PagNum * (TapasMostradasH * TapasMostradasV))
    SelDisco nDiscoSEL
    lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
    'totar la tecla de enar a disco
    Form_KeyDown TeclaOK, 0
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acea"
    Resume Next

End Sub

Private Sub tbrPassImg1_ChangeImg()
    On Error GoTo MiErr
    'si se esta pasando un video no dar bola!!!
    tERR.Anotar "acgz", MP3.IsPlaying, EsVideo
    If MP3.IsPlaying And EsVideo Then
        frmVIDEO.picBigImg.Visible = False
    Else
        frmVIDEO.picBigImg.Visible = False
        
        'cambiar tambien las imágenes grandes de la salida de video
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
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aceb"
    Resume Next

End Sub

Private Sub Timer1_Timer()
    On Error GoTo MiErr
    
    'controla el tiempo sin uso (sin ejecucion de temas)
    If MP3.IsPlaying Then Exit Sub
    'controla el tiempo sin uso (sin ejecucion de temas)
    SecSinUso = SecSinUso + (Timer1.Interval / 1000)
    lblNoUSO = Trim(Str(SecSinUso))
    If SecSinUso >= EsperaMinutos Then 'esperaminutos esta en segundos
        tERR.Anotar "acha", SecSinUso, TemasEnRank(0), TemasEnRank(1)
        SecSinUso = 0
        Dim TemasDisponibles As Long
        If TemasEnRank(1) > 50 Then
            TemasDisponibles = TemasEnRank(1) 'todos los que se escucharon
        Else
            TemasDisponibles = TemasEnRank(0) 'todos los que se escucharon
        End If
        
        Randomize Timer
        
        z = Int(Rnd * TemasDisponibles)
        z = z + 1
        CC = 0
        tERR.Anotar "achb", z
        If FSO.FileExists(AP + "ranking.tbr") = False Then
            FSO.CreateTextFile AP + "ranking.tbr", True
            'me voy al azar ya que no hay para elegirdel rank
            tERR.Anotar "achc.NORANK"
            GoTo AZAR
        End If
        Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
        Dim TT As String
        'antes de entra ver si el archivo no tiene nada
        If TE.AtEndOfStream Then
            tERR.Anotar "achd.NORANK"
            GoTo AZAR
        End If
        Do While Not TE.AtEndOfStream
            CC = CC + 1
            TT = TE.ReadLine
            tERR.Anotar "ache", TT, CC, z
            If CC = z Then
                Dim TemaAzar As String
                TemaAzar = txtInLista(TT, 1, ",")
                'si tuve los discos cargados en una unidad o una ubicación distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                If FSO.FileExists(TemaAzar) Then
                    tERR.Anotar "achg", TemaAzar
                    CORTAR_TEMA = True 'este tema se eligio al azar no va entero
                    SecSinUso = 0
                    TE.Close
                    EjecutarTema TemaAzar, False
                    Exit Sub
                Else
AZAR:
                    tERR.Anotar "achf.AZAR"
                    'ejecutar algun tema de cualquier disco
                    Dim MTX10() As String: zz = 0
                    ruta = AP + "discos\"
                    Dim NombreDir As String
                    NombreDir = Dir$(ruta & "*.*", vbDirectory)
                    Do While Len(NombreDir)
                        If NombreDir = "." Or NombreDir = ".." Then
                            ' excluir las entradas "." y ".."
                        ElseIf (GetAttr(ruta & NombreDir) And vbDirectory) = 0 Then
                            ' este es un archivo normal
                        Else
                            'ver los primeros diez discos. En alguno tiene que haber temas
                            'yo se que el primero no tiene temas por que es
                            '01 - los mas escuchados
                            ReDim Preserve MTX10(zz) As String
                            MTX10(zz) = ruta & NombreDir
                            zz = zz + 1
                        End If
                        NombreDir = Dir$
                        tERR.Anotar "achh", NombreDir
                    Loop
BuscaMP3:
                    
                    'siempre cae en el primer tema del primer directorio habilitado
                    Randomize Timer
                    Dim A As Integer, ContA As Integer
                    A = Int(Rnd * 1000) + 1
                    Dim NombreMP3 As String: zz = 0
                    Dim temaMP As String
                    Do While zz < UBound(MTX10)
                        tERR.Anotar "achi", zz, UBound(MTX10)
                        NombreMP3 = Dir$(MTX10(zz) & "\*.mp3")
                        'si no hay ningun tema se va a la prox carpeta
                        If NombreMP3 = "" Then GoTo NextFolder
                        'da vueltas hasta encontrar un tema valido
                        tERR.Anotar "achj", NombreMP3
                        Do While Len(NombreMP3)
                            temaMP = MTX10(zz) & "\" & NombreMP3
                            tERR.Anotar "achk", temaMP
                            If FSO.FileExists(temaMP) Then
                                ContA = ContA + 1
                                If ContA >= A Then
                                    CORTAR_TEMA = True 'este tema va cortado ya que es de 3PM para que haga ruido
                                    EjecutarTema temaMP, False
                                    'solo sale cueando encuentra un tema valido
                                    SecSinUso = 0
                                    Exit Sub
                                End If
                            End If
                            NombreMP3 = Dir$
                            tERR.Anotar "achl", NombreMP3
                        Loop
NextFolder:
                        zz = zz + 1
                    Loop
                End If
                Exit Do
            End If
         Loop
         tERR.Anotar "achm.REAZAR"
         'xxxxx
         On Local Error Resume Next
         TE.Close
        'si llego aca es por que no encontro el numero sorteado al azar en la lista
        'de los mejores. Entonces elige un tema al azar
        GoTo AZAR
    End If
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acec"
    Resume Next
    
End Sub

Private Sub Timer3_Timer()
    On Error GoTo MiErr
    If Protector = 0 Then Timer3.Interval = 0
    'para el reloj del protector. Lo ha inhabilitado
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + 10
    lblNoTecla = Trim(CStr(SecSinTecla))
    'no protector en video
    If EsVideo Then SecSinTecla = 0
    tERR.Anotar "achn", SecSinTecla, EsperaTecla
    If SecSinTecla > EsperaTecla And EsVideo = False Then
        frmProtect.Show 1
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aced"
    Resume Next

End Sub

Public Function TemasEnRank(MasDeXVotos) As Long
    On Error GoTo MiErr
    'indica cuantos temas hay en el ranking
    tERR.Anotar "acho", MasDeXVotos
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        tERR.Anotar "achp"
        FSO.CreateTextFile AP + "ranking.tbr", True
        TemasEnRank = 0
        Exit Function
    End If
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    
    Dim TT As String
    'antes de entra ver si el archivo no tiene nada
    
    If TE.AtEndOfStream Then
        tERR.Anotar "achq"
        TemasEnRank = 0
        TE.Close
        Exit Function
    End If
    Dim CA As Long
    CA = 0
    Dim PuntosEste  As Long
    
    Do While Not TE.AtEndOfStream
        TT = TE.ReadLine
        tERR.Anotar "achr", TT
        PuntosEste = Val(txtInLista(TT, 0, ","))
        If PuntosEste > MasDeXVotos Then
            CA = CA + 1
        Else
            'todos los que siguen tienen uno (1)
            tERR.Anotar "achs"
            Exit Do
        End If
    Loop
    TE.Close
    TemasEnRank = CA
    
    Exit Function
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acef"
    Resume Next
End Function

Public Sub OrdenarListaModoVideo()
    On Error GoTo MiErr
    'asegurarme que el disco elegido se ve en la lista
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    tERR.Anotar "acht", nDiscoGral, nDiscoSEL, TOTAL_DISCOS
    If L(nDiscoGral).Top > frModoVideo.Height - (L(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = L(nDiscoGral).Top - (frModoVideo.Height - (L(0).Height + 25))
        CL = 0
        Do While CL < TOTAL_DISCOS
            L(CL).Top = L(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    If L(nDiscoGral).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = -L(nDiscoGral).Top
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        Do While CL < TOTAL_DISCOS
            L(CL).Top = L(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aceg"
    Resume Next
End Sub

Public Sub SelTema(n As Integer)
    T(n).BackColor = &H0&
    T(n).ForeColor = &H80FFFF
End Sub

Public Sub UnSelTema(n As Integer)
    T(n).BackColor = &H80FFFF
    T(n).ForeColor = &H0&
End Sub

Public Sub OrdenarListaTemaVideo()
    On Error GoTo MiErr
    'asegurarme que el disco elegido se ve en la lista
    tERR.Anotar "achw"
    Dim CL As Long 'contador de L
    Dim HayQueCorrerse As Long 'cuanto hay que correrse
    'para acomodar
    If T(TemaElegidoModoVideo).Top > frTEMAS.Height - (T(0).Height + 25) Then
        'esta fuera de la vista para abajo
        'correr todo para abajo
        'ver cuanto hay que correrse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = T(TemaElegidoModoVideo).Top - (frTEMAS.Height - (T(0).Height + 25))
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            T(CL).Top = T(CL).Top - HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    If T(TemaElegidoModoVideo).Top < 0 Then
        'ver cuanto hay que corresse
        'en gral es solo una casilla
        'pero si me muevo por paginas
        'puede ser mucho mas
        HayQueCorrerse = -T(TemaElegidoModoVideo).Top
        
        'esta fuera de la vista para arriba
        'correr todo para arriba
        CL = 0
        Do While CL <= UBound(MATRIZ_TEMAS)
            T(CL).Top = T(CL).Top + HayQueCorrerse
            CL = CL + 1
        Loop
    End If
    
Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aceh"
    Resume Next
    
End Sub
