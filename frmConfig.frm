VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configuracion de 3pm"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Abrir MANUAL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   8550
      Width           =   2085
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Teclado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3495
      Left            =   210
      TabIndex        =   45
      Top             =   4890
      Width           =   3435
      Begin VB.TextBox txtPagAd 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2160
         Width           =   300
      End
      Begin VB.TextBox txtPagAt 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2490
         Width           =   300
      End
      Begin VB.TextBox txtnPagAd 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   73
         Top             =   2160
         Width           =   700
      End
      Begin VB.TextBox txtnPagAt 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   72
         Top             =   2490
         Width           =   700
      End
      Begin VB.CheckBox chkApagarPC 
         BackColor       =   &H00000000&
         Caption         =   "Apagar la PC al cerrar el sistema"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   22
         Top             =   3120
         Width           =   3210
      End
      Begin VB.TextBox txtnCLOSE 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2610
         TabIndex        =   59
         Top             =   2820
         Width           =   700
      End
      Begin VB.TextBox txtnCONF 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   58
         Top             =   1845
         Width           =   700
      End
      Begin VB.TextBox txtnNewF 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   57
         Top             =   1515
         Width           =   700
      End
      Begin VB.TextBox txtnESC 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   56
         Top             =   1185
         Width           =   700
      End
      Begin VB.TextBox txtnOK 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2595
         TabIndex        =   55
         Top             =   855
         Width           =   700
      End
      Begin VB.TextBox txtnIZQ 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2600
         TabIndex        =   54
         Top             =   525
         Width           =   700
      End
      Begin VB.TextBox txtnDER 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2600
         TabIndex        =   53
         Top             =   195
         Width           =   700
      End
      Begin VB.TextBox txtCLOSE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2835
         Width           =   300
      End
      Begin VB.TextBox txtCONF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   18
         Top             =   1845
         Width           =   300
      End
      Begin VB.TextBox txtNewF 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1515
         Width           =   300
      End
      Begin VB.TextBox txtESC 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1185
         Width           =   300
      End
      Begin VB.TextBox txtOK 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   15
         Top             =   855
         Width           =   300
      End
      Begin VB.TextBox txtIZQ 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   14
         Top             =   525
         Width           =   300
      End
      Begin VB.TextBox txtDER 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2250
         MaxLength       =   1
         TabIndex        =   13
         Top             =   195
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Página Adelante"
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
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   75
         Top             =   2220
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Página Atras"
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
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   74
         Top             =   2535
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Cerrar Sistema"
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
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   52
         Top             =   2880
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Configurar"
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
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   51
         Top             =   1890
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Nueva ficha"
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
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   50
         Top             =   1560
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla SALIR"
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
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   49
         Top             =   1230
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla OK"
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
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   48
         Top             =   900
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla izquierda"
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
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   570
         Width           =   2205
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla derecha"
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
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   46
         Top             =   240
         Width           =   2200
      End
   End
   Begin VB.CheckBox chkRankToPeople 
      BackColor       =   &H00000000&
      Caption         =   "Exponer el Ranking al publico"
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
      Height          =   210
      Left            =   150
      TabIndex        =   0
      Top             =   390
      Width           =   5500
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar multimedia de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8500
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   4380
      Width           =   2700
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Quitar multimedia de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8500
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   4770
      Width           =   2700
   End
   Begin VB.TextBox txtDuracionProtect 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6510
      TabIndex        =   81
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   3300
      Width           =   600
   End
   Begin VB.VScrollBar vsDuracionProtect 
      Height          =   330
      LargeChange     =   10
      Left            =   7110
      Max             =   0
      Min             =   900
      SmallChange     =   10
      TabIndex        =   31
      Top             =   3300
      Value           =   900
      Width           =   330
   End
   Begin VB.CheckBox chkRotulosArriba 
      BackColor       =   &H00000000&
      Caption         =   "Poner los rotulos arriba de las tapas de los discos"
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
      Height          =   210
      Left            =   150
      TabIndex        =   3
      Top             =   1110
      Width           =   5500
   End
   Begin VB.CheckBox chkMostrarRotulos 
      BackColor       =   &H00000000&
      Caption         =   "Mostrar los rotulos de los discos"
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
      Height          =   210
      Left            =   150
      TabIndex        =   4
      Top             =   1350
      Width           =   5500
   End
   Begin VB.CheckBox chkCargarDuracionTemas 
      BackColor       =   &H00000000&
      Caption         =   "Cargar la duracion de los temas (demora extra)"
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
      Height          =   300
      Left            =   6360
      TabIndex        =   27
      Top             =   1530
      Width           =   5595
   End
   Begin VB.CheckBox chkProtectOriginal 
      BackColor       =   &H00000000&
      Caption         =   "Usar Protector de pantalla original (tapas de los discos)"
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
      Height          =   510
      Left            =   4380
      TabIndex        =   29
      Top             =   2400
      Width           =   3075
   End
   Begin VB.CheckBox chkDistorcionarTapas 
      BackColor       =   &H00000000&
      Caption         =   "Distorcionar tapas de los discos para ocupar 100% pantalla"
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
      Height          =   210
      Left            =   150
      TabIndex        =   2
      Top             =   870
      Width           =   5500
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Poner en 0 contador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5730
      Width           =   1935
   End
   Begin VB.CheckBox chkPasarhoja 
      BackColor       =   &H00000000&
      Caption         =   "Pasar páginas con botones de desplazamiento simple."
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
      Height          =   525
      Left            =   270
      TabIndex        =   23
      Top             =   8430
      Width           =   3495
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Mostrar Indicadores de..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1845
      Left            =   330
      TabIndex        =   80
      Top             =   2550
      Width           =   3195
      Begin VB.CheckBox chkVerLista 
         BackColor       =   &H00000000&
         Caption         =   "Lista de temas pendientes"
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
         Height          =   240
         Left            =   100
         TabIndex        =   12
         Top             =   1470
         Width           =   3000
      End
      Begin VB.CheckBox chkVerPuestoRank 
         BackColor       =   &H00000000&
         Caption         =   "Puesto Rankig tema"
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
         Height          =   240
         Left            =   100
         TabIndex        =   11
         Top             =   1230
         Width           =   3000
      End
      Begin VB.CheckBox chkVerTemasPendientes 
         BackColor       =   &H00000000&
         Caption         =   "Temas Pendientes"
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
         Height          =   240
         Left            =   100
         TabIndex        =   8
         Top             =   540
         Width           =   3000
      End
      Begin VB.CheckBox chkVerTiempoFaltante 
         BackColor       =   &H00000000&
         Caption         =   "Tiempo faltante"
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
         Height          =   240
         Left            =   100
         TabIndex        =   7
         Top             =   300
         Width           =   3000
      End
      Begin VB.CheckBox chkVerTotalDiscos 
         BackColor       =   &H00000000&
         Caption         =   "Total de discos"
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
         Height          =   240
         Left            =   100
         TabIndex        =   10
         Top             =   990
         Width           =   3000
      End
      Begin VB.CheckBox chkVerCreditos 
         BackColor       =   &H00000000&
         Caption         =   "Créditos"
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
         Height          =   240
         Left            =   100
         TabIndex        =   9
         Top             =   750
         Width           =   3000
      End
   End
   Begin VB.TextBox txtDiscosH 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   77
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1620
      Width           =   600
   End
   Begin VB.VScrollBar vsDiscosH 
      Height          =   330
      LargeChange     =   10
      Left            =   2670
      Max             =   1
      Min             =   6
      TabIndex        =   5
      Top             =   1620
      Value           =   1
      Width           =   330
   End
   Begin VB.VScrollBar vsDiscosV 
      Height          =   330
      LargeChange     =   10
      Left            =   2670
      Max             =   1
      Min             =   6
      TabIndex        =   6
      Top             =   1980
      Value           =   1
      Width           =   330
   End
   Begin VB.TextBox txtDiscosV 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2070
      TabIndex        =   76
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1965
      Width           =   600
   End
   Begin VB.CheckBox chkVUMeter 
      BackColor       =   &H00000000&
      Caption         =   "Habilitar VUMetro"
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
      Height          =   210
      Left            =   6360
      TabIndex        =   26
      Top             =   1320
      Width           =   5500
   End
   Begin VB.CheckBox chkFastINI 
      BackColor       =   &H00000000&
      Caption         =   "Inicio rápido (no mostrar imágenes en la presentación)"
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
      Height          =   330
      Left            =   6360
      TabIndex        =   28
      Top             =   1770
      Width           =   5565
   End
   Begin VB.CheckBox chkAutoReDraw 
      BackColor       =   &H00000000&
      Caption         =   "AutoRedibujado de pantalla"
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
      Height          =   210
      Left            =   150
      TabIndex        =   1
      Top             =   630
      Width           =   5500
   End
   Begin VB.TextBox txtPorcTema 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10830
      TabIndex        =   70
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2910
      Width           =   600
   End
   Begin VB.VScrollBar VsPorcTema 
      Height          =   330
      LargeChange     =   10
      Left            =   11415
      Max             =   10
      Min             =   100
      SmallChange     =   10
      TabIndex        =   37
      Top             =   2925
      Value           =   10
      Width           =   330
   End
   Begin VB.TextBox txtEsperaTecla 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6510
      TabIndex        =   66
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2925
      Width           =   600
   End
   Begin VB.VScrollBar vsEsperaTecla 
      Height          =   330
      LargeChange     =   10
      Left            =   7110
      Max             =   30
      Min             =   1200
      SmallChange     =   10
      TabIndex        =   30
      Top             =   2940
      Value           =   30
      Width           =   330
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salir sin grabar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8130
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grabar configuracion"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7710
      Width           =   2085
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Cortes de luz"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   3990
      TabIndex        =   64
      Top             =   4260
      Width           =   3585
      Begin VB.OptionButton OpReiniFull 
         BackColor       =   &H00000000&
         Caption         =   "Se ejecutan todos los temas pendientes en la lista de ejecución"
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
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   3345
      End
      Begin VB.OptionButton OpReiniNULL 
         BackColor       =   &H00000000&
         Caption         =   "Comienza de cero borrando la lista de ejecución."
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
         Height          =   375
         Left            =   150
         TabIndex        =   33
         Top             =   810
         Value           =   -1  'True
         Width           =   3240
      End
   End
   Begin VB.VScrollBar VSSegEspera 
      Height          =   330
      LargeChange     =   10
      Left            =   11415
      Max             =   30
      Min             =   1200
      SmallChange     =   10
      TabIndex        =   36
      Top             =   2580
      Value           =   30
      Width           =   330
   End
   Begin VB.TextBox txtSECwait 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10830
      TabIndex        =   63
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2580
      Width           =   600
   End
   Begin VB.VScrollBar VSmaxFichas 
      Height          =   330
      Left            =   7230
      Max             =   5
      Min             =   200
      TabIndex        =   34
      Top             =   5820
      Value           =   5
      Width           =   330
   End
   Begin VB.TextBox txtMaxFichas 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6600
      TabIndex        =   61
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   5820
      Width           =   600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Imágenes en memoria"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   6030
      TabIndex        =   42
      Top             =   270
      Width           =   6045
      Begin VB.OptionButton OpImgSIS 
         BackColor       =   &H00000000&
         Caption         =   "Cargar las imagenes a pedido"
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
         Height          =   375
         Left            =   270
         TabIndex        =   25
         Top             =   570
         Value           =   -1  'True
         Width           =   4350
      End
      Begin VB.OptionButton OpImgINI 
         BackColor       =   &H00000000&
         Caption         =   "Cargar imagenes al inicio. Recomendado"
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
         Height          =   315
         Left            =   240
         TabIndex        =   24
         Top             =   270
         Width           =   5100
      End
   End
   Begin VB.HScrollBar HSvolumen 
      Height          =   240
      LargeChange     =   10
      Left            =   4230
      Max             =   100
      TabIndex        =   35
      Top             =   6540
      Width           =   2505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   5
      X1              =   7830
      X2              =   7830
      Y1              =   2130
      Y2              =   6900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   4
      X1              =   0
      X2              =   3750
      Y1              =   4500
      Y2              =   4500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   3
      X1              =   5910
      X2              =   5910
      Y1              =   30
      Y2              =   2100
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   2
      X1              =   3750
      X2              =   3750
      Y1              =   2130
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   1
      X1              =   3780
      X2              =   12000
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      BorderWidth     =   3
      Index           =   0
      X1              =   3750
      X2              =   12000
      Y1              =   2100
      Y2              =   2100
   End
   Begin VB.Label lblHLP 
      BackColor       =   &H0000FFFF&
      Caption         =   "Detalle/Ayuda de la opcion elegida"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   3780
      TabIndex        =   90
      Top             =   7050
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Agregar/quitar multimedia de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   24
      Left            =   7900
      TabIndex        =   89
      Top             =   3960
      Width           =   4000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Otras configuraciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   23
      Left            =   3960
      TabIndex        =   88
      Top             =   3960
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Autoejecucion de temas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   22
      Left            =   8580
      TabIndex        =   87
      Top             =   2160
      Width           =   3165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Protector de pantalla"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   21
      Left            =   4320
      TabIndex        =   86
      Top             =   2130
      Width           =   2835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Aceleracion de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   20
      Left            =   6000
      TabIndex        =   85
      Top             =   30
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones de teclado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   19
      Left            =   120
      TabIndex        =   84
      Top             =   4590
      Width           =   3495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Opciones de Visualizacion/Presentacion de 3PM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   18
      Left            =   0
      TabIndex        =   83
      Top             =   60
      Width           =   6015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Duracion del protector"
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
      Height          =   285
      Index           =   17
      Left            =   4320
      TabIndex        =   82
      Top             =   3360
      Width           =   2145
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discos Horizontal"
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
      Height          =   285
      Index           =   16
      Left            =   210
      TabIndex        =   79
      Top             =   1650
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Discos Vertical"
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
      Height          =   285
      Index           =   15
      Left            =   210
      TabIndex        =   78
      Top             =   1980
      Width           =   1800
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Porcentaje ejecutar tema"
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
      Height          =   285
      Index           =   12
      Left            =   7950
      TabIndex        =   71
      Top             =   3000
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Contador de Fichas Cargadas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Index           =   10
      Left            =   7830
      TabIndex        =   69
      Top             =   5400
      Width           =   4005
   End
   Begin VB.Label lblContador 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "20264536538"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   345
      Left            =   8040
      TabIndex        =   68
      Top             =   5730
      Width           =   1890
   End
   Begin VB.Image Image2 
      Height          =   1200
      Left            =   10830
      Picture         =   "frmConfig.frx":0442
      Stretch         =   -1  'True
      Top             =   6150
      Width           =   1020
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Espera protector de pantalla"
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
      Height          =   285
      Index           =   7
      Left            =   3690
      TabIndex        =   67
      Top             =   3000
      Width           =   2805
   End
   Begin VB.Line LineSCROLL 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   4
      Visible         =   0   'False
      X1              =   4230
      X2              =   6720
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label LblVol 
      BackStyle       =   0  'Transparent
      Caption         =   "Volumen"
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
      Height          =   285
      Left            =   4200
      TabIndex        =   65
      Top             =   6270
      Width           =   1260
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Espera autoejecutar tema"
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
      Height          =   285
      Index           =   9
      Left            =   7920
      TabIndex        =   62
      Top             =   2640
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Maximo de fichas permitidas"
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
      Height          =   240
      Index           =   8
      Left            =   3990
      TabIndex        =   60
      Top             =   5895
      Width           =   2535
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TeclaConfOK As String
Dim TeclaConfESC As String

Private Sub chkApagarPC_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkApagarPC.ForeColor = vbYellow
    HLP "Tecla de cierre de 3PM. Si esta habilitado el apagado. Windows se " + _
    "cerrara tambien. Este cambio es automatico, no necesita reiniciar 3PM"
End Sub

Private Sub chkApagarPC_LostFocus()
    chkApagarPC.ForeColor = vbWhite
End Sub

Private Sub chkAutoReDraw_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkAutoReDraw.ForeColor = vbYellow
    HLP "En general no necesitara habilitar el AutoRedibujado (AutoReDraw), " + _
    "si tiene inconvenientes con la visualización de pantalla en los pasos " + _
    "de página active esta propiedad."
End Sub

Private Sub chkAutoReDraw_LostFocus()
    chkAutoReDraw.ForeColor = vbWhite
End Sub

Private Sub chkCargarDuracionTemas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkCargarDuracionTemas.ForeColor = vbYellow
    HLP "Cada vez que se habra un disco se pueden mostrar las duraciones de los" + _
    " temas. No se recomienda habilitar esta funcion salvo que cuente con un equipo potente"
End Sub

Private Sub chkCargarDuracionTemas_LostFocus()
    chkCargarDuracionTemas.ForeColor = vbWhite
End Sub

Private Sub chkDistorcionarTapas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkDistorcionarTapas.ForeColor = vbYellow
    HLP "Como 3PM permite definir la cantidad de discos mostrados por pantalla" + _
    " es posible que su eleccion no este relacionada con las proporciones de " + _
    "la pantalla. Si habilita esta opcion las fotos se distorcionaran para " + _
    "ocupar todo el espacio disponible. Caso contrario se dejara el espacio " + _
    "sobrante como libre. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkDistorcionarTapas_LostFocus()
    chkDistorcionarTapas.ForeColor = vbWhite
End Sub

Private Sub chkFastINI_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkFastINI.ForeColor = vbYellow
    HLP "A modo de atractivo grafico mientras se inicia 3PM se pueden mostrar" + _
    " todas las tapas de los discos. Si habilita esta funcion se acelerara el inicio " + _
    "y las imagenes no se mostraran"
End Sub

Private Sub chkFastINI_LostFocus()
    chkFastINI.ForeColor = vbWhite
End Sub

Private Sub chkMostrarRotulos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkMostrarRotulos.ForeColor = vbYellow
    HLP "Se recomienda dejar esta opcion habilitada, ya que sino el usuario" + _
    " final debera identificar un disco solo por su tapa (no estara disponible" + _
    " el nombre del interprete y el nombre del disco). Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkMostrarRotulos_LostFocus()
    chkMostrarRotulos.ForeColor = vbWhite
End Sub

Private Sub chkPasarhoja_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkPasarhoja.ForeColor = vbYellow
    HLP "Habilitar a las teclas de desplazamiento simple para pasar paginas. Si" + _
    " esta inhabilitado al llegar al ultimo disco de una página volvera al " + _
    "primero disco de la misma (y viceversa). Este cambio es automatico, no " + _
    "necesita reiniciar 3PM"
End Sub

Private Sub chkPasarhoja_LostFocus()
    chkPasarhoja.ForeColor = vbWhite
End Sub

Private Sub chkProtectOriginal_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectOriginal.ForeColor = vbYellow
    HLP "Puede usar para proteger la pantalla el protector por defecto. Este muestra " + _
    "las tapas de los discos. Si desea mostrar otras imagenes debera cargarlas en " + _
    "la carpeta FOTOS de la carpeta en que se instalacion y deshabilitar esta funcion. " + _
    "No use imagenes muy pesadas ya que puede afectar el rendimiento de 3PM. Se recomienda" + _
    "no sobrepasar los 100 KB"
End Sub

Private Sub chkProtectOriginal_LostFocus()
    chkProtectOriginal.ForeColor = vbWhite
End Sub

Private Sub chkRankToPeople_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkRankToPeople.ForeColor = vbYellow
    HLP "3PM a traves de la ejecucion de temas multimedia (musica o videos) " + _
    "acumula los totales por temas. Esto esta ordenado, es consultable" + _
    " y puede mostrarse o no a los usuarios finales. Si se muestra permite" + _
    " tambien cargar temas desde aqui evitando la busqueda de discos. Se " + _
    "recomienda dejar activado. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkRankToPeople_LostFocus()
    chkRankToPeople.ForeColor = vbWhite
End Sub

Private Sub chkRotulosArriba_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkRotulosArriba.ForeColor = vbYellow
    HLP "Se dice rotulo al indicador del nombre de cada disco. Esta opcion " + _
    "sirve para colocarlo encima de la foto. Si deshabilita esta opcion el rotulo " + _
    "aparecera por debajo de la foto (valor recomendado). Este cambio solo se vera una vez reiniciado" + _
    "3PM"
    
End Sub

Private Sub chkRotulosArriba_LostFocus()
    chkRotulosArriba.ForeColor = vbWhite
End Sub

Private Sub chkVerCreditos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerCreditos.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra la cantidad de credito disponible. Esto es la cantidad de temas" + _
    " que todavia se pueden programar. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerCreditos_LostFocus()
    chkVerCreditos.ForeColor = vbWhite
End Sub

Private Sub chkVerLista_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerLista.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra todos los temas pendientes de reproduccion. Si se quita este desaparecera e logo de 3PM" + _
    ". Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerLista_LostFocus()
    chkVerLista.ForeColor = vbWhite
End Sub

Private Sub chkVerPuestoRank_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerPuestoRank.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra el puesto en el ranking 3PM del tema que se este ejecutando" + _
    ". Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerPuestoRank_LostFocus()
    chkVerPuestoRank.ForeColor = vbWhite
End Sub

Private Sub chkVerTemasPendientes_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerTemasPendientes.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra la cantidad de temas pendientes de ejecucion " + _
    " ademas del tema que este ejecutandose. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerTemasPendientes_LostFocus()
    chkVerTemasPendientes.ForeColor = vbWhite
End Sub

Private Sub chkVerTiempoFaltante_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerTiempoFaltante.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra el tiempo faltante para que termine el tema que se este" + _
    " reproduciendo. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerTiempoFaltante_LostFocus()
    chkVerTiempoFaltante.ForeColor = vbWhite
End Sub

Private Sub chkVerTotalDiscos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVerTotalDiscos.ForeColor = vbYellow
    HLP "De los indicadores de la parte inferior de la pantalla este" + _
    "muestra la cantidad de discos totales y el numero de disco elegido" + _
    ". Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVerTotalDiscos_LostFocus()
    chkVerTotalDiscos.ForeColor = vbWhite
End Sub

Private Sub chkVUMeter_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVUMeter.ForeColor = vbYellow
    HLP "Se llama VuMetro al medidor de nivel de sonido. Este es muy" + _
    " atractivo a la vista pero consume muchos recursos de la PC. Por esto" + _
    " solo deberá usarse cuando el rendimiento del equipo no se vea afectado " + _
    "con el uso de este. Para PCs de bajos recursos (procesador y RAM)" + _
    " se recomienda dejar desactivado. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub chkVUMeter_LostFocus()
    chkVUMeter.ForeColor = vbWhite
End Sub

Private Sub Command1_Click()
    'cargar los datos del archivo config.tbr
    Dim TE As TextStream
    Set TE = FSO.CreateTextFile(AP + "config.tbr", True)
    TE.WriteLine "CargarImagenInicio=" + CStr(OpImgINI)
    TE.WriteLine "AutoReDraw=" + CStr(chkAutoReDraw)
    TE.WriteLine "TeclaDerecha=" + txtnDER
    TE.WriteLine "TeclaIzquierda=" + txtnIZQ
    TE.WriteLine "TeclaPagAd=" + txtnPagAd
    TE.WriteLine "TeclaPagAt=" + txtnPagAt
    TE.WriteLine "TeclaOK=" + txtnOK
    TE.WriteLine "TeclaESC=" + txtnESC
    TE.WriteLine "TeclaNuevaFicha=" + txtnNewF
    TE.WriteLine "TeclaConfig=" + txtnCONF
    TE.WriteLine "TeclaCerrarSistema=" + txtnCLOSE
    TE.WriteLine "ApagarAlCierre= " + CStr(chkApagarPC)
    TE.WriteLine "RankToPeople= " + CStr(chkRankToPeople)
    TE.WriteLine "MaximoFichas=" + txtMaxFichas
    TE.WriteLine "EsperaMinutos=" + txtSECwait
    TE.WriteLine "FastIni=" + CStr(chkFastINI)
    TE.WriteLine "HabilitarVUMetro=" + CStr(chkVUMeter)
    'Valores de ReIni LISTA=solo lista NADA=arranca de cero
    If OpReiniFull Then
        TE.WriteLine "ReINI=LISTA"
    Else
        TE.WriteLine "ReINI=NADA"
    End If
    TE.WriteLine "Volumen=" + Trim(Str(HSvolumen))
    TE.WriteLine "EsperaTecla=" + txtEsperaTecla
    TE.WriteLine "PorcentajeTema=" + txtPorcTema
    TE.WriteLine "DiscosH=" + txtDiscosH
    TE.WriteLine "DiscosV=" + txtDiscosV
    TE.WriteLine "DuracionProtect=" + txtDuracionProtect
    
    TE.WriteLine "verTiempoRestante=" + CStr(chkVerTiempoFaltante)
    TE.WriteLine "verTemasEnLista=" + CStr(chkVerTemasPendientes)
    TE.WriteLine "verCreditos=" + CStr(chkVerCreditos)
    TE.WriteLine "verTOTdiscos=" + CStr(chkVerTotalDiscos)
    TE.WriteLine "verPuesto=" + CStr(chkVerPuestoRank)
    TE.WriteLine "verLista=" + CStr(chkVerLista)
    TE.WriteLine "PasarHoja=" + CStr(chkPasarhoja)
    TE.WriteLine "DistorcionarTapas=" + CStr(chkDistorcionarTapas)
    TE.WriteLine "ProtectOriginal=" + CStr(chkProtectOriginal)
    TE.WriteLine "CargarDuracionTemas=" + CStr(chkCargarDuracionTemas)
    TE.WriteLine "MostrarRotulos=" + CStr(chkMostrarRotulos)
    TE.WriteLine "RotulosArriba=" + CStr(chkRotulosArriba)
    
    TE.Close
    
    'todas las propiedades se quedan sin reiniciar
    'algunas no se necesitan
    'NO NECESITO CargarIMGinicio = LeerConfig("CargarImagenInicio")
    AutoReDibuj = LeerConfig("AutoReDraw")
    'NO DEBO TeclaDER = Val(LeerConfig("TeclaDerecha"))
    'NO DEBO TeclaIZQ = Val(LeerConfig("TeclaIzquierda") )
    'NO DEBO TeclaOK = Val(LeerConfig("TeclaOK"))
    'NO DEBO TeclaESC = Val(LeerConfig("TeclaESC"))
    'NO DEBO TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha"))
    'NO DEBO TeclaConfig = Val(LeerConfig("TeclaConfig"))
    'NO DEBO TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre")
    MaximoFichas = Val(LeerConfig("MaximoFichas"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos"))
    'NO DEBO ReINI = LeerConfig("ReINI")
    VolumenIni = CLng(LeerConfig("Volumen"))
    EsperaTecla = Val(LeerConfig("EsperaTecla"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema"))
    'NO NECESITO FASTini = LeerConfig("FastIni")
    PasarHoja = LeerConfig("PasarHoja")
    ProtectOriginal = LeerConfig("ProtectOriginal")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas")
    VolumenIni = HSvolumen
    DuracionProtect = LeerConfig("DuracionProtect")
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command1.BackColor = vbYellow
    HLP "Grabar los datos cargados"
End Sub

Private Sub Command1_LostFocus()
    Command1.BackColor = &HFFC0C0
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command2_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command2.BackColor = vbYellow
    HLP "Salir ignorando los cambios realizados"
End Sub

Private Sub Command2_LostFocus()
    Command2.BackColor = &HFFC0C0
End Sub

Private Sub Command3_Click()
    frmCLAVE.Show 1
    'ver que la contraseña se tome desde el teclado al usuario
    If ClaveIngresada = "fRF4247L000wZ" Then '13 caracteres
        SumarContadorCreditos -CONTADOR 'esto lo deja en cero
        lblContador = STRceros(CONTADOR, 11)
    Else
        MsgBox "La clave ingresada no es correcta"
    End If
End Sub

Private Sub Command3_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command3.BackColor = vbYellow
    HLP "Dejar en cero el contador de creditos, requiere el uso del teclado y una " + _
    "contraseña"
End Sub

Private Sub Command3_LostFocus()
    Command3.BackColor = &HFFC0C0
End Sub

Private Sub Command4_Click()
    frmAddRemoveMusic.Show 1
End Sub

Private Sub Command4_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command4.BackColor = vbYellow
    HLP "Quitar discos o temas de 3PM. Requiere el uso del teclado "
End Sub

Private Sub Command4_LostFocus()
    Command4.BackColor = &HFFC0C0
End Sub

Private Sub Command5_Click()
    frmAddMusic.Show 1
End Sub

Private Sub Command5_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command5.BackColor = vbYellow
    HLP "Agregar desde CD o desde otro disco mas discos a 3PM. Requiere el uso del teclado "
End Sub

Private Sub Command5_LostFocus()
    Command5.BackColor = &HFFC0C0
End Sub

Private Sub Command7_Click()
    AbrirArchivo AP + "manual.doc", Me
End Sub

Private Sub Command7_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command7.BackColor = vbYellow
    HLP "Abrir el manual de uso de 3PM"
End Sub

Private Sub Command7_LostFocus()
    Command7.BackColor = &HFFC0C0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmINDEX.MP3.DoClose
            End
        Case TeclaNewFicha
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS <= MaximoFichas Then
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + 1
                SumarContadorCreditos 1
                lblContador = STRceros(CONTADOR, 11)
                If CREDITOS >= 10 Then
                    frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Else
                    frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                End If
            Else
                OnOffCAPS vbKeyScrollLock, False
            End If
        Case TeclaDER
            SendKeys "{TAB}"
        Case TeclaIZQ
            SendKeys "+{TAB}"
        Case TeclaOK
            SendKeys TeclaConfOK
        Case TeclaESC
            SendKeys TeclaConfESC
    End Select
    SecSinTecla = 0
    frmINDEX.lblNoTecla = 0
End Sub

Private Sub Form_Load()
    MostrarCursor True
    AjustarFRM Me, 12000
    lblContador = STRceros(CONTADOR, 11)
    
    lblTIT = "3PM - Sistema de reproducción de ficheros MP3." + vbCrLf + vbCrLf + _
    "Este sistema se distribuye sin ficheros MP3 y esta pensado para su utilización" + _
    " en lugares publicos como herramienta de entretenimiento. De ninguna manera " + _
    "deberá utilizarse para difundir ficheros cuya expresa autorización no haya " + _
    "sido solicitada a los titulares de los mismos. Los autores de 3PM creen " + _
    "firmemente en el respeto a los derechos de autor. Por lo tanto solo se podrá" + _
    " hacer uso de este sistema sobre la base de una utlización dentro del marco " + _
    "que impone la ley en en este sentido. " + vbCrLf + _
    "La reponsabilidad del uso de este sistema cae en los usuarios finales y " + _
    "los autores del sistema no se hacen responsables por utilizaciones fuera del " + _
    "marco legal del pais en que se utilize"
    
    'leer el archivo de configuracion ap+"config.tbr"
    CargarIMGinicio = LeerConfig("CargarImagenInicio")
    AutoReDibuj = LeerConfig("AutoReDraw")
    TeclaDER = Val(LeerConfig("TeclaDerecha"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda"))
    TeclaOK = Val(LeerConfig("TeclaOK"))
    TeclaESC = Val(LeerConfig("TeclaESC"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha"))
    TeclaConfig = Val(LeerConfig("TeclaConfig"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre")
    MaximoFichas = Val(LeerConfig("MaximoFichas"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos"))
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI")
    VolumenIni = CLng(LeerConfig("Volumen"))
    EsperaTecla = Val(LeerConfig("EsperaTecla"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema"))
    FASTini = LeerConfig("FastIni")
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro")
    verTiempoRestante = LeerConfig("VerTiempoRestante")
    verTemasEnLista = LeerConfig("verTemasEnLista")
    verCreditos = LeerConfig("verCreditos")
    verTOTdiscos = LeerConfig("verTotDiscos")
    verPuesto = LeerConfig("verPuesto")
    verLista = LeerConfig("verLista")
    PasarHoja = LeerConfig("PasarHoja")
    DistorcionarTapas = LeerConfig("DistorcionarTapas")
    ProtectOriginal = LeerConfig("ProtectOriginal")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas")
    MostrarRotulos = LeerConfig("MostrarRotulos")
    RotulosArriba = LeerConfig("RotulosArriba")
    DuracionProtect = LeerConfig("DuracionProtect")
    RankToPeople = LeerConfig("RankToPeople")
    
    'las variables ya se cargaron al inicio
    OpImgINI = CargarIMGinicio
    chkAutoReDraw = -AutoReDibuj
    txtnDER = TeclaDER
    txtDER = Chr(TeclaDER)
    txtnIZQ = TeclaIZQ
    txtIZQ = Chr(TeclaIZQ)
    txtnOK = TeclaOK
    txtOK = Chr(TeclaOK)
    txtnESC = TeclaESC
    txtESC = Chr(TeclaESC)
    txtnPagAd = TeclaPagAd
    txtPagAd = Chr(TeclaPagAd)
    txtnPagAt = TeclaPagAt
    txtPagAt = Chr(TeclaPagAt)
    txtnNewF = TeclaNewFicha
    txtNewF = Chr(TeclaNewFicha)
    txtnCONF = TeclaConfig
    txtCONF = Chr(TeclaConfig)
    txtnCLOSE = TeclaCerrarSistema
    txtCLOSE = Chr(TeclaCerrarSistema)
    chkApagarPC = -ApagarAlCierre
    chkVerTiempoFaltante = -verTiempoRestante
    chkVerTemasPendientes = -verTemasEnLista
    chkVerCreditos = -verCreditos
    chkVerTotalDiscos = -verTOTdiscos
    chkVerPuestoRank = -verPuesto
    chkVerLista = -verLista
    chkDistorcionarTapas = -DistorcionarTapas
    chkRankToPeople = -RankToPeople
    txtMaxFichas = MaximoFichas
    VSmaxFichas = MaximoFichas
    txtSECwait = EsperaMinutos
    VSSegEspera = EsperaMinutos
    vsDuracionProtect = DuracionProtect
    'Valores de ReIni LISTA=solo lista NADA=arranca de cero
    If ReINI = "LISTA" Then
        OpReiniFull = True
    Else
        OpReiniNULL = True
    End If
    HSvolumen = VolumenIni
    LblVol = "Volumen: " + CStr(VolumenIni)
    txtEsperaTecla = EsperaTecla
    vsEsperaTecla = EsperaTecla
    txtPorcTema = PorcentajeTEMA
    VsPorcTema = PorcentajeTEMA
    chkFastINI = -FASTini
    chkVUMeter = -HabilitarVUMetro
    vsDiscosH = TapasMostradasH
    vsDiscosV = TapasMostradasV
    TeclaConfOK = "{UP}"
    TeclaConfESC = "{DOWN}"
    chkPasarhoja = -PasarHoja
    chkProtectOriginal = -ProtectOriginal
    chkCargarDuracionTemas = -CargarDuracionTemas
    
    chkMostrarRotulos = -MostrarRotulos
    chkRotulosArriba = -RotulosArriba
    
End Sub

Private Sub HSvolumen_Change()
    If frmINDEX.MP3.IsPlaying Then frmINDEX.MP3.Volumen = HSvolumen
    LblVol = "Volumen: " + Trim(Str(HSvolumen))
    HLP "Volumen del sonido actual."
End Sub

Private Sub HSvolumen_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineSCROLL.Visible = True
End Sub

Private Sub HSvolumen_LostFocus()
    LineSCROLL.Visible = False
End Sub

Private Sub OpImgINI_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpImgINI.ForeColor = vbYellow
    HLP "Todas las imagenes se cargan en memoria al iniciar el sistema. " + _
    "Arranque del sistema mas lento, funcionamiento general más agil. " + _
    "Recomendado para PCs viejas"
End Sub

Private Sub OpImgINI_LostFocus()
    OpImgINI.ForeColor = vbWhite
End Sub

Private Sub OpImgSIS_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpImgSIS.ForeColor = vbYellow
    HLP "Las imágenes se cargan a pedido durante el uso del sistema. " + _
    "Arranque rápido. Recomendado para PCs nuevas"
End Sub

Private Sub OpImgSIS_LostFocus()
    OpImgSIS.ForeColor = vbWhite
End Sub

Private Sub OpReiniFull_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpReiniFull.ForeColor = vbYellow
    HLP "Al iniciar 3PM este ejecuta todos los temas pendientes" + _
    " de reproduccion que habia al cerrarse 3PM"
End Sub

Private Sub OpReiniFull_LostFocus()
    OpReiniFull.ForeColor = vbWhite
End Sub

Private Sub OpReiniNULL_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpReiniNULL.ForeColor = vbYellow
    HLP "Al iniciar 3PM este borra (no ejecuta) todos los temas pendientes" + _
    " de reproduccion que habia al cerrarse 3PM"
End Sub

Private Sub OpReiniNULL_LostFocus()
    OpReiniNULL.ForeColor = vbWhite
End Sub

Private Sub txtCLOSE_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCLOSE.BackColor = vbYellow
    txtnCLOSE.BackColor = vbYellow
    
End Sub

Private Sub txtCLOSE_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnCLOSE = KeyCode
    txtCLOSE = Chr(KeyCode)
End Sub

Private Sub txtCLOSE_LostFocus()
    txtCLOSE.BackColor = vbWhite
    txtnCLOSE.BackColor = vbWhite
End Sub

Private Sub txtCONF_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCONF.BackColor = vbYellow
    txtnCONF.BackColor = vbYellow
    HLP "Tecla de para ingresar a esta pagina de configuracion" + _
    ". Se recomienda usar la tecla ENTER." + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtCONF_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnCONF = KeyCode
    txtCONF = Chr(KeyCode)
End Sub

Private Sub txtCONF_LostFocus()
    txtCONF.BackColor = vbWhite
    txtnCONF.BackColor = vbWhite
End Sub

Private Sub txtDER_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDER.BackColor = vbYellow
    txtnDER.BackColor = vbYellow
    HLP "Tecla de desplazamiento de disco a la derecha" + _
    ". Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtDER_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnDER = KeyCode
    txtDER = Chr(KeyCode)
End Sub

Private Sub txtDER_LostFocus()
    txtDER.BackColor = vbWhite
    txtnDER.BackColor = vbWhite
End Sub

Private Sub txtESC_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtESC.BackColor = vbYellow
    txtnESC.BackColor = vbYellow
    HLP "Tecla de salida. Se usa para salir de un discos sin " + _
    "ejecutar algun tema. Se recomienda usar la tecla ESC." + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtESC_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnESC = KeyCode
    txtESC = Chr(KeyCode)
End Sub

Private Sub txtEXE_KeyDown(KeyCode As Integer, Shift As Integer)
    txtnEXE = KeyCode
    txtEXE = Chr(KeyCode)
End Sub

Private Sub txtESC_LostFocus()
    txtESC.BackColor = vbWhite
    txtnESC.BackColor = vbWhite
End Sub

Private Sub txtIZQ_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtIZQ.BackColor = vbYellow
    txtnIZQ.BackColor = vbYellow
    HLP "Tecla de desplazamiento de disco a la izquierda" + _
    ". Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtIZQ_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnIZQ = KeyCode
    txtIZQ = Chr(KeyCode)
End Sub

Private Sub txtIZQ_LostFocus()
    txtIZQ.BackColor = vbWhite
    txtnIZQ.BackColor = vbWhite
End Sub

Private Sub txtNewF_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtNewF.BackColor = vbYellow
    txtnNewF.BackColor = vbYellow
    HLP "Tecla de carga de credito. Esta tecla no estra expuesta al publico." + _
    " Esta tecla se debera conectar al receptor de fichas o monedas." + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtNewF_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnNewF = KeyCode
    txtNewF = Chr(KeyCode)
End Sub

Private Sub txtNewF_LostFocus()
    txtNewF.BackColor = vbWhite
    txtnNewF.BackColor = vbWhite
End Sub

Private Sub txtOK_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtOK.BackColor = vbYellow
    txtnOK.BackColor = vbYellow
    HLP "Tecla de aceptacion. Se usa para ingresar a un discos o para" + _
    " ejecutar algun tema. Se recomienda usar la tecla ENTER." + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtOK_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnOK = KeyCode
    txtOK = Chr(KeyCode)
End Sub

Private Sub txtOK_LostFocus()
    txtOK.BackColor = vbWhite
    txtnOK.BackColor = vbWhite
End Sub

Private Sub txtPagAd_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPagAd.BackColor = vbYellow
    txtnPagAd.BackColor = vbYellow
    HLP "Tecla de desplazamiento de pagina completa a la derecha (adelante)" + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnPagAd = KeyCode
    txtPagAd = Chr(KeyCode)
End Sub

Private Sub txtPagAd_LostFocus()
    txtPagAd.BackColor = vbWhite
    txtnPagAd.BackColor = vbWhite
End Sub

Private Sub txtPagAt_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPagAt.BackColor = vbYellow
    txtnPagAt.BackColor = vbYellow
    HLP "Tecla de desplazamiento de pagina completa a la izquierda (atras)" + _
    " Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub txtPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnPagAt = KeyCode
    txtPagAt = Chr(KeyCode)
End Sub

Private Sub txtPagAt_LostFocus()
    txtPagAt.BackColor = vbWhite
    txtnPagAt.BackColor = vbWhite
End Sub

Private Sub vsDiscosH_Change()
    txtDiscosH = vsDiscosH
    
End Sub

Private Sub vsDiscosH_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDiscosH.BackColor = vbYellow
    HLP "Cantidad de discos que se distribuiran horizontalmente. tbrSoft" + _
    " recomienda usar 5 (y 3 vertical). Puede usted probar distintos " + _
    "valores que sean de su agrado. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub vsDiscosH_LostFocus()
    txtDiscosH.BackColor = vbWhite
End Sub

Private Sub vsDiscosV_Change()
    txtDiscosV = vsDiscosV
    
End Sub

Private Sub vsDiscosV_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDiscosV.BackColor = vbYellow
    HLP "Cantidad de discos que se distribuiran verticalmente. tbrSoft" + _
    " recomienda usar 3 (y 5 horizontal). Puede usted probar distintos " + _
    "valores que sean de su agrado. Este cambio solo se vera una vez reiniciado" + _
    "3PM"
End Sub

Private Sub vsDiscosV_LostFocus()
    txtDiscosV.BackColor = vbWhite
End Sub

Private Sub vsDuracionProtect_Change()
    txtDuracionProtect = vsDuracionProtect
End Sub

Private Sub vsDuracionProtect_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDuracionProtect.BackColor = vbYellow
    HLP "Tiempo en segundos que el protector de pantalla se muestra" + _
    ". Si deja en cero el protector de pantalla solo se desactivara " + _
    "con la presion de alguna tecla"
End Sub

Private Sub vsDuracionProtect_LostFocus()
    txtDuracionProtect.BackColor = vbWhite
End Sub

Private Sub vsEsperaTecla_Change()
    txtEsperaTecla = vsEsperaTecla
End Sub

Private Sub vsEsperaTecla_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtEsperaTecla.BackColor = vbYellow
    HLP "Tiempo en segundos que deben pasar (sin la presion de ninguna tecla)" + _
    " para que se active el protector de pantalla."
End Sub

Private Sub vsEsperaTecla_LostFocus()
    txtEsperaTecla.BackColor = vbWhite
End Sub

Private Sub VSmaxFichas_Change()
    txtMaxFichas = VSmaxFichas
End Sub

Private Sub VSmaxFichas_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtMaxFichas.BackColor = vbYellow
    HLP "Si se cargan mas creditos (fichas, monedas) que este valor 3PM" + _
    " no los tomara y se perderan"
End Sub

Private Sub VSmaxFichas_LostFocus()
    txtMaxFichas.BackColor = vbWhite
End Sub

Private Sub VsPorcTema_Change()
    txtPorcTema = VsPorcTema
End Sub

Private Sub VsPorcTema_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPorcTema.BackColor = vbYellow
    HLP "Porcentaje de tema ejecutado automaticamente que se va a reproducir." + _
    " Si deja en 100 los temas automaticos se reproduciran completamente, de lo" + _
    " contrario se cortaran."
End Sub

Private Sub VsPorcTema_LostFocus()
    txtPorcTema.BackColor = vbWhite
End Sub

Private Sub VSSegEspera_Change()
    txtSECwait = VSSegEspera
End Sub

Private Sub VSSegEspera_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtSECwait.BackColor = vbYellow
    HLP "Tiempo en segundos que deben pasar (sin la ejecucion de ningun tema)" + _
    " para que se autoejecute algun tema. Este es sacado del ranking al azar"
End Sub

Private Sub VSSegEspera_LostFocus()
    txtSECwait.BackColor = vbWhite
End Sub

Public Sub HLP(TXT As String)
    lblHLP = "Detalle/Ayuda de la opcion elegida:" + vbCrLf + TXT
End Sub
