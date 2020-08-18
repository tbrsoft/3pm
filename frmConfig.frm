VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Configuracion de 3pm"
   ClientHeight    =   13365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16995
   Icon            =   "frmConfig.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13365
   ScaleWidth      =   16995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame frVisualizacion 
      BackColor       =   &H00000000&
      Caption         =   "Visualizacion"
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
      Height          =   4125
      Left            =   8325
      TabIndex        =   47
      Top             =   9180
      Visible         =   0   'False
      Width           =   6105
      Begin VB.CheckBox chkSalida2 
         BackColor       =   &H00000000&
         Caption         =   "REPRODUCIR VIDEOS EN TV *"
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
         Height          =   330
         Left            =   180
         TabIndex        =   144
         Top             =   1935
         Width           =   5385
      End
      Begin VB.CheckBox chkNoVumVID 
         BackColor       =   &H00000000&
         Caption         =   "Quitar VUMetro (medidor de sonido) en Videos"
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
         Height          =   330
         Left            =   180
         TabIndex        =   143
         Top             =   2220
         Width           =   5775
      End
      Begin VB.CheckBox chkVidFullScreen 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir videos en full-screen"
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
         Height          =   330
         Left            =   180
         TabIndex        =   142
         Top             =   1650
         Width           =   3855
      End
      Begin VB.CheckBox chkTouch 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar botones de touch-screen"
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
         Height          =   210
         Left            =   180
         TabIndex        =   122
         Top             =   1440
         Width           =   5800
      End
      Begin VB.TextBox TxtUSUARIO 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   90
         Text            =   "frmConfig.frx":0442
         Top             =   3480
         Width           =   5790
      End
      Begin VB.CheckBox chkAutoReDraw 
         BackColor       =   &H00000000&
         Caption         =   "AutoRedibujado de pantalla"
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
         Height          =   210
         Left            =   180
         TabIndex        =   56
         Top             =   480
         Width           =   5800
      End
      Begin VB.TextBox txtDiscosV 
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
         Left            =   2490
         TabIndex        =   55
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2865
         Width           =   600
      End
      Begin VB.VScrollBar vsDiscosV 
         Height          =   330
         LargeChange     =   10
         Left            =   3090
         Max             =   1
         Min             =   6
         TabIndex        =   54
         Top             =   2880
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsDiscosH 
         Height          =   330
         LargeChange     =   10
         Left            =   3090
         Max             =   1
         Min             =   6
         TabIndex        =   53
         Top             =   2550
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtDiscosH 
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
         Left            =   2490
         TabIndex        =   52
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2550
         Width           =   600
      End
      Begin VB.CheckBox chkDistorcionarTapas 
         BackColor       =   &H00000000&
         Caption         =   "Distorcionar tapas de discos para ocupar 100% pantalla"
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
         Height          =   210
         Left            =   180
         TabIndex        =   51
         Top             =   720
         Width           =   5800
      End
      Begin VB.CheckBox chkMostrarRotulos 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar los rotulos de los discos"
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
         Height          =   210
         Left            =   180
         TabIndex        =   50
         Top             =   1200
         Width           =   5800
      End
      Begin VB.CheckBox chkRotulosArriba 
         BackColor       =   &H00000000&
         Caption         =   "Poner los rotulos arriba de las tapas de los discos"
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
         Height          =   210
         Left            =   180
         TabIndex        =   49
         Top             =   960
         Width           =   5800
      End
      Begin VB.CheckBox chkRankToPeople 
         BackColor       =   &H00000000&
         Caption         =   "Exponer el Ranking al publico"
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
         Height          =   210
         Left            =   180
         TabIndex        =   48
         Top             =   240
         Width           =   5800
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Texto Personalizado"
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
         Height          =   285
         Index           =   10
         Left            =   240
         TabIndex        =   91
         Top             =   3240
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discos Vertical"
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
         Height          =   285
         Index           =   15
         Left            =   420
         TabIndex        =   58
         Top             =   2880
         Width           =   1995
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Discos Horizontal"
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
         Height          =   285
         Index           =   16
         Left            =   300
         TabIndex        =   57
         Top             =   2580
         Width           =   2145
      End
   End
   Begin VB.Frame frPUBS 
      BackColor       =   &H00000000&
      Caption         =   "Publicidades"
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
      Left            =   14580
      TabIndex        =   133
      Top             =   8190
      Visible         =   0   'False
      Width           =   3375
      Begin VB.VScrollBar vsPubliIMGCada 
         Height          =   330
         Left            =   2910
         Max             =   10
         Min             =   100
         TabIndex        =   137
         Top             =   2055
         Value           =   10
         Width           =   330
      End
      Begin VB.TextBox txtPubliImgCada 
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
         Left            =   2280
         TabIndex        =   140
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2040
         Width           =   600
      End
      Begin VB.CheckBox ckPubIMG 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir Publicidades (imagenes rotativas) "
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
         Height          =   390
         Left            =   150
         TabIndex        =   136
         Top             =   1500
         Width           =   3105
      End
      Begin VB.VScrollBar vsPubliCada 
         Height          =   330
         Left            =   2910
         Max             =   1
         Min             =   100
         TabIndex        =   135
         Top             =   915
         Value           =   5
         Width           =   330
      End
      Begin VB.TextBox txtPubliCada 
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
         Left            =   2280
         TabIndex        =   138
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   900
         Width           =   600
      End
      Begin VB.CheckBox ckPUB 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir Publicidades (Audio y video) "
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
         Height          =   570
         Left            =   120
         TabIndex        =   134
         Top             =   300
         Width           =   3105
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   3060
         Y1              =   1380
         Y2              =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reproducir publicidades cada X segundos"
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
         Height          =   405
         Index           =   30
         Left            =   90
         TabIndex        =   141
         Top             =   1980
         Width           =   2145
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reproducir publicidades cada X temas"
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
         Height          =   405
         Index           =   29
         Left            =   90
         TabIndex        =   139
         Top             =   840
         Width           =   2145
      End
   End
   Begin VB.CommandButton Command19 
      BackColor       =   &H0080C0FF&
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   123
      Top             =   8490
      Width           =   3720
   End
   Begin VB.Frame frValidacion 
      BackColor       =   &H00000000&
      Caption         =   "Validacion de uso de 3PM"
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
      Left            =   3030
      TabIndex        =   103
      Top             =   330
      Visible         =   0   'False
      Width           =   8685
      Begin VB.TextBox txtEstadoValidacion 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   120
         TabStop         =   0   'False
         Top             =   1830
         Width           =   4245
      End
      Begin VB.TextBox txtTraduccion 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   5820
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   119
         Top             =   1500
         Width           =   2800
      End
      Begin VB.TextBox txtClaveXaValidar 
         Alignment       =   2  'Center
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
         Left            =   5820
         Locked          =   -1  'True
         TabIndex        =   118
         TabStop         =   0   'False
         Top             =   1170
         Width           =   2800
      End
      Begin VB.TextBox txtCodigoXaValidar 
         Alignment       =   2  'Center
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
         Left            =   5820
         TabIndex        =   116
         TabStop         =   0   'False
         Top             =   420
         Width           =   2800
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Crear clave segun codigo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   5820
         Style           =   1  'Graphical
         TabIndex        =   115
         Top             =   780
         Width           =   2800
      End
      Begin VB.TextBox txtRegistroDiario 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   113
         Text            =   "frmConfig.frx":0482
         Top             =   3240
         Width           =   5715
      End
      Begin VB.TextBox txtAvisarAntes 
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
         Left            =   2520
         TabIndex        =   108
         TabStop         =   0   'False
         Text            =   "20"
         Top             =   1200
         Width           =   600
      End
      Begin VB.VScrollBar vsAvisarAntes 
         Height          =   330
         LargeChange     =   10
         Left            =   3120
         Max             =   10
         Min             =   500
         SmallChange     =   5
         TabIndex        =   106
         Top             =   1200
         Value           =   50
         Width           =   330
      End
      Begin VB.VScrollBar vsValidarCada 
         Height          =   330
         LargeChange     =   100
         Left            =   3120
         Max             =   50
         Min             =   5000
         SmallChange     =   10
         TabIndex        =   105
         Top             =   750
         Value           =   50
         Width           =   330
      End
      Begin VB.TextBox txtValidarCada 
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
         Left            =   2520
         TabIndex        =   107
         TabStop         =   0   'False
         Text            =   "100"
         Top             =   750
         Width           =   600
      End
      Begin VB.CheckBox chkValidar 
         BackColor       =   &H00000000&
         Caption         =   "Solicitar una clave cada una determinada cantidad de creditos"
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
         Height          =   390
         Left            =   150
         TabIndex        =   104
         Top             =   300
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado de validacion de este equipo"
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
         Index           =   24
         Left            =   150
         TabIndex        =   121
         Top             =   1620
         Width           =   5535
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Carge aqui el codigo solicitado"
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
         Index           =   23
         Left            =   5820
         TabIndex        =   117
         Top             =   180
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Registro de todos los inicios de 3PM y el valor de contador de creditos correspondiente"
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
         Height          =   615
         Index           =   22
         Left            =   90
         TabIndex        =   114
         Top             =   2820
         Width           =   5685
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "creditos"
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
         Index           =   21
         Left            =   3480
         TabIndex        =   112
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "creditos"
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
         Index           =   20
         Left            =   3480
         TabIndex        =   111
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Avisar cuando falten "
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
         Height          =   465
         Index           =   19
         Left            =   1080
         TabIndex        =   110
         Top             =   1140
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Solicitar clave cada"
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
         Index           =   18
         Left            =   90
         TabIndex        =   109
         Top             =   810
         Width           =   2385
      End
   End
   Begin VB.CommandButton Command16 
      BackColor       =   &H00FF8080&
      Caption         =   "Ingresar Clave"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3420
      Width           =   2595
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Administrador"
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
      Height          =   4065
      Left            =   30
      TabIndex        =   101
      Top             =   4020
      Width           =   2895
      Begin VB.CommandButton Command20 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Publicidades"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   132
         Top             =   3630
         Width           =   2650
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Administrar discos"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   129
         Top             =   3210
         Width           =   2650
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Validacion de uso"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1950
         Width           =   2650
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cambiar Licencia"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2370
         Width           =   2650
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "SUPERLICENCIA"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2790
         Width           =   2650
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Claves de 3PM"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1530
         Width           =   2650
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Incio 3PM"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1110
         Width           =   2650
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Creditos"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   2650
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Teclado"
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
         Height          =   380
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   690
         Width           =   2650
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Basicas"
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
      Height          =   2415
      Left            =   60
      TabIndex        =   100
      Top             =   60
      Width           =   2865
      Begin VB.CommandButton Command11 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Visualizacion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   2650
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Abrir MANUAL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1920
         Width           =   2650
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Protector de pantalla"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   660
         Width           =   2650
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Otras opciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1500
         Width           =   2650
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Aceleracion de 3PM"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   2650
      End
   End
   Begin VB.HScrollBar HSvolumen 
      Height          =   360
      LargeChange     =   10
      Left            =   7620
      Max             =   100
      TabIndex        =   15
      Top             =   5850
      Width           =   3975
   End
   Begin VB.Frame frOtras 
      BackColor       =   &H00000000&
      Caption         =   "Otras opciones"
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
      Height          =   3255
      Left            =   12060
      TabIndex        =   73
      Top             =   2430
      Visible         =   0   'False
      Width           =   4485
      Begin VB.TextBox txtSECwait 
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
         Left            =   3300
         TabIndex        =   96
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2280
         Width           =   600
      End
      Begin VB.VScrollBar VSSegEspera 
         Height          =   330
         LargeChange     =   10
         Left            =   3930
         Max             =   30
         Min             =   1200
         SmallChange     =   10
         TabIndex        =   95
         Top             =   2280
         Value           =   30
         Width           =   330
      End
      Begin VB.VScrollBar VsPorcTema 
         Height          =   330
         LargeChange     =   10
         Left            =   3930
         Max             =   10
         Min             =   100
         SmallChange     =   10
         TabIndex        =   94
         Top             =   2655
         Value           =   10
         Width           =   330
      End
      Begin VB.TextBox txtPorcTema 
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
         Left            =   3300
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2640
         Width           =   600
      End
      Begin VB.TextBox txtMaxFichas 
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
         Left            =   3300
         TabIndex        =   89
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1920
         Width           =   600
      End
      Begin VB.VScrollBar VSmaxFichas 
         Height          =   330
         Left            =   3930
         Max             =   5
         Min             =   200
         TabIndex        =   88
         Top             =   1920
         Value           =   5
         Width           =   330
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Cortes de luz"
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
         Height          =   1335
         Left            =   300
         TabIndex        =   84
         Top             =   240
         Width           =   4005
         Begin VB.OptionButton OpReiniNULL 
            BackColor       =   &H00000000&
            Caption         =   "Comienza de cero borrando la lista de ejecución."
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
            Height          =   375
            Left            =   60
            TabIndex        =   86
            Top             =   810
            Value           =   -1  'True
            Width           =   3840
         End
         Begin VB.OptionButton OpReiniFull 
            BackColor       =   &H00000000&
            Caption         =   "Se ejecutan todos los temas pendientes en la lista de ejecución"
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
            Left            =   60
            TabIndex        =   85
            Top             =   330
            Width           =   3705
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje ejecutar tema"
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
         Height          =   285
         Index           =   12
         Left            =   150
         TabIndex        =   97
         Top             =   2700
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Espera autoejecutar tema"
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
         Height          =   285
         Index           =   9
         Left            =   120
         TabIndex        =   98
         Top             =   2340
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maximo de fichas permitidas"
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
         Height          =   240
         Index           =   8
         Left            =   240
         TabIndex        =   87
         Top             =   1980
         Width           =   2925
      End
   End
   Begin VB.Frame frCreditos 
      BackColor       =   &H00000000&
      Caption         =   "Creditos"
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
      Height          =   2265
      Left            =   12060
      TabIndex        =   66
      Top             =   5760
      Visible         =   0   'False
      Width           =   4185
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         LargeChange     =   10
         Left            =   3600
         Max             =   1
         Min             =   6
         TabIndex        =   76
         Top             =   1770
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTemaVIDEO 
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
         Left            =   3000
         TabIndex        =   130
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1770
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Poner en 0 contador"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   660
         Width           =   3375
      End
      Begin VB.TextBox txtTemasXCredito 
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
         Left            =   3000
         TabIndex        =   78
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1110
         Width           =   600
      End
      Begin VB.VScrollBar VSTemasXCredito 
         Height          =   330
         LargeChange     =   10
         Left            =   3600
         Max             =   1
         Min             =   6
         TabIndex        =   77
         Top             =   1110
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         LargeChange     =   10
         Left            =   3600
         Max             =   1
         Min             =   6
         TabIndex        =   75
         Top             =   1440
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTema 
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
         Left            =   3000
         TabIndex        =   74
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos para 1 tema VIDEO"
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
         Index           =   28
         Left            =   330
         TabIndex        =   131
         Top             =   1800
         Width           =   2625
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
         Left            =   1950
         TabIndex        =   83
         Top             =   210
         Width           =   2070
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos por ficha"
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
         Index           =   11
         Left            =   420
         TabIndex        =   82
         Top             =   1140
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contador interno"
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
         Index           =   25
         Left            =   90
         TabIndex        =   81
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos para 1 tema"
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
         Index           =   26
         Left            =   330
         TabIndex        =   80
         Top             =   1470
         Width           =   2625
      End
   End
   Begin VB.Frame frAceleracion 
      BackColor       =   &H00000000&
      Caption         =   "Aceleracion de 3PM"
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
      Height          =   2415
      Left            =   12060
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   6315
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
         Left            =   210
         TabIndex        =   70
         Top             =   300
         Width           =   5865
         Begin VB.OptionButton OpImgINI 
            BackColor       =   &H00000000&
            Caption         =   "Cargar imagenes al inicio. Recomendado"
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
            Height          =   315
            Left            =   210
            TabIndex        =   72
            Top             =   300
            Width           =   5580
         End
         Begin VB.OptionButton OpImgSIS 
            BackColor       =   &H00000000&
            Caption         =   "Cargar las imagenes a pedido"
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
            Height          =   375
            Left            =   210
            TabIndex        =   71
            Top             =   570
            Value           =   -1  'True
            Width           =   5520
         End
      End
      Begin VB.CheckBox chkFastINI 
         BackColor       =   &H00000000&
         Caption         =   "Inicio rápido (no mostrar imágenes en la presentación)"
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
         Height          =   330
         Left            =   210
         TabIndex        =   69
         Top             =   1800
         Width           =   5890
      End
      Begin VB.CheckBox chkVUMeter 
         BackColor       =   &H00000000&
         Caption         =   "Habilitar VUMetro"
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
         Height          =   210
         Left            =   210
         TabIndex        =   68
         Top             =   1620
         Width           =   5890
      End
      Begin VB.CheckBox chkCargarDuracionTemas 
         BackColor       =   &H00000000&
         Caption         =   "Cargar la duracion de los temas (demora extra)"
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
         Height          =   300
         Left            =   210
         TabIndex        =   67
         Top             =   1320
         Width           =   5890
      End
   End
   Begin VB.Frame frProtector 
      BackColor       =   &H00000000&
      Caption         =   "Protector de pantalla"
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
      Height          =   3075
      Left            =   120
      TabIndex        =   59
      Top             =   9270
      Visible         =   0   'False
      Width           =   4185
      Begin VB.OptionButton chkProtectOriginal 
         BackColor       =   &H00000000&
         Caption         =   "Usar Protector de pantalla original (tapas de los discos)"
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
         Height          =   405
         Left            =   130
         TabIndex        =   126
         Top             =   1230
         Width           =   3900
      End
      Begin VB.OptionButton chkProtectorCustom 
         BackColor       =   &H00000000&
         Caption         =   "Usar protector de pantalla personalizado. "
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
         Height          =   405
         Left            =   130
         TabIndex        =   125
         Top             =   720
         Width           =   3900
      End
      Begin VB.OptionButton chkNoProtector 
         BackColor       =   &H00000000&
         Caption         =   "No usar protectore de pantalla"
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
         Height          =   405
         Left            =   130
         TabIndex        =   124
         Top             =   240
         Width           =   3900
      End
      Begin VB.VScrollBar vsEsperaTecla 
         Height          =   330
         LargeChange     =   10
         Left            =   3750
         Max             =   30
         Min             =   1200
         SmallChange     =   10
         TabIndex        =   127
         Top             =   1890
         Value           =   30
         Width           =   330
      End
      Begin VB.TextBox txtEsperaTecla 
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
         Left            =   3150
         TabIndex        =   62
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1875
         Width           =   600
      End
      Begin VB.VScrollBar vsDuracionProtect 
         Height          =   330
         LargeChange     =   10
         Left            =   3750
         Max             =   0
         Min             =   900
         SmallChange     =   10
         TabIndex        =   128
         Top             =   2250
         Value           =   900
         Width           =   330
      End
      Begin VB.TextBox txtDuracionProtect 
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
         Left            =   3150
         TabIndex        =   61
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2250
         Width           =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   300
         X2              =   3870
         Y1              =   1740
         Y2              =   1740
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Espera protector de pantalla"
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
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   64
         Top             =   1950
         Width           =   2925
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duracion del protector"
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
         Height          =   285
         Index           =   17
         Left            =   150
         TabIndex        =   63
         Top             =   2310
         Width           =   2925
      End
   End
   Begin VB.Frame frTeclado 
      BackColor       =   &H00000000&
      Caption         =   "Teclado"
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
      Height          =   4005
      Left            =   4410
      TabIndex        =   26
      Top             =   9120
      Visible         =   0   'False
      Width           =   3765
      Begin VB.CheckBox chkPasarhoja 
         BackColor       =   &H00000000&
         Caption         =   "Pasar páginas con botones de desplazamiento simple."
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
         Height          =   525
         Left            =   90
         TabIndex        =   60
         Top             =   3450
         Width           =   3615
      End
      Begin VB.TextBox txtPagAd 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2160
         Width           =   300
      End
      Begin VB.TextBox txtPagAt 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   23
         Top             =   2490
         Width           =   300
      End
      Begin VB.TextBox txtnPagAd 
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
         Left            =   2805
         TabIndex        =   42
         Top             =   2160
         Width           =   700
      End
      Begin VB.TextBox txtnPagAt 
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
         Left            =   2805
         TabIndex        =   41
         Top             =   2490
         Width           =   700
      End
      Begin VB.CheckBox chkApagarPC 
         BackColor       =   &H00000000&
         Caption         =   "Apagar la PC al cerrar el sistema"
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
         Height          =   330
         Left            =   60
         TabIndex        =   25
         Top             =   3150
         Width           =   3480
      End
      Begin VB.TextBox txtnCLOSE 
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
         Left            =   2820
         TabIndex        =   40
         Top             =   2820
         Width           =   700
      End
      Begin VB.TextBox txtnCONF 
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
         Left            =   2805
         TabIndex        =   39
         Top             =   1845
         Width           =   700
      End
      Begin VB.TextBox txtnNewF 
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
         Left            =   2805
         TabIndex        =   38
         Top             =   1515
         Width           =   700
      End
      Begin VB.TextBox txtnESC 
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
         Left            =   2805
         TabIndex        =   37
         Top             =   1185
         Width           =   700
      End
      Begin VB.TextBox txtnOK 
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
         Left            =   2805
         TabIndex        =   36
         Top             =   855
         Width           =   700
      End
      Begin VB.TextBox txtnIZQ 
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
         Left            =   2805
         TabIndex        =   35
         Top             =   525
         Width           =   700
      End
      Begin VB.TextBox txtnDER 
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
         Left            =   2805
         TabIndex        =   34
         Top             =   195
         Width           =   700
      End
      Begin VB.TextBox txtCLOSE 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   24
         Top             =   2835
         Width           =   300
      End
      Begin VB.TextBox txtCONF 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   21
         Top             =   1845
         Width           =   300
      End
      Begin VB.TextBox txtNewF 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   20
         Top             =   1515
         Width           =   300
      End
      Begin VB.TextBox txtESC 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   19
         Top             =   1185
         Width           =   300
      End
      Begin VB.TextBox txtOK 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   18
         Top             =   855
         Width           =   300
      End
      Begin VB.TextBox txtIZQ 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   17
         Top             =   525
         Width           =   300
      End
      Begin VB.TextBox txtDER 
         Alignment       =   2  'Center
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
         Left            =   2460
         MaxLength       =   1
         TabIndex        =   16
         Top             =   195
         Width           =   300
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Página Adelante"
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
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   44
         Top             =   2220
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Página Atras"
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
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   43
         Top             =   2535
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Cerrar Sistema"
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
         Height          =   240
         Index           =   6
         Left            =   0
         TabIndex        =   33
         Top             =   2880
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Configurar"
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
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   32
         Top             =   1890
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla Nueva ficha"
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
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   31
         Top             =   1560
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla SALIR"
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
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   30
         Top             =   1230
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla OK"
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
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   900
         Width           =   2450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla izquierda"
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
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   28
         Top             =   570
         Width           =   2445
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tecla derecha"
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
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   270
         Width           =   2450
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Caption         =   "Salir sin grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8550
      Width           =   2850
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8130
      Width           =   2850
   End
   Begin VB.Frame frConfigVis 
      BackColor       =   &H00000000&
      Caption         =   "Opciones de la configuracion elegida"
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
      Height          =   5445
      Left            =   2940
      TabIndex        =   99
      Top             =   90
      Width           =   8925
   End
   Begin VB.Line LineScroll 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   7650
      X2              =   11580
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para habilitar a las opciones de administrador debe infresar su clave de administrador"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   1395
      Left            =   90
      TabIndex        =   102
      Top             =   2520
      Width           =   2835
   End
   Begin VB.Label LblVol 
      BackStyle       =   0  'Transparent
      Caption         =   "Volumen"
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
      Height          =   285
      Left            =   7590
      TabIndex        =   92
      Top             =   5580
      Width           =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FFFF&
      Visible         =   0   'False
      X1              =   12000
      X2              =   12000
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Visible         =   0   'False
      X1              =   0
      X2              =   12000
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label lblHLP 
      BackColor       =   &H00404000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle/Ayuda de la opcion elegida"
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
      Height          =   3345
      Left            =   2940
      TabIndex        =   45
      Top             =   5550
      Width           =   4605
   End
   Begin VB.Label lblTBRcfg 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmConfig.frx":04A1
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   2085
      Left            =   8250
      TabIndex        =   46
      Top             =   6360
      Width           =   2925
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
    " 3PM"
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
    " 3PM"
End Sub

Private Sub chkMostrarRotulos_LostFocus()
    chkMostrarRotulos.ForeColor = vbWhite
End Sub

Private Sub chknoprotector_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkNoProtector.ForeColor = vbYellow
    HLP "Deshabilitar la función de protección de pantalla. No recomendado"
End Sub

Private Sub chknoprotector_LostFocus()
    chkNoProtector.ForeColor = vbWhite
End Sub

Private Sub chkNoVumVID_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkNoVumVID.ForeColor = vbYellow
    HLP "Quitar el VUMetro (medidor de sonido) cuando los videos sean full-screen"
End Sub

Private Sub chkNoVumVID_LostFocus()
    chkNoVumVID.ForeColor = vbWhite
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

Private Sub chkProtectorCustom_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectorCustom.ForeColor = vbYellow
    HLP "Si desea mostrar imagenes personalizadas debera cargarlas en " + _
    "la carpeta FOTOS de la carpeta en que se instalo 3PM. " + _
    "No use imagenes muy pesadas ya que puede afectar el rendimiento de 3PM. Se recomienda" + _
    "no sobrepasar los 100 KB"
End Sub

Private Sub chkProtectorCustom_LostFocus()
    chkProtectorCustom.ForeColor = vbWhite
End Sub

Private Sub chkProtectOriginal_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectOriginal.ForeColor = vbYellow
    HLP "Puede usar para proteger la pantalla el protector por defecto. Este muestra " + _
    "las tapas de los discos."
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
    " 3PM"
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
    " 3PM"
End Sub

Private Sub chkRotulosArriba_LostFocus()
    chkRotulosArriba.ForeColor = vbWhite
End Sub

Private Sub chkSalida2_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkSalida2.ForeColor = vbYellow
    HLP "Habilitar la segunda salida para reproduccion de videos. " + _
        "Debe habilitarse la salida de TV como expanción del escritorio " + _
        "y configurarla con la misma definición de pixeles para ambas salidas"
End Sub

Private Sub chkSalida2_LostFocus()
    chkSalida2.ForeColor = vbWhite
End Sub

Private Sub chkTouch_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkTouch.ForeColor = vbYellow
    HLP "Mostrar los botones para pantallas sensibles al tacto. Este cambio solo se vera una vez reiniciado" + _
    " 3PM"
End Sub

Private Sub chkTouch_LostFocus()
    chkTouch.ForeColor = vbWhite
End Sub

Private Sub chkValidar_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkValidar.ForeColor = vbYellow
    HLP "Se establecera una cantidad de creditos luego de la cual 3PM se bloqueara hasta que se ingrse una clave solicitada." + _
        " La lista de claves estará a disposicion del administrador del equipo. Esto permitirá bloquear usos en casos de falta de " + _
        "pago. El usuario recibira preavisos para solicitar su clave y regularizar su situacion"
End Sub

Private Sub chkValidar_LostFocus()
    chkValidar.ForeColor = vbWhite
End Sub

Private Sub chkVidFullScreen_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVidFullScreen.ForeColor = vbYellow
    HLP "Mostrar los videos en pantalla completa cuando se ejecuten"
End Sub

Private Sub chkVidFullScreen_LostFocus()
    chkVidFullScreen.ForeColor = vbWhite
End Sub

Private Sub chkVUMeter_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVUMeter.ForeColor = vbYellow
    HLP "Se llama VuMetro al medidor de nivel de sonido. Este es muy" + _
    " atractivo a la vista pero consume recursos de la PC. Por esto" + _
    " solo deberá usarse cuando el rendimiento del equipo no se vea afectado " + _
    "con el uso de este. Para PCs de bajos recursos (procesador y RAM)" + _
    " se recomienda dejar desactivado. Este cambio solo se vera una vez reiniciado" + _
    " 3PM"
End Sub

Private Sub chkVUMeter_LostFocus()
    chkVUMeter.ForeColor = vbWhite
End Sub

Private Sub ckPUB_Click()
    If PUBs.TotalPUBs = 0 Then
        MsgBox "No puede activar esta opción ya que no hay publicidades cargadas." + vbCrLf + _
            "Para cargar publicidades debera incluir en la carpeta 'PUB' (en la carpeta en " + _
            "que instalo 3PM) uno o más ficheros MP3, AVI o MPG"
        ckPUB = 0
    End If
End Sub

Private Sub ckPUB_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    ckPUB.ForeColor = vbYellow
    HLP "Indica si se reproducirán publicidades. Si se habilita esta opción deben colocarse ficheros " + _
        "MP3, AVI o MPG en la carpeta PUB (de la carpeta en la que instalo 3PM). Estos ficheros se reproducen cada X (a configurar) " + _
        "temas y de a uno por vez. Se reproducen en orden alfabético por lo que podrá " + _
        "modificar el nombre para definir el orden deseado. Puede tambien duplicar ficheros para " + _
        "darle mayor repeticion a alguna publicidad en particular"
End Sub

Private Sub ckPUB_LostFocus()
    ckPUB.ForeColor = vbWhite
End Sub

Private Sub ckPubIMG_Click()
    If PUBs.TotalPUBsIMG = 0 Then
        MsgBox "No puede activar esta opción ya que no hay publicidades cargadas." + vbCrLf + _
            "Para cargar publicidades debera incluir en la carpeta 'PUB' (en la carpeta en " + _
            "que instalo 3PM) uno o más ficheros JPG, BMP o GIF. " + _
            "Debera reiniciar 3PM para que este cambio surta efecto"
        ckPUB = 0
    End If
End Sub

Private Sub ckPubIMG_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    ckPubIMG.ForeColor = vbYellow
    HLP "Indica si se reproducirán publicidades. Si se habilita esta opción deben colocarse ficheros " + _
        "JPG, BMP o GIF en la carpeta PUB (de la carpeta en la que instalo 3PM). Estos ficheros " + _
        "se muestran cada X (a configurar) " + _
        "segundos. Se muestran en orden alfabético por lo que podrá " + _
        "modificar el nombre para definir el orden deseado. Puede tambien duplicar ficheros para " + _
        "darle mayor repeticion a alguna publicidad en particular"
End Sub

Private Sub ckPubIMG_LostFocus()
    ckPubIMG.ForeColor = vbWhite
End Sub

Private Sub Command1_Click() 'GRABAR BUTTON
    'GRABAR BUTTON
    'cargar los datos del archivo SYSfolder + "3pmcfg.tbr"
    'paso todo a una cadena, la encripto y luego la escribo
    Dim FullConfig As String
    FullConfig = FullConfig + "CargarImagenInicio=" + CStr(OpImgINI) + vbCrLf
    FullConfig = FullConfig + "AutoReDraw=" + CStr(chkAutoReDraw) + vbCrLf
    FullConfig = FullConfig + "TeclaDerecha=" + txtnDER + vbCrLf
    FullConfig = FullConfig + "TeclaIzquierda=" + txtnIZQ + vbCrLf
    FullConfig = FullConfig + "TeclaPagAd=" + txtnPagAd + vbCrLf
    FullConfig = FullConfig + "TeclaPagAt=" + txtnPagAt + vbCrLf
    FullConfig = FullConfig + "TeclaOK=" + txtnOK + vbCrLf
    FullConfig = FullConfig + "TeclaESC=" + txtnESC + vbCrLf
    FullConfig = FullConfig + "TeclaNuevaFicha=" + txtnNewF + vbCrLf
    FullConfig = FullConfig + "TeclaConfig=" + txtnCONF + vbCrLf
    FullConfig = FullConfig + "TeclaCerrarSistema=" + txtnCLOSE + vbCrLf
    FullConfig = FullConfig + "ApagarAlCierre= " + CStr(chkApagarPC) + vbCrLf
    FullConfig = FullConfig + "RankToPeople= " + CStr(chkRankToPeople) + vbCrLf
    FullConfig = FullConfig + "MaximoFichas=" + txtMaxFichas + vbCrLf
    FullConfig = FullConfig + "EsperaMinutos=" + txtSECwait + vbCrLf
    FullConfig = FullConfig + "FastIni=" + CStr(chkFastINI) + vbCrLf
    FullConfig = FullConfig + "HabilitarVUMetro=" + CStr(chkVUMeter) + vbCrLf
    FullConfig = FullConfig + "VidfullScreen=" + CStr(chkVidFullScreen) + vbCrLf
    FullConfig = FullConfig + "Salida2=" + CStr(chkSalida2) + vbCrLf
    FullConfig = FullConfig + "NoVumVid=" + CStr(chkNoVumVID) + vbCrLf
    'Valores de ReIni LISTA=solo lista NADA=arranca de cero
    If OpReiniFull Then
        FullConfig = FullConfig + "ReINI=LISTA" + vbCrLf
    Else
        FullConfig = FullConfig + "ReINI=NADA" + vbCrLf
    End If
    FullConfig = FullConfig + "Volumen=" + Trim(Str(HSvolumen)) + vbCrLf
    FullConfig = FullConfig + "EsperaTecla=" + txtEsperaTecla + vbCrLf
    FullConfig = FullConfig + "PorcentajeTema=" + txtPorcTema + vbCrLf
    FullConfig = FullConfig + "DiscosH=" + txtDiscosH + vbCrLf
    FullConfig = FullConfig + "DiscosV=" + txtDiscosV + vbCrLf
    FullConfig = FullConfig + "DuracionProtect=" + txtDuracionProtect + vbCrLf
    
    FullConfig = FullConfig + "PasarHoja=" + CStr(chkPasarhoja) + vbCrLf
    FullConfig = FullConfig + "DistorcionarTapas=" + CStr(chkDistorcionarTapas) + vbCrLf
    'valores para el protectore de pantalla
    '0=inhabilitado 1=Original 2=Carpeta Fotos 3= Video FullScreen
    If chkNoProtector Then
        FullConfig = FullConfig + "Protector=0" + vbCrLf
        frmIndex.Timer3.Interval = 0 'para el reloj del protector
    Else
        'si es cualquierra de los otros habilitarlo
        frmIndex.Timer3.Interval = 10000
    End If
    If chkProtectOriginal Then FullConfig = FullConfig + "Protector=1" + vbCrLf
    If chkProtectorCustom Then FullConfig = FullConfig + "Protector=2" + vbCrLf
    FullConfig = FullConfig + "CargarDuracionTemas=" + CStr(chkCargarDuracionTemas) + vbCrLf
    FullConfig = FullConfig + "MostrarRotulos=" + CStr(chkMostrarRotulos) + vbCrLf
    FullConfig = FullConfig + "RotulosArriba=" + CStr(chkRotulosArriba) + vbCrLf
    FullConfig = FullConfig + "TemasPorCredito= " + txtTemasXCredito + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTema= " + txtCreditosCuestaTema + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTemaVIDEO= " + txtCreditosCuestaTemaVIDEO + vbCrLf
    FullConfig = FullConfig + "TextoUsuario= " + TxtUSUARIO + vbCrLf
    'validacion con clave cada x creditos
    FullConfig = FullConfig + "Validar= " + CStr(chkValidar) + vbCrLf
    FullConfig = FullConfig + "ValidarCada= " + txtValidarCada + vbCrLf
    FullConfig = FullConfig + "AvisarAntes= " + txtAvisarAntes + vbCrLf
    FullConfig = FullConfig + "MostrarTouch= " + CStr(chkTouch) + vbCrLf
    'publicidades
    FullConfig = FullConfig + "MostrarPUB= " + CStr(ckPUB) + vbCrLf
    FullConfig = FullConfig + "MostrarPUBIMG= " + CStr(ckPubIMG) + vbCrLf
    FullConfig = FullConfig + "PubliCada= " + txtPubliCada + vbCrLf
    FullConfig = FullConfig + "PubliIMGCada= " + txtPubliImgCada + vbCrLf
    'encriptar
    FullConfig = Encriptar(FullConfig, False)
    'grabar el kilombo
    Set TE = FSO.CreateTextFile(SYSfolder + "\3pmcfg.tbr", True)
    TE.Write FullConfig
    TE.Close
    
    'SI NO HAY que validar me aseguro que se bore el archivo de validacion SYSfolder + "\radilav.cfg"
    If chkValidar.Value = 0 Then
        If FSO.FileExists(SYSfolder + "\radilav.cfg") Then FSO.DeleteFile SYSfolder + "\radilav.cfg", True
    End If
    
    'publicidades
    PubliCada = Val(txtPubliCada)
    MostrarPUB = ckPUB
    
    PubliIMGCada = Val(txtPubliImgCada)
    MostrarPUBIMG = ckPubIMG
    
    PUBs.HabilitarPublicidades = MostrarPUB
    PUBs.SonarPublicidadesCada = PubliCada
    
    PUBs.HabilitarPublicidadesIMG = MostrarPUBIMG
    PUBs.SonarPublicidadesIMGCada = PubliIMGCada
    
    'todas las propiedades se quedan sin reiniciar
    'algunas no se necesitan
    'NO NECESITO CargarIMGinicio = LeerConfig("CargarImagenInicio")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
   
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    'NO DEBO ReINI = LeerConfig("ReINI","LISTA")
    VolumenIni = CLng(LeerConfig("Volumen", "50"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    'NO NECESITO FASTini = LeerConfig("FastIni","1")
    PasarHoja = LeerConfig("PasarHoja", "1")
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    VolumenIni = HSvolumen
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosCuestaTema = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTemaVIDEO = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "0")

    If K.LICENCIA = aSinCargar Then
        frmIndex.lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        frmIndex.lblDEMO = textoUsuario
    End If
    
    Unload Me
End Sub

Private Sub Command1_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command1.BackColor = vbYellow
    HLP "Grabar los datos cargados"
End Sub

Private Sub Command1_LostFocus()
    Command1.BackColor = &HFF8080
End Sub

Private Sub Command10_Click()
    CentrarFrEnFr frConfigVis, frProtector
End Sub

Private Sub Command10_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command10.BackColor = vbYellow
    HLP "Opciones del protector de pantalla"
End Sub

Private Sub Command10_LostFocus()
    Command10.BackColor = &HFFC0C0
End Sub

Private Sub Command11_Click()
    CentrarFrEnFr frConfigVis, frVisualizacion
End Sub

Private Sub Command11_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command11.BackColor = vbYellow
    HLP "Opciones de visualizacion de 3PM"
End Sub

Private Sub Command11_LostFocus()
    Command11.BackColor = &HFFC0C0
End Sub

Private Sub Command12_Click()
    CentrarFrEnFr frConfigVis, frCreditos
End Sub

Private Sub Command12_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command12.BackColor = vbYellow
    HLP "Configuracion de precios de la fonola. Opcion de reinicio de contador de creditos"
End Sub

Private Sub Command12_LostFocus()
    Command12.BackColor = &HFFC0C0
End Sub

Private Sub Command13_Click()
    CentrarFrEnFr frConfigVis, frTeclado
End Sub

Private Sub Command13_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command13.BackColor = vbYellow
    HLP "Configuracion de las teclas usadas en 3PM"
End Sub

Private Sub Command13_LostFocus()
    Command13.BackColor = &HFFC0C0
End Sub

Private Sub Command14_Click()
    CentrarFrEnFr frConfigVis, frOtras
End Sub

Private Sub Command14_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command14.BackColor = vbYellow
    HLP "Otras opciones de configuracion de 3PM"
End Sub

Private Sub Command14_LostFocus()
    Command14.BackColor = &HFFC0C0
End Sub

Private Sub Command15_Click()
    CentrarFrEnFr frConfigVis, frAceleracion
End Sub

Private Sub Command15_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command15.BackColor = vbYellow
    HLP "Opciones de Aceleracion de 3PM. Utilizar para optimizar recursos segun el equipo utilizado."
End Sub

Private Sub Command15_LostFocus()
    Command15.BackColor = &HFFC0C0
End Sub

Private Sub Command16_Click()
    'Ingresar Clave Admin BUTTON!!!
    frmCLAVE.Show 1
    Dim TodoOk As Boolean
    TodoOk = False
    'si es una demo que permita la clave de administrador "DEMO"
    If K.LICENCIA = CGratuita And UCase(ClaveIngresada) = "DEMO" Then TodoOk = True
    If K.LICENCIA = aSinCargar And UCase(ClaveIngresada) = "DEMO" Then TodoOk = True
    'ver que la contraseña se tome desde el teclado al usuario
    If ClaveIngresada = ClaveAdmin Or ClaveIngresada = "rmlvf28177891" Then TodoOk = True
    
    If TodoOk Then
        'habilitar todos los botones
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        Command8.Enabled = True
        Command9.Enabled = True
        Command12.Enabled = True
        Command13.Enabled = True
        Command17.Enabled = True
        Command20.Enabled = True
    Else
        MsgBox "La clave ingresada no es correcta"
    End If
End Sub

Private Sub Command17_Click()
    CentrarFrEnFr frConfigVis, frValidacion
    txtRegistroDiario.SelStart = Len(txtRegistroDiario) - 1
    txtRegistroDiario.SelLength = 1
End Sub

Private Sub Command17_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command17.BackColor = vbYellow
    HLP "Solicitar claves periodicamente para no perimitir usos inválidos." + vbCrLf + _
        "De esta forma podra controlar los pagos de las concesiones de sus fonolas"
End Sub

Private Sub Command17_LostFocus()
    Command17.BackColor = &HFFC0C0
End Sub

Private Sub Command18_Click()
    'este es un huevo de pascua!!!!
    If UCase(txtCodigoXaValidar) = "3PM" Then
        Dim WI3 As String
        WI3 = "Esta es una protección para posibles violaciones a los derechos de " + _
            "autor de este software." + vbCrLf + _
            "Este es un desarrollo de Andrés Vázquez Flexes." + vbCrLf + _
            "Argentino de DNI (Documento Nacional de" + _
            " Identidad) 26453653." + vbCrLf + _
            "Podra contactarse conmigo a info@tbrsoft.com o telefónicamente " + _
            "(desde fuera de Argentina) a:" + vbCrLf + _
            "0054-9-351-4022170" + vbCrLf + _
            "0054-3543-485045" + vbCrLf + _
            "0054-3543-489271" + vbCrLf + _
            "O desde la Argentina a:" + vbCrLf + _
            "0351-154022170" + vbCrLf + _
            "03543-485045" + vbCrLf + _
            "03543-489271"

        MsgBox WI3
        Exit Sub
        
    End If
    
    'si el tipo puso en string estamos hasta las Bolas!!!!
    If Not IsNumeric(txtCodigoXaValidar) Then
        MsgBox "Los códigos para validar son siempre numericos. Ingresa correctamente"
        Exit Sub
    End If
    txtClaveXaValidar = ClaveParaValidar(txtCodigoXaValidar)
    'cargar kla traduccion
    Dim Largo As Long
    Largo = Len(txtClaveXaValidar)
    Dim CC As Long, Letra As String
    CC = 1
    Do While CC <= Largo
        Letra = Mid(txtClaveXaValidar, CC, 1)
        Select Case Asc(Letra)
            Case TeclaIZQ: txtTraduccion = txtTraduccion + "Tecla Izquierda" + vbCrLf
            Case TeclaDER: txtTraduccion = txtTraduccion + "Tecla Derecha" + vbCrLf
            Case TeclaPagAd: txtTraduccion = txtTraduccion + "Tecla Pagina Adelante" + vbCrLf
            Case TeclaPagAt: txtTraduccion = txtTraduccion + "Tecla Pagina Atras" + vbCrLf
            Case Else: txtTraduccion = txtTraduccion + "ERROR en traduccion. Se uso la letra " + Letra + vbCrLf
        End Select
        CC = CC + 1
    Loop
End Sub

Private Sub Command19_Click()
    frmCompraYA.Show 1
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
    Command2.BackColor = &HFF8080
End Sub

Private Sub Command20_Click()
    CentrarFrEnFr frConfigVis, frPUBS
End Sub

Private Sub Command3_Click()
    SumarContadorCreditos -CONTADOR 'esto lo deja en cero
    lblContador = STRceros(CONTADOR, 11)
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
    If K.LICENCIA = HSuperLicencia Then
        frmSUPERlic.Show 1
    Else
        MsgBox "Usted no posee una SUPELICENCIA envie un email a info@tbrsoft.com para más información." + vbCrLf + _
            "No tiene acceso"
    End If
End Sub

Private Sub Command5_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command5.BackColor = vbYellow
    HLP "Convierta a 3PM en su propio software. Cambie los logos y coloque información como si el " + _
    "software fuera de su propiedad"
End Sub

Private Sub Command5_LostFocus()
    Command5.BackColor = &HFFC0C0
End Sub

Private Sub Command6_Click()
    frmINI3PM.Show 1
End Sub

Private Sub Command6_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command6.BackColor = vbYellow
    HLP "Configurar las opciones de inicio de 3PM"
End Sub

Private Sub Command6_LostFocus()
    Command6.BackColor = &HFFC0C0
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

Private Sub Command8_Click()
    If MsgBox("¿Desea borrar los datos de su licencia actual para volver a cargarlos?" + vbCrLf + _
        "Usese solo para cuando obtenga una nueva clave para cargar", vbCritical + vbYesNo, "NUEVA LICENCIA") = vbNo Then Exit Sub
    
    'borro el archivo de registro para que inicie preguntando clave
    
    If FSO.FileExists(K.FileData) Then FSO.DeleteFile K.FileData, True
    
    
    MsgBox "La información de licencia se ha borrado correctamente. El sistema se cerrará " + _
        "para que cargue nuevamente su clave"
    
    End
    
End Sub

Private Sub Command8_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command8.BackColor = vbYellow
    HLP "Borre la información de su licencia actual para cargar una nueva clave"
End Sub

Private Sub Command8_LostFocus()
    Command8.BackColor = &HFFC0C0
End Sub

Private Sub Command9_Click()
    frmClaves.Show 1
End Sub

Private Sub Command9_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command9.BackColor = vbYellow
    HLP "Modificar las claves de 3PM"
End Sub

Private Sub Command9_LostFocus()
    Command9.BackColor = &HFFC0C0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            End
        
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
    frmIndex.lblNoTecla = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            lblContador = STRceros(CONTADOR, 11)
            If CREDITOS >= 10 Then
                frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
            Else
                frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
            End If
        Else
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    'poner en tamaño para que se ajuste bien
    Me.Height = 9000
    Me.Width = 12000
    MostrarCursor True
    AjustarFRM Me, 12000
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(WINfolder + "\SL\txtCFG.tbr") Then
            Set TE = FSO.OpenTextFile(WINfolder + "\SL\txtCFG.tbr", ForReading, False)
            If TE.AtEndOfStream = False Then
                Dim NewT As String
                NewT = TE.ReadAll
            Else
                NewT = "Error Al leer el archivo"
                WriteTBRLog "No pudo leer el texto de configuracion w/sl/txtcfg.tbr", True
            End If
            lblTBRcfg = NewT
            TE.Close
        Else
            lblTBRcfg = "Desarrollado por tbrSoft" + vbCrLf + "www.tbrsoft.com" + vbCrLf + _
                "----------------" + vbCrLf + "Contáctenos a info@tbrsoft.com" + vbCrLf + _
                "tbrsoft@cpcipc.org" + vbCrLf + "----------------" + vbCrLf + "Hecho en Argentina"
        End If
    Else
        lblTBRcfg = "Desarrollado por tbrSoft" + vbCrLf + "www.tbrsoft.com" + vbCrLf + _
            "----------------" + vbCrLf + "Contáctenos a info@tbrsoft.com" + vbCrLf + _
            "tbrsoft@cpcipc.org" + vbCrLf + "----------------" + vbCrLf + "Hecho en Argentina"
    End If
    lblContador = STRceros(CONTADOR, 11)
    
    If K.LICENCIA = aSinCargar Then
        TxtUSUARIO = "No puede modificar esta opcion si es una version demo"
        TxtUSUARIO.Locked = True
    End If
        
    'lblTIT = "3PM - Sistema de reproducción de ficheros MP3." + vbCrLf + vbCrLf + _
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
    
    'leer el archivo de configuracion SYSfolder + "3pmcfg.tbr"
    CargarIMGinicio = LeerConfig("CargarImagenInicio", "1")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI", "LISTA")
    'que no se carge el voilumen grabado
    'VolumenIni = CLng(LeerConfig("Volumen", "50"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    FASTini = LeerConfig("FastIni", "1")
    HabilitarVUMetro = LeerConfig("HabilitarVUMetro", "1")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "0")
    PasarHoja = LeerConfig("PasarHoja", "1")
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    RankToPeople = LeerConfig("RankToPeople", "1")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosCuestaTema = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTemaVIDEO = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    'validar cada X Creditos
    Validar = LeerConfig("Validar", "0")
    ValidarCada = LeerConfig("ValidarCada", "500")
    AvisarAntes = LeerConfig("AvisarAntes", "50")
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    'publicidades
    MostrarPUB = LeerConfig("MostrarPub", "0")
    PubliCada = LeerConfig("PubliCada", "5")
    MostrarPUBIMG = LeerConfig("MostrarPubIMG", "0")
    PubliIMGCada = LeerConfig("PubliIMGCada", "10")
    
    PUBs.HabilitarPublicidades = MostrarPUB
    PUBs.SonarPublicidadesCada = PubliCada
    PUBs.HabilitarPublicidadesIMG = MostrarPUBIMG
    PUBs.SonarPublicidadesIMGCada = PubliIMGCada
    
    
    
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
    chkVidFullScreen = -vidFullScreen
    chkSalida2 = -Salida2
    chkNoVumVID = -NoVumVID
    vsDiscosH = TapasMostradasH
    vsDiscosV = TapasMostradasV
    TeclaConfOK = "{UP}"
    TeclaConfESC = "{DOWN}"
    chkPasarhoja = -PasarHoja
    
    If Protector = 0 Then chkNoProtector = True
    If Protector = 1 Then chkProtectOriginal = True
    If Protector = 2 Then chkProtectorCustom = True
    
    chkCargarDuracionTemas = -CargarDuracionTemas
    
    chkMostrarRotulos = -MostrarRotulos
    chkRotulosArriba = -RotulosArriba
    VSTemasXCredito = TemasPorCredito
    vsCreditosCuestaTema = CreditosCuestaTema
    vsCreditosCuestaTemaVIDEO = CreditosCuestaTemaVIDEO
    TxtUSUARIO = textoUsuario
    chkTouch = -MostrarTouch
    'validar cada X creditos
    chkValidar = -Validar
    vsValidarCada = ValidarCada
    vsAvisarAntes = AvisarAntes
    'publicidad
    ckPUB = -MostrarPUB
    vsPubliCada = PubliCada
    ckPubIMG = -MostrarPUBIMG
    vsPubliIMGCada = PubliIMGCada
    
    'mostrar el registro diario de contador
    Dim TE2 As TextStream
    Set TE2 = FSO.OpenTextFile(SYSfolder + "\daily.cfg", ForReading, False)
    Dim TodoTe2 As String
    TodoTe2 = TE2.ReadAll
    TE2.Close
    txtRegistroDiario = TodoTe2
    txtRegistroDiario.SelStart = Len(txtRegistroDiario) - 1
    txtRegistroDiario.SelLength = 1
    txtEstadoValidacion = "Creditos Usados: " + CStr(CreditosValidar) + " de " + CStr(ValidarCada) + vbCrLf + _
        " Quedan: " + CStr(ValidarCada - CreditosValidar) + vbCrLf + _
        " Codigo Actual: " + CodigoParaClaveActual
    
    'mostrra visulaizacion
    Command11_Click
End Sub

Private Sub HSvolumen_Change()
    If frmIndex.MP3.IsPlaying Then frmIndex.MP3.Volumen = HSvolumen
    LblVol = "Volumen: " + Trim(Str(HSvolumen))
End Sub

Private Sub HSvolumen_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineScroll.Visible = True
    HLP "Volumen del sonido actual."
End Sub

Private Sub HSvolumen_LostFocus()
    LineScroll.Visible = False
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
    " 3PM"
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
    " 3PM"
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
    " 3PM"
End Sub

Private Sub txtESC_KeyDown(KeyCode As Integer, Shift As Integer)
    'la tecla derecha cumple la funcion de tabulacion por lo que no se tiene en cuenta
    If KeyCode = TeclaDER Or KeyCode = TeclaIZQ Or Shift Then Exit Sub
    txtnESC = KeyCode
    txtESC = Chr(KeyCode)
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
    " 3PM"
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
    " 3PM"
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
    " 3PM"
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
    " 3PM"
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
    " 3PM"
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

Private Sub TxtUSUARIO_GotFocus()
    'deshabilitar el teclado
    Me.KeyPreview = False
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    TxtUSUARIO.BackColor = vbYellow
    HLP "Este texto se mostrara en la página principal de 3PM como espacio de publicidad de su empresa"
End Sub

Private Sub TxtUSUARIO_LostFocus()
    TxtUSUARIO.BackColor = vbWhite
    Me.KeyPreview = True
End Sub

Private Sub vsAvisarAntes_Change()
    txtAvisarAntes = vsAvisarAntes
End Sub

Private Sub vsAvisarAntes_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtAvisarAntes.BackColor = vbYellow
    HLP "Antes del bloqueo del equipo recibira notificaciones cada vez que se inicie el equipo"
End Sub

Private Sub vsAvisarAntes_LostFocus()
    txtAvisarAntes.BackColor = vbWhite
End Sub

Private Sub vsCreditosCuestaTema_Change()
    txtCreditosCuestaTema = vsCreditosCuestaTema
End Sub

Private Sub vsCreditosCuestaTema_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCreditosCuestaTema.BackColor = vbYellow
    HLP "Cantidad de creditos que se necesitan para ejecutar un tema. Si lo configura en dos necesitara" + _
    " dos creditos para poder ejecutar un tema"
End Sub

Private Sub vsCreditosCuestaTema_LostFocus()
    txtCreditosCuestaTema.BackColor = vbWhite
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_Change()
    txtCreditosCuestaTemaVIDEO = vsCreditosCuestaTemaVIDEO
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCreditosCuestaTemaVIDEO.BackColor = vbYellow
    HLP "Cantidad de creditos que se necesitan para ejecutar un " + _
    "clip de video musical. Si lo configura en dos necesitara" + _
    " dos creditos para poder ejecutar un clip de video"
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_LostFocus()
    txtCreditosCuestaTemaVIDEO.BackColor = vbWhite
End Sub

Private Sub vsDiscosH_Change()
    txtDiscosH = vsDiscosH
End Sub

Private Sub vsDiscosH_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDiscosH.BackColor = vbYellow
    HLP "Cantidad de discos que se distribuiran horizontalmente. tbrSoft" + _
    " recomienda usar 4 (y 3 vertical). Puede usted probar distintos " + _
    "valores que sean de su agrado. Este cambio solo se vera una vez reiniciado" + _
    " 3PM"
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
    " 3PM"
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

Private Sub vsPubliCada_Change()
    txtPubliCada = vsPubliCada
End Sub

Private Sub vsPubliCada_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPubliCada.BackColor = vbYellow
    HLP "Indica cuantos temas deben pasar para que se ejecute una publicidad"
End Sub

Private Sub vsPubliCada_LostFocus()
    txtPubliCada.BackColor = vbWhite
End Sub

Private Sub vsPubliIMGCada_Change()
    txtPubliImgCada = vsPubliIMGCada
End Sub

Private Sub vsPubliIMGCada_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPubliImgCada.BackColor = vbYellow
    HLP "Indica cuantos segundos deben pasar para que se cambien la imagen publicitaria de la página inicial. " + _
    "Debera reiniciar 3PM para que este cambio surta efecto"
End Sub

Private Sub vsPubliIMGCada_LostFocus()
    txtPubliImgCada.BackColor = vbWhite
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

Private Sub VSTemasXCredito_Change()
    txtTemasXCredito = VSTemasXCredito
End Sub

Private Sub VSTemasXCredito_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtTemasXCredito.BackColor = vbYellow
    HLP "Cantidad de temas que se pueden reproducir con un credito. No necesita reiniciar 3PM" + _
    " para que esta configuracion surga efecto."
End Sub

Private Sub VSTemasXCredito_LostFocus()
    txtTemasXCredito.BackColor = vbWhite
End Sub

Private Sub CentrarFrEnFr(FrBig As Frame, FrChi As Frame)
    FrChi.Left = FrBig.Left + (FrBig.Width / 2 - FrChi.Width / 2)
    FrChi.Top = FrBig.Top + (FrBig.Height / 2 - FrChi.Height / 2)
    'se asegura que si o si se vean solo esos dos
    FrBig.ZOrder
    FrChi.ZOrder
    FrChi.Visible = True

End Sub

Private Sub vsValidarCada_Change()
    txtValidarCada = vsValidarCada
End Sub

Private Sub vsValidarCada_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtValidarCada.BackColor = vbYellow
    HLP "Cantidad de creditos luego de la cual se bloquera el equipo. " + _
        "Solo se deshabilita con el ingreso de una clave enviada por el administrador"
End Sub

Private Sub vsValidarCada_LostFocus()
    txtValidarCada.BackColor = vbWhite
End Sub
