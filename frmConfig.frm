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
      Height          =   5295
      Left            =   13230
      TabIndex        =   42
      Top             =   10050
      Visible         =   0   'False
      Width           =   4485
      Begin VB.CommandButton Command28 
         Caption         =   "Especiales monedero"
         Height          =   465
         Left            =   3210
         TabIndex        =   187
         Top             =   4380
         Width           =   1125
      End
      Begin VB.CheckBox chkCS 
         BackColor       =   &H00000000&
         Caption         =   "Activar correcion de señales"
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
         Left            =   360
         TabIndex        =   186
         Top             =   4440
         Width           =   2760
      End
      Begin VB.VScrollBar vsSegFade 
         Height          =   330
         Left            =   3960
         Max             =   3
         Min             =   20
         TabIndex        =   183
         Top             =   3855
         Value           =   10
         Width           =   330
      End
      Begin VB.TextBox txtSegFade 
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
         Left            =   3330
         TabIndex        =   182
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3870
         Width           =   600
      End
      Begin VB.CheckBox chkActivarERROR 
         BackColor       =   &H00000000&
         Caption         =   "ACTIVAR REGISTRO DE ERROR PERMANENETE"
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
         Height          =   585
         Left            =   270
         TabIndex        =   174
         Top             =   3180
         Width           =   4020
      End
      Begin VB.TextBox txtCortaMusicaPaga 
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
         Left            =   3330
         TabIndex        =   172
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2760
         Width           =   600
      End
      Begin VB.VScrollBar vsCortaMusicaPaga 
         Height          =   330
         LargeChange     =   10
         Left            =   3960
         Max             =   10
         Min             =   100
         SmallChange     =   10
         TabIndex        =   171
         Top             =   2760
         Value           =   10
         Width           =   330
      End
      Begin VB.ComboBox cmbIDIOMA 
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
         ItemData        =   "frmConfig.frx":0442
         Left            =   2070
         List            =   "frmConfig.frx":0452
         Style           =   2  'Dropdown List
         TabIndex        =   101
         Top             =   5070
         Visible         =   0   'False
         Width           =   2205
      End
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
         Left            =   3210
         TabIndex        =   54
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2010
         Width           =   720
      End
      Begin VB.VScrollBar VSSegEspera 
         Height          =   330
         LargeChange     =   10
         Left            =   3960
         Max             =   0
         Min             =   7200
         SmallChange     =   10
         TabIndex        =   53
         Top             =   2010
         Value           =   30
         Width           =   330
      End
      Begin VB.VScrollBar VsPorcTema 
         Height          =   330
         LargeChange     =   10
         Left            =   3960
         Max             =   10
         Min             =   100
         SmallChange     =   10
         TabIndex        =   52
         Top             =   2385
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
         Left            =   3330
         TabIndex        =   51
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2370
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
         Left            =   3330
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1650
         Width           =   600
      End
      Begin VB.VScrollBar VSmaxFichas 
         Height          =   330
         Left            =   3960
         Max             =   5
         Min             =   200
         TabIndex        =   47
         Top             =   1650
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
         Left            =   210
         TabIndex        =   43
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   330
            Width           =   3885
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo de fade in / fade out al enganchar canciones"
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
         Index           =   25
         Left            =   270
         TabIndex        =   184
         Top             =   3840
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cortar canciones pagas en %"
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
         Index           =   40
         Left            =   180
         TabIndex        =   173
         Top             =   2820
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "IDIOMA"
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
         Index           =   27
         Left            =   900
         TabIndex        =   102
         Top             =   5160
         Visible         =   0   'False
         Width           =   1065
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
         Left            =   180
         TabIndex        =   55
         Top             =   2430
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
         Left            =   90
         TabIndex        =   56
         Top             =   2070
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
         Left            =   270
         TabIndex        =   46
         Top             =   1710
         Width           =   2925
      End
   End
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
      Height          =   5055
      Left            =   4440
      TabIndex        =   18
      Top             =   10080
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox chkOutTemasWhenSel 
         BackColor       =   &H00000000&
         Caption         =   "Salir de listado de musica al hacer una selección"
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
         TabIndex        =   170
         Top             =   1710
         Width           =   4875
      End
      Begin VB.CommandButton Command25 
         Caption         =   "elegir imagen"
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
         Height          =   285
         Left            =   7200
         TabIndex        =   166
         Top             =   3510
         Width           =   1365
      End
      Begin VB.CheckBox chkLoadImgFondo2 
         BackColor       =   &H00000000&
         Caption         =   "Cargar imagen de fondo de los indicadores inferiores"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   4440
         TabIndex        =   168
         Top             =   3270
         Width           =   3855
      End
      Begin VB.TextBox txtImgFondo2 
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
         Height          =   705
         Left            =   4350
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   167
         Top             =   3870
         Width           =   4250
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
         Left            =   5190
         TabIndex        =   27
         Top             =   420
         Width           =   2745
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
         Left            =   5190
         TabIndex        =   79
         Top             =   180
         Width           =   3345
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
         Left            =   60
         TabIndex        =   21
         Top             =   930
         Width           =   3435
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
         Left            =   5190
         TabIndex        =   97
         Top             =   630
         Width           =   3315
      End
      Begin VB.CheckBox chkVisEXCL 
         BackColor       =   &H00000000&
         Caption         =   "CAMBIAR A VISUALIZACION  SUPERLICENCIA EXCLUSIVA"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   90
         TabIndex        =   112
         Top             =   4680
         Width           =   8475
      End
      Begin VB.CommandButton Command22 
         Caption         =   "elegir imagen"
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
         Height          =   285
         Left            =   60
         TabIndex        =   111
         Top             =   3510
         Width           =   1305
      End
      Begin VB.TextBox txtImgFondo 
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
         Height          =   705
         Left            =   60
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   110
         Top             =   3870
         Width           =   4250
      End
      Begin VB.CheckBox chkLoadImgFondo 
         BackColor       =   &H00000000&
         Caption         =   "Cargar imagen de fondo de las portadas"
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   60
         TabIndex        =   109
         Top             =   3210
         Width           =   4365
      End
      Begin VB.CheckBox chkBloquearMusicaElegida 
         BackColor       =   &H00000000&
         Caption         =   "Evitar selección multiple de un mismo tema en un disco"
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
         Left            =   60
         TabIndex        =   99
         Top             =   1440
         Width           =   5115
      End
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
         Height          =   315
         Left            =   5190
         TabIndex        =   100
         Top             =   900
         Width           =   3285
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
         Left            =   60
         TabIndex        =   98
         Top             =   1140
         Width           =   4875
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
         Height          =   495
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Text            =   "frmConfig.frx":047C
         Top             =   2400
         Width           =   5910
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
         Left            =   5190
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1515
         Width           =   600
      End
      Begin VB.VScrollBar vsDiscosV 
         Height          =   330
         LargeChange     =   10
         Left            =   5790
         Max             =   1
         Min             =   6
         TabIndex        =   25
         Top             =   1530
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsDiscosH 
         Height          =   330
         LargeChange     =   10
         Left            =   5790
         Max             =   1
         Min             =   6
         TabIndex        =   24
         Top             =   1200
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
         Left            =   5190
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1200
         Width           =   600
      End
      Begin VB.CheckBox chkDistorcionarTapas 
         BackColor       =   &H00000000&
         Caption         =   "Distorcionar tapas de discos ocupando 100% pantalla"
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
         Left            =   60
         TabIndex        =   22
         Top             =   450
         Width           =   4935
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
         Left            =   60
         TabIndex        =   20
         Top             =   690
         Width           =   5355
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
         Left            =   60
         TabIndex        =   19
         Top             =   210
         Width           =   5295
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "SOLO SUPERLICENCIA"
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
         Left            =   60
         TabIndex        =   169
         Top             =   2940
         Width           =   8535
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
         Left            =   90
         TabIndex        =   50
         Top             =   2160
         Width           =   1995
      End
      Begin VB.Label Label1 
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
         Left            =   6150
         TabIndex        =   29
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label1 
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
         Left            =   6150
         TabIndex        =   28
         Top             =   1230
         Width           =   1635
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
      Height          =   4230
      Left            =   12390
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   8160
      Begin VB.PictureBox PicContLetras 
         BackColor       =   &H00000040&
         Height          =   2715
         Left            =   150
         ScaleHeight     =   2655
         ScaleWidth      =   7875
         TabIndex        =   113
         Top             =   210
         Width           =   7935
         Begin VB.CommandButton Command24 
            Height          =   1270
            Left            =   7380
            Picture         =   "frmConfig.frx":04BC
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   1320
            Width           =   465
         End
         Begin VB.CommandButton Command23 
            Height          =   1270
            Left            =   7380
            Picture         =   "frmConfig.frx":08FE
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   30
            Width           =   465
         End
         Begin VB.PictureBox PicLetras 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   5535
            Left            =   0
            ScaleHeight     =   5535
            ScaleWidth      =   7275
            TabIndex        =   114
            Top             =   0
            Width           =   7275
            Begin VB.TextBox txtTeclas 
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
               Index           =   15
               Left            =   7260
               TabIndex        =   232
               Top             =   5100
               Width           =   700
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   15
               ItemData        =   "frmConfig.frx":0D40
               Left            =   2010
               List            =   "frmConfig.frx":0E6D
               Style           =   2  'Dropdown List
               TabIndex        =   230
               Top             =   5070
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   8
               ItemData        =   "frmConfig.frx":15EF
               Left            =   2010
               List            =   "frmConfig.frx":171C
               Style           =   2  'Dropdown List
               TabIndex        =   144
               Top             =   2745
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   7
               ItemData        =   "frmConfig.frx":1E9E
               Left            =   2010
               List            =   "frmConfig.frx":1FCB
               Style           =   2  'Dropdown List
               TabIndex        =   143
               Top             =   2415
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   6
               ItemData        =   "frmConfig.frx":274D
               Left            =   2010
               List            =   "frmConfig.frx":287A
               Style           =   2  'Dropdown List
               TabIndex        =   142
               Top             =   2085
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   5
               ItemData        =   "frmConfig.frx":2FFC
               Left            =   2010
               List            =   "frmConfig.frx":3129
               Style           =   2  'Dropdown List
               TabIndex        =   141
               Top             =   1755
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   4
               ItemData        =   "frmConfig.frx":38AB
               Left            =   2010
               List            =   "frmConfig.frx":39D8
               Style           =   2  'Dropdown List
               TabIndex        =   140
               Top             =   1425
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   3
               ItemData        =   "frmConfig.frx":415A
               Left            =   2010
               List            =   "frmConfig.frx":4287
               Style           =   2  'Dropdown List
               TabIndex        =   139
               Top             =   1095
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   2
               ItemData        =   "frmConfig.frx":4A09
               Left            =   2010
               List            =   "frmConfig.frx":4B36
               Style           =   2  'Dropdown List
               TabIndex        =   138
               Top             =   765
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   1
               ItemData        =   "frmConfig.frx":52B8
               Left            =   2010
               List            =   "frmConfig.frx":53E5
               Style           =   2  'Dropdown List
               TabIndex        =   137
               Top             =   435
               Width           =   5235
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   0
               ItemData        =   "frmConfig.frx":5B67
               Left            =   2010
               List            =   "frmConfig.frx":5C94
               Style           =   2  'Dropdown List
               TabIndex        =   136
               Top             =   90
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   6
               Left            =   7275
               TabIndex        =   135
               Top             =   2055
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   7
               Left            =   7275
               TabIndex        =   134
               Top             =   2385
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   8
               Left            =   7275
               TabIndex        =   133
               Top             =   2745
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   5
               Left            =   7275
               TabIndex        =   132
               Top             =   1740
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   4
               Left            =   7275
               TabIndex        =   131
               Top             =   1410
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   3
               Left            =   7275
               TabIndex        =   130
               Top             =   1080
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   2
               Left            =   7275
               TabIndex        =   129
               Top             =   750
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   1
               Left            =   7290
               TabIndex        =   128
               Top             =   420
               Width           =   700
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   0
               Left            =   7260
               TabIndex        =   127
               Top             =   60
               Width           =   180
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   9
               ItemData        =   "frmConfig.frx":6416
               Left            =   2010
               List            =   "frmConfig.frx":6543
               Style           =   2  'Dropdown List
               TabIndex        =   126
               Top             =   3075
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   9
               Left            =   7275
               TabIndex        =   125
               Top             =   3075
               Width           =   700
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   10
               ItemData        =   "frmConfig.frx":6CC5
               Left            =   2010
               List            =   "frmConfig.frx":6DF2
               Style           =   2  'Dropdown List
               TabIndex        =   124
               Top             =   3405
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   10
               Left            =   7320
               TabIndex        =   123
               Top             =   3405
               Width           =   660
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   11
               ItemData        =   "frmConfig.frx":7574
               Left            =   2010
               List            =   "frmConfig.frx":76A1
               Style           =   2  'Dropdown List
               TabIndex        =   122
               Top             =   3735
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   11
               Left            =   7275
               TabIndex        =   121
               Top             =   3735
               Width           =   700
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   12
               ItemData        =   "frmConfig.frx":7E23
               Left            =   2010
               List            =   "frmConfig.frx":7F50
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   4065
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   12
               Left            =   7275
               TabIndex        =   119
               Top             =   4065
               Width           =   700
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   13
               ItemData        =   "frmConfig.frx":86D2
               Left            =   2010
               List            =   "frmConfig.frx":87FF
               Style           =   2  'Dropdown List
               TabIndex        =   118
               Top             =   4410
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   13
               Left            =   7275
               TabIndex        =   117
               Top             =   4410
               Width           =   700
            End
            Begin VB.ComboBox cmbTECLAS 
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
               Index           =   14
               ItemData        =   "frmConfig.frx":8F81
               Left            =   2010
               List            =   "frmConfig.frx":90AE
               Style           =   2  'Dropdown List
               TabIndex        =   116
               Top             =   4740
               Width           =   5235
            End
            Begin VB.TextBox txtTeclas 
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
               Index           =   14
               Left            =   7275
               TabIndex        =   115
               Top             =   4740
               Width           =   700
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Nueva ficha (2)"
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
               Index           =   44
               Left            =   -480
               TabIndex        =   231
               Top             =   5115
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
               Left            =   -480
               TabIndex        =   159
               Top             =   105
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Pag. Adelante / Abajo"
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
               Left            =   -480
               TabIndex        =   158
               Top             =   2085
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Página Atras / Arriba"
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
               Left            =   -480
               TabIndex        =   157
               Top             =   2445
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   156
               Top             =   2745
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   155
               Top             =   1785
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   154
               Top             =   1455
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   153
               Top             =   1125
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   152
               Top             =   795
               Width           =   2445
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
               Left            =   -480
               TabIndex        =   151
               Top             =   480
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Mostrar Contador"
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
               Index           =   33
               Left            =   -480
               TabIndex        =   150
               Top             =   3075
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Poner Cero Contador"
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
               Index           =   34
               Left            =   -480
               TabIndex        =   149
               Top             =   3405
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tecla Fast Forward (FF)"
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
               Index           =   35
               Left            =   -480
               TabIndex        =   148
               Top             =   3735
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Bajar Volumen"
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
               Index           =   36
               Left            =   -480
               TabIndex        =   147
               Top             =   4065
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Subir Volumen"
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
               Index           =   37
               Left            =   -480
               TabIndex        =   146
               Top             =   4425
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Siguiente Tema"
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
               Index           =   38
               Left            =   -480
               TabIndex        =   145
               Top             =   4755
               Width           =   2445
            End
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Modo teclado"
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
         Height          =   780
         Left            =   150
         TabIndex        =   104
         Top             =   3360
         Width           =   7980
         Begin VB.OptionButton opModo4Teclas 
            BackColor       =   &H00000000&
            Caption         =   "Modo 4/6 teclas"
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
            Left            =   90
            TabIndex        =   105
            Top             =   210
            Width           =   1695
         End
         Begin VB.OptionButton opModo5Teclas 
            BackColor       =   &H00000000&
            Caption         =   "Modo 5 teclas"
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
            Left            =   90
            TabIndex        =   106
            Top             =   450
            Width           =   1590
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Adelante, Atras, Arriba, Abajo y OK. Sale de los discos con Adel-Atras"
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
            Height          =   270
            Index           =   32
            Left            =   1800
            TabIndex        =   108
            Top             =   480
            Width           =   6135
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Adelante, Atras, OK, ESC (Pagina Adelante y Pagina Atras)"
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
            Height          =   270
            Index           =   31
            Left            =   1830
            TabIndex        =   107
            Top             =   210
            Width           =   6045
         End
      End
      Begin VB.CheckBox chkPasarhoja 
         BackColor       =   &H00000000&
         Caption         =   "Pasa páginas con botones Adel-Atras"
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
         Left            =   3900
         TabIndex        =   31
         Top             =   3030
         Width           =   3510
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
         Left            =   300
         TabIndex        =   14
         Top             =   3000
         Width           =   3480
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
      Left            =   1170
      TabIndex        =   30
      Top             =   6180
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
         TabIndex        =   83
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   84
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
         TabIndex        =   33
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
         TabIndex        =   85
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
         TabIndex        =   32
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
         TabIndex        =   35
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
         TabIndex        =   34
         Top             =   2310
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
      Height          =   5505
      Left            =   12390
      TabIndex        =   188
      Top             =   4440
      Visible         =   0   'False
      Width           =   6285
      Begin VB.ComboBox cmbSCM 
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
         ItemData        =   "frmConfig.frx":9830
         Left            =   2790
         List            =   "frmConfig.frx":983A
         Style           =   2  'Dropdown List
         TabIndex        =   233
         Top             =   4920
         Width           =   2205
      End
      Begin VB.TextBox txtPrecioBase2 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   5400
         TabIndex        =   214
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   1140
         Width           =   810
      End
      Begin VB.TextBox txtExplicPrecios 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   213
         Top             =   3420
         Width           =   6015
      End
      Begin VB.TextBox txtPrecioBASE 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         TabIndex        =   212
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   780
         Width           =   810
      End
      Begin VB.VScrollBar vsCreditosBilletes 
         Height          =   330
         LargeChange     =   10
         Left            =   4590
         Max             =   1
         Min             =   100
         TabIndex        =   211
         Top             =   1140
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosBilletes 
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
         Left            =   3990
         TabIndex        =   210
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1140
         Width           =   600
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
         Index           =   2
         Left            =   1710
         TabIndex        =   209
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2985
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   2
         LargeChange     =   10
         Left            =   2310
         Max             =   0
         Min             =   100
         TabIndex        =   208
         Top             =   2985
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
         Index           =   2
         Left            =   3960
         TabIndex        =   207
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2985
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   2
         LargeChange     =   10
         Left            =   4560
         Max             =   0
         Min             =   100
         TabIndex        =   206
         Top             =   2985
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioM 
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
         Index           =   2
         Left            =   2700
         TabIndex        =   205
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2985
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
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
         Index           =   2
         Left            =   4920
         TabIndex        =   204
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2985
         Width           =   1100
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
         Index           =   1
         Left            =   1710
         TabIndex        =   203
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2625
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   1
         LargeChange     =   10
         Left            =   2310
         Max             =   0
         Min             =   100
         TabIndex        =   202
         Top             =   2625
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
         Index           =   1
         Left            =   3960
         TabIndex        =   201
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2625
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   1
         LargeChange     =   10
         Left            =   4560
         Max             =   0
         Min             =   100
         TabIndex        =   200
         Top             =   2625
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioM 
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
         Index           =   1
         Left            =   2700
         TabIndex        =   199
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2625
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
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
         Index           =   1
         Left            =   4920
         TabIndex        =   198
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2625
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
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
         Index           =   0
         Left            =   4920
         TabIndex        =   197
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2265
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioM 
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
         Index           =   0
         Left            =   2700
         TabIndex        =   196
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2265
         Width           =   1100
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   0
         LargeChange     =   10
         Left            =   4560
         Max             =   0
         Min             =   100
         TabIndex        =   195
         Top             =   2265
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
         Index           =   0
         Left            =   3960
         TabIndex        =   194
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2265
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "En cero"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   210
         Width           =   525
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
         Left            =   3990
         TabIndex        =   192
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   795
         Width           =   600
      End
      Begin VB.VScrollBar VSTemasXCredito 
         Height          =   330
         LargeChange     =   10
         Left            =   4590
         Max             =   1
         Min             =   100
         TabIndex        =   191
         Top             =   780
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   0
         LargeChange     =   10
         Left            =   2310
         Max             =   0
         Min             =   100
         TabIndex        =   190
         Top             =   2265
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
         Index           =   0
         Left            =   1710
         TabIndex        =   189
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2265
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mostar los creditos como"
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
         Height          =   165
         Index           =   45
         Left            =   300
         TabIndex        =   234
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "= $"
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
         Height          =   315
         Index           =   42
         Left            =   4920
         TabIndex        =   229
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X1"
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
         Height          =   315
         Index           =   54
         Left            =   1470
         TabIndex        =   228
         Top             =   2310
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X3"
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
         Height          =   315
         Index           =   46
         Left            =   1470
         TabIndex        =   227
         Top             =   3060
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Poner en cero X1 es modo gratuito. Poner en cero X2 o X3 es no usar promociones."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   1455
         Index           =   53
         Left            =   120
         TabIndex        =   226
         Top             =   1890
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "= $"
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
         Height          =   315
         Index           =   52
         Left            =   4920
         TabIndex        =   225
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   795
         Index           =   51
         Left            =   60
         TabIndex        =   224
         Top             =   1770
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Los créditos no son necesariamente canciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   225
         Index           =   50
         Left            =   120
         TabIndex        =   223
         Top             =   1470
         Width           =   4065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos por cada señal de billetero (S)"
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
         Index           =   49
         Left            =   90
         TabIndex        =   222
         Top             =   1170
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X2"
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
         Height          =   315
         Index           =   43
         Left            =   1470
         TabIndex        =   221
         Top             =   2700
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00404080&
         X1              =   120
         X2              =   6150
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblContador2 
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
         ForeColor       =   &H0000FFFF&
         Height          =   345
         Left            =   1680
         TabIndex        =   220
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contador historico/Interno"
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
         Index           =   39
         Left            =   -30
         TabIndex        =   219
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos para VIDEO/KARAOKE"
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
         Index           =   28
         Left            =   3960
         TabIndex        =   218
         Top             =   1770
         Width           =   2205
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
         Left            =   3660
         TabIndex        =   217
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos por cada señal de monedero (Q)"
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
         Index           =   11
         Left            =   90
         TabIndex        =   216
         Top             =   810
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Creditos para musica"
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
         Index           =   26
         Left            =   1710
         TabIndex        =   215
         Top             =   1770
         Width           =   2205
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
      Height          =   2835
      Left            =   6390
      TabIndex        =   88
      Top             =   270
      Visible         =   0   'False
      Width           =   5385
      Begin VB.CheckBox chkVidMudos 
         BackColor       =   &H00000000&
         Caption         =   "Usar la salida de TV para reproducir videos MUDOS. Esto anula las imagenes grandes en el TV"
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
         Left            =   180
         TabIndex        =   185
         Top             =   2100
         Width           =   4995
      End
      Begin VB.VScrollBar vsPubliIMGCada 
         Height          =   330
         Left            =   4800
         Max             =   10
         Min             =   100
         TabIndex        =   92
         Top             =   600
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
         Left            =   4170
         TabIndex        =   95
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
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
         Height          =   240
         Left            =   210
         TabIndex        =   91
         Top             =   300
         Width           =   4515
      End
      Begin VB.VScrollBar vsPubliCada 
         Height          =   330
         Left            =   4890
         Max             =   1
         Min             =   100
         TabIndex        =   90
         Top             =   1620
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
         Left            =   4260
         TabIndex        =   93
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1620
         Width           =   600
      End
      Begin VB.CheckBox ckPUB 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir Publicidades (Audio y video)  CON SONIDO altercando la reproducciones pagadas."
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
         Height          =   450
         Left            =   270
         TabIndex        =   89
         Top             =   1170
         Width           =   4665
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         X1              =   300
         X2              =   4770
         Y1              =   1020
         Y2              =   1020
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
         Height          =   255
         Index           =   30
         Left            =   210
         TabIndex        =   96
         Top             =   630
         Width           =   3795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Reproducir estas publicidades cada X temas"
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
         Index           =   29
         Left            =   375
         TabIndex        =   94
         Top             =   1650
         Width           =   3840
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
      Height          =   2265
      Left            =   3150
      TabIndex        =   36
      Top             =   3180
      Visible         =   0   'False
      Width           =   7275
      Begin VB.TextBox txtTamanoTapaPermitido 
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
         Left            =   6240
         TabIndex        =   236
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1770
         Width           =   480
      End
      Begin VB.VScrollBar vsTamanoTapaPermitido 
         Height          =   330
         LargeChange     =   10
         Left            =   6750
         Max             =   20
         Min             =   200
         SmallChange     =   10
         TabIndex        =   235
         Top             =   1770
         Value           =   200
         Width           =   330
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
         Left            =   150
         TabIndex        =   39
         Top             =   300
         Width           =   6885
         Begin VB.OptionButton OpImgINI 
            BackColor       =   &H00000000&
            Caption         =   "Cargar imagenes al inicio. Recomendado hasta 150 discos"
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
            TabIndex        =   41
            Top             =   300
            Width           =   6390
         End
         Begin VB.OptionButton OpImgSIS 
            BackColor       =   &H00000000&
            Caption         =   "Cargar las imagenes a pedido. Recomendado mas de 150 discos"
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
            TabIndex        =   40
            Top             =   570
            Value           =   -1  'True
            Width           =   6570
         End
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
         TabIndex        =   38
         Top             =   1620
         Width           =   2175
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
         TabIndex        =   37
         Top             =   1320
         Width           =   5890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño maximo en KB permitido para portadas"
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
         Index           =   47
         Left            =   3390
         TabIndex        =   237
         Top             =   1710
         Width           =   2805
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00000000&
      Caption         =   "Clave"
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
      Height          =   1935
      Left            =   60
      TabIndex        =   175
      Top             =   2220
      Width           =   2865
      Begin VB.CommandButton Command16 
         BackColor       =   &H00FFC0C0&
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   180
         Top             =   240
         Width           =   525
      End
      Begin VB.TextBox txtClaveAdmin 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   178
         Top             =   1110
         Width           =   2625
      End
      Begin VB.CommandButton Command31 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ingreso Administrador"
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
         TabIndex        =   177
         Top             =   1470
         Width           =   2650
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cambiar/Crear Clave"
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
         TabIndex        =   176
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escriba aqui su clave de administrador"
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
         Index           =   41
         Left            =   150
         TabIndex        =   179
         Top             =   660
         Width           =   2535
      End
   End
   Begin VB.HScrollBar HSvolumen 
      Height          =   240
      LargeChange     =   10
      Left            =   8970
      Max             =   100
      TabIndex        =   163
      Top             =   5610
      Width           =   2895
   End
   Begin VB.HScrollBar HSVolumen2 
      Height          =   240
      LargeChange     =   10
      Left            =   9000
      Max             =   100
      TabIndex        =   162
      Top             =   5940
      Width           =   2895
   End
   Begin VB.CommandButton Command21 
      BackColor       =   &H0080C0FF&
      Caption         =   "CLUF"
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
      Left            =   8070
      Style           =   1  'Graphical
      TabIndex        =   103
      Top             =   8490
      Width           =   1620
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
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   80
      Top             =   8490
      Width           =   2220
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
      Left            =   60
      TabIndex        =   60
      Top             =   9120
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
         TabIndex        =   77
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
         TabIndex        =   76
         Top             =   1530
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
         TabIndex        =   75
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
         TabIndex        =   73
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
         TabIndex        =   72
         Top             =   780
         Width           =   2800
      End
      Begin VB.TextBox txtRegistroDiario 
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   70
         Text            =   "frmConfig.frx":9856
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
         TabIndex        =   65
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         TabIndex        =   64
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
         Left            =   120
         TabIndex        =   61
         Top             =   270
         Width           =   3975
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
         Left            =   120
         TabIndex        =   78
         Top             =   1620
         Width           =   4875
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
         TabIndex        =   74
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
         TabIndex        =   71
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
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
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
         TabIndex        =   66
         Top             =   810
         Width           =   2385
      End
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
      Height          =   4005
      Left            =   60
      TabIndex        =   59
      Top             =   4140
      Width           =   2895
      Begin VB.CommandButton Command26 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Importar/Exportar CONFIG"
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
         TabIndex        =   181
         Top             =   3585
         Width           =   2650
      End
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
         TabIndex        =   87
         Top             =   3210
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
         TabIndex        =   86
         Top             =   2835
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
         TabIndex        =   11
         Top             =   1710
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
         TabIndex        =   10
         Top             =   2085
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
         TabIndex        =   9
         Top             =   2460
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
         TabIndex        =   8
         Top             =   1335
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
         TabIndex        =   7
         Top             =   960
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
         TabIndex        =   5
         Top             =   210
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
         TabIndex        =   6
         Top             =   585
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
      Height          =   2145
      Left            =   60
      TabIndex        =   58
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
         Top             =   600
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
         Top             =   960
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
         Top             =   1320
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
         Top             =   1680
         Width           =   2650
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FF8080&
      Cancel          =   -1  'True
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
      TabIndex        =   13
      Top             =   8580
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
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8190
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
      Left            =   2970
      TabIndex        =   57
      Top             =   60
      Width           =   8925
   End
   Begin VB.Label LblVol 
      Alignment       =   1  'Right Justify
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
      Left            =   7560
      TabIndex        =   165
      Top             =   5610
      Width           =   1380
   End
   Begin VB.Line LineScroll 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   11850
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line LineScroll2 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      X1              =   9000
      X2              =   11880
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Label lblVol2 
      Alignment       =   1  'Right Justify
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
      Left            =   7560
      TabIndex        =   164
      Top             =   5910
      Width           =   1380
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
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle/Ayuda de la opcion elegida"
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
      Height          =   3345
      Left            =   3000
      TabIndex        =   16
      Top             =   5550
      Width           =   4575
   End
   Begin VB.Label lblTBRcfg 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmConfig.frx":9875
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
      Left            =   8070
      TabIndex        =   17
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

Public Sub SendW()
    Form_KeyDown TeclaCerrarSistema, 0
End Sub

Private Sub chkActivarERROR_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    HLP "Active solo en caso de que 3PM se ciere bruscamente con errores. NO ACTIVAR " + _
        "SI 3PM FUNCIONA CORRECTAMENTE. " + _
        "Luego de activar reinicie 3PM y luego de que se cierre con fallo busque " + _
        "en la carpeta de 3PM todos los archivos " + _
        "'REG*****.W15' y envíelos a tbrsoft (info@tbrsoft.com) detallando el mensaje de " + _
        "error que informa 3PM antes de cerrarse. Luego de esto recibira un email " + _
        "con el detalle de su error y la solución correspondiente"
End Sub

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

Private Sub chkBloquearMusicaElegida_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkBloquearMusicaElegida.ForeColor = vbYellow
    HLP "Si la activa cuando ingrese a algún disco y seleccione algun tema " + _
        "este quedará bloqueado hasta que vuelva a abrir el disco. Esto" + _
        " evita la seleccion multiple de un mismo tema varias veces continuadas"
End Sub

Private Sub chkBloquearMusicaElegida_LostFocus()
    chkBloquearMusicaElegida.ForeColor = vbWhite
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

Private Sub chkCS_Click()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkApagarPC.ForeColor = vbYellow
    HLP "Le permite corregir errores en la recepcion de las señales de su " + _
        "monedero / billetero electrónico. No lo active si no es muy necesario"
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

Private Sub chkLoadImgFondo_Click()
    'si lo tilda la imagen ya se carga
    If chkLoadImgFondo Then
        If txtImgFondo = "" Then
            MsgBox "No hay imagen para cargar!"
            chkLoadImgFondo = 0
            Exit Sub
        End If
        If FSO.FileExists(txtImgFondo) = False Then
            MsgBox "La imagen elegida no existe"
            chkLoadImgFondo = 0
            Exit Sub
        End If
        'cargarla!
        frmIndex.picFondoDisco.Picture = LoadPicture(txtImgFondo)
    Else
        'limpiarla!
        frmIndex.picFondoDisco.Picture = LoadPicture
    End If
End Sub

Private Sub chkLoadImgFondo_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkLoadImgFondo.ForeColor = vbYellow
    HLP "Modificar la imagen de fondo de los discos. Debera coincidir o superar el" + _
        " tamaño que este ocupa dependiendo de la configuración de pantalla que utilice"
End Sub

Private Sub chkLoadImgFondo_LostFocus()
    chkLoadImgFondo.ForeColor = vbWhite
End Sub

Private Sub chkLoadImgFondo2_Click()
    'si lo tilda la imagen ya se carga
    If chkLoadImgFondo2 Then
        If txtImgFondo2 = "" Then
            MsgBox "No hay imagen para cargar!"
            chkLoadImgFondo2 = 0
            Exit Sub
        End If
        If FSO.FileExists(txtImgFondo2) = False Then
            MsgBox "La imagen elegida no existe"
            chkLoadImgFondo2 = 0
            Exit Sub
        End If
        'cargarla!
        frmIndex.picFondo.Picture = LoadPicture(txtImgFondo2)
        frmIndex.VU21.Picture = LoadPicture(txtImgFondo2)
        
    Else
        'limpiarla!
        frmIndex.picFondo.Picture = LoadPicture
        frmIndex.VU21.Picture = LoadPicture
    End If
End Sub

Private Sub chkLoadImgFondo2_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkLoadImgFondo2.ForeColor = vbYellow
    HLP "Modificar la imagen de fondo de los indicadores de la parte inferior de la pantalla principal. Debera coincidir o superar el" + _
        " tamaño que este ocupa dependiendo de la configuración de pantalla que utilice"
End Sub

Private Sub chkLoadImgFondo2_LostFocus()
    chkLoadImgFondo2.ForeColor = vbWhite
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

Private Sub chkOutTemasWhenSel_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkOutTemasWhenSel.ForeColor = vbYellow
    HLP "Salir inmediatamente del listado de musica al hacer una selección"
End Sub

Private Sub chkOutTemasWhenSel_LostFocus()
    chkOutTemasWhenSel.ForeColor = vbWhite
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
        "Debe habilitarse la salida de TV como expansión del escritorio " + _
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

Private Sub chkVidMudos_Click()
    
    If PUBs.TotalPUBsMUTE = 0 Then
        MsgBox "No puede activar esta opción ya que no hay publicidades cargadas." + vbCrLf + _
            "Para cargar publicidades debera incluir en la carpeta 'PUBMUTE' (en la carpeta en " + _
            "que instalo 3PM) uno o más ficheros AVI, MPG, DAT (VCD) o VOB (DVD)"
        chkVidMudos = 0
    End If

End Sub

Private Sub chkVidMudos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVidMudos.ForeColor = vbYellow
    HLP "Indica si se reproducirán publicidades por la salida de TV " + _
        "sin sonido. Esto no interrumpe ninguna otra reproduccion " + _
        "de la rockola. Si se habilita esta opción deben colocarse ficheros" + _
        " de video AVI, MPG, VOB (DVD) o DAT (VCD) en la carpeta PUBMUTE (de la " + _
        "carpeta en la que instalo 3PM). Estos ficheros continuamente" + _
        " salvo que algun usuario cargue algun video pago." + _
        " Se reproducen en orden alfabético por lo que podrá " + _
        "modificar el nombre para definir el orden deseado. Habilitar" + _
        " esta opcion anulas las imagenes publictarias destinadas al tv"
End Sub

Private Sub chkVidMudos_LostFocus()
    chkVidMudos.ForeColor = vbWhite
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
            "que instalo 3PM) uno o más ficheros MP3, WMA, AVI, MPG, VOB (DVD) o DAT (VCD)"
        ckPUB = 0
    End If
End Sub

Private Sub ckPUB_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    ckPUB.ForeColor = vbYellow
    HLP "Indica si se reproducirán publicidades. Si se habilita esta opción deben colocarse ficheros " + _
        "MP3, WMA, AVI, MPG, VOB (DVD) o DAT (VCD) en la carpeta PUB (de la carpeta en la que instalo 3PM). Estos ficheros se reproducen cada X (a configurar) " + _
        "temas y de a uno por vez. Se reproducen en orden alfabético por lo que podrá " + _
        "modificar el nombre para definir el orden deseado. Puede tambien duplicar ficheros para " + _
        "darle mayor repeticion a alguna publicidad en particular"
End Sub

Private Sub ckPUB_LostFocus()
    ckPUB.ForeColor = vbWhite
End Sub

Private Sub ckPubIMG_Click()
    If ckPubIMG Then
        If PUBs.TotalPUBsIMG = 0 Then
            MsgBox "No puede activar esta opción ya que no hay publicidades (de menos de 50KB) cargadas." + vbCrLf + _
                "Para cargar publicidades debera incluir en la carpeta 'PUB' (en la carpeta en " + _
                "que instalo 3PM) uno o más ficheros JPG, BMP o GIF. " + _
                "Debera reiniciar 3PM para que este cambio surta efecto"
            ckPubIMG = 0
        End If
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
    On Error GoTo MiErr
    tERR.Anotar "aclp"
    'GRABAR BUTTON
    'cargar los datos del archivo GPF("config")
    'paso todo a una cadena, la encripto y luego la escribo
    Dim FullConfig As String
    FullConfig = FullConfig + "ClaveAdmin=" + ClaveAdmin + vbCrLf
    FullConfig = FullConfig + "CargarImagenInicio=" + CStr(OpImgINI) + vbCrLf
    FullConfig = FullConfig + "AutoReDraw=" + CStr(chkAutoReDraw) + vbCrLf
    FullConfig = FullConfig + "TeclaDerecha=" + txtTeclas(0) + vbCrLf
    FullConfig = FullConfig + "TeclaIzquierda=" + txtTeclas(1) + vbCrLf
    FullConfig = FullConfig + "TeclaOK=" + txtTeclas(2) + vbCrLf
    FullConfig = FullConfig + "TeclaESC=" + txtTeclas(3) + vbCrLf
    FullConfig = FullConfig + "TeclaNuevaFicha=" + txtTeclas(4) + vbCrLf
    FullConfig = FullConfig + "TeclaNuevaFicha2=" + txtTeclas(15) + vbCrLf
    FullConfig = FullConfig + "TeclaConfig=" + txtTeclas(5) + vbCrLf
    FullConfig = FullConfig + "TeclaPagAd=" + txtTeclas(6) + vbCrLf
    FullConfig = FullConfig + "TeclaPagAt=" + txtTeclas(7) + vbCrLf
    FullConfig = FullConfig + "TeclaCerrarSistema=" + txtTeclas(8) + vbCrLf
    tERR.Anotar "aclq"
    FullConfig = FullConfig + "ShowCreditsMode=" + CStr(cmbSCM.ListIndex) + vbCrLf
    ShowCreditsMode = cmbSCM.ListIndex
    FullConfig = FullConfig + "TeclaShowContador=" + txtTeclas(9) + vbCrLf
    FullConfig = FullConfig + "TeclaPutCeroContador=" + txtTeclas(10) + vbCrLf
    FullConfig = FullConfig + "TeclaFF=" + txtTeclas(11) + vbCrLf
    FullConfig = FullConfig + "TeclaBajaVolumen=" + txtTeclas(12) + vbCrLf
    FullConfig = FullConfig + "TeclaSubeVolumen=" + txtTeclas(13) + vbCrLf
    FullConfig = FullConfig + "TeclaNextMusic=" + txtTeclas(14) + vbCrLf
    FullConfig = FullConfig + "ActivarCorreccionSignal=" + CStr(chkCS) + vbCrLf
    FullConfig = FullConfig + "ApagarAlCierre= " + CStr(chkApagarPC) + vbCrLf
    FullConfig = FullConfig + "ActivarERR= " + CStr(chkActivarERROR) + vbCrLf
    FullConfig = FullConfig + "TamanoTapaPermitido= " + CStr(vsTamanoTapaPermitido) + vbCrLf
    
    tERR.Anotar "aclr"
    If opModo4Teclas Then
        FullConfig = FullConfig + "IsMod46Teclas= 46" + vbCrLf
        IsMod46Teclas = 46
    End If
    If opModo5Teclas Then
        FullConfig = FullConfig + "IsMod46Teclas= 5" + vbCrLf
        IsMod46Teclas = 5
    End If
    FullConfig = FullConfig + "RankToPeople= " + CStr(chkRankToPeople) + vbCrLf
    FullConfig = FullConfig + "MaximoFichas=" + txtMaxFichas + vbCrLf
    FullConfig = FullConfig + "EsperaMinutos=" + txtSECwait + vbCrLf
    FullConfig = FullConfig + "FastIni=" + CStr(chkFastINI) + vbCrLf
    FullConfig = FullConfig + "HabilitarVUMetro=" + CStr(chkVUMeter) + vbCrLf
    FullConfig = FullConfig + "VidfullScreen=" + CStr(chkVidFullScreen) + vbCrLf
    tERR.Anotar "acls"
    FullConfig = FullConfig + "Salida2=" + CStr(chkSalida2) + vbCrLf
    FullConfig = FullConfig + "NoVumVid=" + CStr(chkNoVumVID) + vbCrLf
    FullConfig = FullConfig + "OutTemasWhenSel=" + CStr(chkOutTemasWhenSel) + vbCrLf
    FullConfig = FullConfig + "BloquearMusicaElegida=" + CStr(chkBloquearMusicaElegida) + vbCrLf
    'Valores de ReIni LISTA=solo lista NADA=arranca de cero
    If OpReiniFull Then
        FullConfig = FullConfig + "ReINI=LISTA" + vbCrLf
    Else
        FullConfig = FullConfig + "ReINI=NADA" + vbCrLf
    End If
    tERR.Anotar "aclt"
    FullConfig = FullConfig + "Volumen=" + Trim(CStr(HSvolumen)) + vbCrLf
    FullConfig = FullConfig + "Volumen2=" + Trim(CStr(HSVolumen2)) + vbCrLf
    FullConfig = FullConfig + "EsperaTecla=" + txtEsperaTecla + vbCrLf
    FullConfig = FullConfig + "PorcentajeTema=" + txtPorcTema + vbCrLf
    FullConfig = FullConfig + "SegFade=" + txtSegFade + vbCrLf
    SegFade = vsSegFade
    FullConfig = FullConfig + "DiscosH=" + txtDiscosH + vbCrLf
    FullConfig = FullConfig + "DiscosV=" + txtDiscosV + vbCrLf
    FullConfig = FullConfig + "DuracionProtect=" + txtDuracionProtect + vbCrLf
    tERR.Anotar "aclu"
    FullConfig = FullConfig + "PasarHoja=" + CStr(chkPasarhoja) + vbCrLf
    FullConfig = FullConfig + "DistorcionarTapas=" + CStr(chkDistorcionarTapas) + vbCrLf
    FullConfig = FullConfig + "3pmExcl=" + CStr(chkVisEXCL) + vbCrLf
    'valores para el protectore de pantalla
    '0=inhabilitado 1=Original 2=Carpeta Fotos 3= Video FullScreen
    If chkNoProtector Then
        FullConfig = FullConfig + "Protector=0" + vbCrLf
        frmIndex.Timer3.Interval = 0 'para el reloj del protector
    Else
        'si es cualquierra de los otros habilitarlo
        frmIndex.Timer3.Interval = 10000
    End If
    tERR.Anotar "aclv"
    If chkProtectOriginal Then FullConfig = FullConfig + "Protector=1" + vbCrLf
    If chkProtectorCustom Then FullConfig = FullConfig + "Protector=2" + vbCrLf
    FullConfig = FullConfig + "CargarDuracionTemas=" + CStr(chkCargarDuracionTemas) + vbCrLf
    FullConfig = FullConfig + "MostrarRotulos=" + CStr(chkMostrarRotulos) + vbCrLf
    FullConfig = FullConfig + "RotulosArriba=" + CStr(chkRotulosArriba) + vbCrLf
    FullConfig = FullConfig + "TemasPorCredito= " + txtTemasXCredito + vbCrLf
    FullConfig = FullConfig + "CreditosBilletes= " + txtCreditosBilletes + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTema= " + txtCreditosCuestaTema(0) + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTema2= " + txtCreditosCuestaTema(1) + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTema3= " + txtCreditosCuestaTema(2) + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTemaVIDEO= " + txtCreditosCuestaTemaVIDEO(0) + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTemaVIDEO2= " + txtCreditosCuestaTemaVIDEO(1) + vbCrLf
    FullConfig = FullConfig + "CreditosCuestaTemaVIDEO3= " + txtCreditosCuestaTemaVIDEO(2) + vbCrLf
    FullConfig = FullConfig + "PrecioBase= " + txtPrecioBASE + vbCrLf
    FullConfig = FullConfig + "PrecioBase2= " + txtPrecioBase2 + vbCrLf
    FullConfig = FullConfig + "TextoUsuario= " + TxtUSUARIO + vbCrLf
    tERR.Anotar "aclw"
    'validacion con clave cada x creditos
    FullConfig = FullConfig + "Validar= " + CStr(chkValidar) + vbCrLf
    FullConfig = FullConfig + "ValidarCada= " + txtValidarCada + vbCrLf
    FullConfig = FullConfig + "AvisarAntes= " + txtAvisarAntes + vbCrLf
    FullConfig = FullConfig + "MostrarTouch= " + CStr(chkTouch) + vbCrLf
    tERR.Anotar "aclx"
    'publicidades
    FullConfig = FullConfig + "MostrarPUB= " + CStr(ckPUB) + vbCrLf
    FullConfig = FullConfig + "MostrarPUBMute= " + CStr(chkVidMudos) + vbCrLf
    FullConfig = FullConfig + "MostrarPUBIMG= " + CStr(ckPubIMG) + vbCrLf
    FullConfig = FullConfig + "PubliCada= " + txtPubliCada + vbCrLf
    FullConfig = FullConfig + "PubliIMGCada= " + txtPubliImgCada + vbCrLf
    FullConfig = FullConfig + "Idioma= " + cmbIDIOMA + vbCrLf
    tERR.Anotar "acly"
    If chkLoadImgFondo Then
        FullConfig = FullConfig + "ImgFondo= " + txtImgFondo + vbCrLf
    Else
        FullConfig = FullConfig + "ImgFondo= NO" + vbCrLf
    End If
    
    If chkLoadImgFondo2 Then
        FullConfig = FullConfig + "ImgFondo2= " + txtImgFondo2 + vbCrLf
    Else
        FullConfig = FullConfig + "ImgFondo2= NO" + vbCrLf
    End If
    tERR.Anotar "aclz"
    'encriptar
    FullConfig = Encriptar(FullConfig, False)
    'grabar el kilombo
    Set TE = FSO.CreateTextFile(GPF("config"), True)
        TE.Write FullConfig
    TE.Close
    'hacer una copia de seguridad cada vez que haya cambios
    'xxxx ver desde que punto restaurarlo
    FSO.CopyFile GPF("config"), GPF("config2")
    
    'SI NO HAY que validar me aseguro que se borre el archivo de validacion sf + "radilav.cfg"
    If chkValidar.Value = 0 Then
        If FSO.FileExists(GPF("radliv")) Then FSO.DeleteFile GPF("radliv"), True
    End If
    tERR.Anotar "acma"
    'publicidades
    PUBs.SonarPublicidadesCada = Val(txtPubliCada)
    PUBs.HabilitarPublicidadesMp3Vid = ckPUB
    PUBs.HabilitarPublicidadesVMute = chkVidMudos
    
    PUBs.SonarPublicidadesIMGCada = Val(txtPubliImgCada)
    PUBs.HabilitarPublicidadesIMG = ckPubIMG
    IDIOMA = cmbIDIOMA
    tERR.Anotar "acmb"
    
    'todas las propiedades se quedan sin reiniciar
    'algunas no se necesitan
    'NO NECESITO CargarIMGinicio = LeerConfig("CargarImagenInicio")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
   
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    'solo se hace al inicio
    'ActivarERR = LeerConfig("ActivarERR", "0")
    tERR.Anotar "acmc"
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    'NO DEBO ReINI = LeerConfig("ReINI","LISTA")
    VolumenIni = CLng(LeerConfig("Volumen", "50"))
    VolumenIni2 = CLng(LeerConfig("Volumen2", "20"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    'NO NECESITO FASTini = LeerConfig("FastIni","1")
    PasarHoja = LeerConfig("PasarHoja", "1")
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    tERR.Anotar "acmd"
    
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosBilletes = LeerConfig("CreditosBilletes", "10")
    PrecioBase = LeerConfig("PrecioBase", "0,50")
    PrecioBase2 = LeerConfig("PrecioBase2", "10")
    CreditosCuestaTema(0) = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTema(1) = LeerConfig("CreditosCuestaTema2", "0")
    CreditosCuestaTema(2) = LeerConfig("CreditosCuestaTema3", "0")
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "0")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "0")
    
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "0")
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    tERR.Anotar "acme"
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    If K.LICENCIA = aSinCargar Then
        frmIndex.lblDEMO = "Este espacio sera suyo cuando adquiera la version full de 3PM"
    Else
        frmIndex.lblDEMO = textoUsuario
    End If
    tERR.Anotar "acmf"
    Unload Me
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aclp"
    Resume Next

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
    MsgBox "Si usted usa una version demo su clave es 'DEMO' y no se pude cambiar" + vbCrLf + _
        "Si ya dispone de una licencia paga su clave predeterminada es 'ADMIN' hasta " + _
        "que la cambia." + vbCrLf + _
        "Es muy recomendado que la cambie. Si tenia usted una clave ya recibida de " + _
        "versiones anteriores esta deja de tener validez. A partir de esta version cambia " + _
        "a 'ADMIN' hasta que la cambie usted"
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
    txtTraduccion = ""
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

Private Sub Command21_Click()
    frmCLUF.Show 1
End Sub

Private Sub Command22_Click()
    Dim CmdLg As New CommonDialog
    
    CmdLg.Filter = "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp"
    CmdLg.DialogTitle = "Seleccione imagen de fondo de 3PM"
    CmdLg.ShowOpen
    If CmdLg.FileName <> "" Then txtImgFondo = CmdLg.FileName
    Set CmdLg = Nothing
End Sub

Private Sub Command23_Click()
    If PicLetras.Top < 0 Then PicLetras.Top = PicLetras.Top + 300
    If PicLetras.Top > 0 Then PicLetras.Top = 0
End Sub

Private Sub Command24_Click()
    If PicLetras.Top > -PicLetras.Height + PicContLetras.Height Then PicLetras.Top = PicLetras.Top - 300
    If PicLetras.Top < -PicLetras.Height + PicContLetras.Height Then PicLetras.Top = -PicLetras.Height + PicContLetras.Height
End Sub

Private Sub Command25_Click()
    Dim CmdLg As New CommonDialog
    
    CmdLg.Filter = "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp"
    CmdLg.DialogTitle = "Seleccione imagen de fondo INDICADORES de 3PM"
    CmdLg.ShowOpen
    If CmdLg.FileName <> "" Then txtImgFondo2 = CmdLg.FileName
    Set CmdLg = Nothing
End Sub

Private Sub Command26_Click()
    frmImpExpCONFIG.Show 1
End Sub

Private Sub Command27_Click()
    If K.LICENCIA < GFull Then
        MsgBox "No puede cambiar la clave. Para versiones demo la clave es 'DEMO'"
        Exit Sub
    End If
    
    Dim ClaveSel As String
    ClaveSel = InputBox("Ingrese la anterior clave de administrador", "3PM CLAVE")
    
    If UCase(ClaveSel) = UCase(ClaveAdmin) Or UCase(ClaveSel) = "RMLVF" Then
        ClaveSel = InputBox("Ingreso Correcto." + vbCrLf + vbCrLf + _
            "Ingrese la nueva clave:", "3PM CLAVE")
        
        If ClaveSel = "" Then Exit Sub
        
        ClaveAdmin = ClaveSel
        MsgBox "Recuerde colocar 'GRABAR' al salir de esta pagina para que el cambio tenga efecto luego de reiniciado 3PM"
    Else
        MsgBox "Clave erronea"
    End If
        
    
End Sub

Private Sub Command27_GotFocus()
    TeclaConfOK = "{ENTER}"
    Command27.BackColor = vbYellow
    HLP "Si usted usa una version demo su clave es 'DEMO' y no se pude cambiar" + vbCrLf + _
        "Si ya dispone de una licencia paga su clave predeterminada es 'ADMIN' hasta " + _
        "que la cambia." + vbCrLf + _
        "Es muy recomendado que la cambie. Si tenia usted una clave ya recibida de " + _
        "versiones anteriores esta deja de tener validez. A partir de esta version cambia " + _
        "a 'ADMIN' hasta que la cambie usted"
End Sub

Private Sub Command27_LostFocus()
    Command27.BackColor = &HFFC0C0
End Sub

Private Sub Command28_Click()
    frmEspecialMonedero.Show 1
End Sub

Private Sub Command3_Click()
    SumarContadorCreditos -CONTADOR 'esto lo deja en cero
    lblContador = STRceros(CONTADOR, 11)
    lblContador2 = STRceros(CONTADOR2, 11)
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

Private Sub Command31_Click()
    'Ingresar Clave Admin BUTTON!!!
    'ClaveIngresada
    Dim TodoOk As Boolean
    TodoOk = False
    'si es una demo que permita la clave de administrador "DEMO"
    If K.LICENCIA = CGratuita And UCase(txtClaveAdmin) = "DEMO" Then TodoOk = True
    If K.LICENCIA = aSinCargar And UCase(txtClaveAdmin) = "DEMO" Then TodoOk = True
    'ver que la contraseña se tome desde el teclado al usuario
    If UCase(txtClaveAdmin) = UCase(ClaveAdmin) Or LCase(txtClaveAdmin) = "rmlvf" Then TodoOk = True
    
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
        Command26.Enabled = True
    Else
        MsgBox "La clave ingresada no es correcta"
    End If
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
    
    Dim V As vWindows
    V = vW.GetVersion
    Select Case V
    Case Win98, Win98SE, WinME
        frmINI3PM.Show 1
    Case Win2000, WinNT4, WinXp, WinXP2
        frmINI3PMxp.Show 1
    End Select

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
    
    'borrar el original...
    If FSO.FileExists(GPF("cd5pm")) Then FSO.DeleteFile GPF("cd5pm"), True
    '... y la copia
    If FSO.FileExists(GPF("cd6pm")) Then FSO.DeleteFile GPF("cd6pm"), True
    
    If FSO.FileExists(GPF("cd5pm")) Or FSO.FileExists(GPF("cd6pm")) Then
        MsgBox "No se ha podido borrar el archivo de licencia"
    Else
        MsgBox "La información de licencia se ha borrado correctamente. El sistema se cerrará " + _
            "para que cargue nuevamente su clave"
    End If
    
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
            YaCerrar3PM
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
        LTE 1
        VarCreditos CSng(TemasPorCredito)
            
        lblContador = STRceros(CONTADOR, 11)
        lblContador2 = STRceros(CONTADOR2, 11)
        
    End If
    
    If KeyCode = TeclaNewFicha2 Then
        LTE 2
        VarCreditos CSng(CreditosBilletes)
            
        lblContador = STRceros(CONTADOR, 11)
        lblContador2 = STRceros(CONTADOR2, 11)
        
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo MiErr
    tERR.Anotar "acmg", ClaveAdmin
    Dim S5 As String
    S5 = "Esta configuración dependera de si dispone usted " + _
        "monederos multimoneda o de unica moneda. " + vbCrLf + _
        "3PM toma como base las señales que envia el monedero y/o billetero" + _
        ", cada señal representa una X cantidad de créditos.  " + vbCrLf + _
        "Si tiene un monedero de moneda única por ejemplo puede usar monedas" + _
        " de $5. En este caso para que una canción cueste $10 hay que colocar los" + _
        " Creditos para Musica X1 en 2. Puede por ejemplo colocar Créditos " + _
        "para Musica X2 en 3 para que una cancion cueste $10 y 2 x $15. En " + _
        "este mismo caso si una cancion cuesta $5 no tendría sentido usar 'X2'" + _
        " y si por ejemplo porner X3 en 2. Con esto una canción costaría $5 " + _
        "y 3 canciones por $10. Para ocultar la promocion X2 sería recomendable" + _
        " ponerla en cero. Todo esto poniendo 'Creditos por señal' en 2 =$10" + _
        vbCrLf + "Si tiene monedero multimoneda las opciones son parecidas " + _
        "mejorarán los sobrantes que quedan sin usar. Se recomienda programar " + _
        "el monedero al valor menor para que los precios puedan manejarse mas " + _
        "comodamente. Por ejemplo el monedero recibe monedas de $1, $2, $5, $10." + _
        "Se programa para que mande señal cada $1. De esta forma coloca " + _
        "'Creditos por señal' en 1 =$1. Credito para musica X1=5 ($5) X2=8 ($8)" + _
        "X3=11 ($11). Esto a modo de ejemplo. Si se usan las monedas adecuadas " + _
        "no habrá sobrante de crédito nunca."
    
    txtExplicPrecios.Text = S5
    
    
    'caso especial Eduardo rodirguez
    If ClaveAdmin = "ERO77701192FF" Or ClaveAdmin = "MARC777" Then
        Command19.Visible = False
        Command21.Visible = False
    End If
    
    'poner en tamaño para que se ajuste bien
    Me.Height = 9000
    Me.Width = 12000
    MostrarCursor True
    AjustarFRM Me, 12000
    tERR.Anotar "acmh", K.LICENCIA
    If K.LICENCIA = HSuperLicencia Then
        'Permitir modificar la imagen de fondo
        chkLoadImgFondo.Enabled = True
        chkLoadImgFondo2.Enabled = True
        txtImgFondo.Enabled = True
        txtImgFondo2.Enabled = True
        Command22.Enabled = True
        Command25.Enabled = True
        'permite pasara a modo exclusivo
        chkVisEXCL.Enabled = True
        tERR.Anotar "acmi"
        If FSO.FileExists(GPF("telcnot")) Then
            Set TE = FSO.OpenTextFile(GPF("telcnot"), ForReading, False)
            If TE.AtEndOfStream = False Then
                Dim NewT As String
                NewT = TE.ReadAll
            Else
                NewT = "Error Al leer el archivo"
                tERR.AppendLog "NOLEE.w/sl/txtcfg.tbr", Me.Name + ".acpm"
            End If
            lblTBRcfg = NewT
            TE.Close
        Else
            lblTBRcfg = "Desarrollado por tbrSoft" + vbCrLf + "www.tbrsoft.com" + vbCrLf + _
                "----------------" + vbCrLf + "Contáctenos a info@tbrsoft.com" + vbCrLf + _
                "tbrsoft@cpcipc.org" + vbCrLf + "----------------" + vbCrLf + "Hecho en Argentina"
        End If
    Else
        tERR.Anotar "acmj"
        'no permitir modificar la imagen de fondo
        chkLoadImgFondo.Enabled = False
        chkLoadImgFondo2.Enabled = False
        txtImgFondo.Enabled = False
        txtImgFondo2.Enabled = False
        Command22.Enabled = False
        Command25.Enabled = False
        lblTBRcfg = "Desarrollado por tbrSoft" + vbCrLf + "www.tbrsoft.com" + vbCrLf + _
            "----------------" + vbCrLf + "Contáctenos a info@tbrsoft.com" + vbCrLf + _
            "tbrsoft@cpcipc.org" + vbCrLf + "----------------" + vbCrLf + "Hecho en Argentina"
    End If
    tERR.Anotar "acmk"
    lblContador = STRceros(CONTADOR, 11)
    lblContador2 = STRceros(CONTADOR2, 11)
    
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
    
    'leer el archivo de configuracion GPF("config")
    CargarIMGinicio = LeerConfig("CargarImagenInicio", "1")
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    AutoReDibuj = LeerConfig("AutoReDraw", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaNewFicha2 = Val(LeerConfig("TeclaNuevaFicha2", "83"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    tERR.Anotar "acml"
    TeclaShowContador = Val(LeerConfig("TeclaShowContador", "85")) 'U
    TeclaPutCeroContador = Val(LeerConfig("TeclaPutCeroContador", "86")) 'V
    TeclaFF = Val(LeerConfig("TeclaFF", "74")) 'J
    TeclaBajaVolumen = Val(LeerConfig("TeclaBajaVolumen", "68")) 'D
    TeclaSubeVolumen = Val(LeerConfig("TeclaSubeVolumen", "69")) 'E
    TeclaNextMusic = Val(LeerConfig("TeclaNextMusic", "66")) 'B
    cmbSCM.ListIndex = ShowCreditsMode
    
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    
    vsTamanoTapaPermitido = TamanoTapaPermitido
    
    Dim ModTec As Long
    ModTec = CLng(LeerConfig("IsMod46Teclas", "46"))
    If ModTec = 46 Then opModo4Teclas = True
    If ModTec = 5 Then opModo5Teclas = True
    tERR.Anotar "acmm"
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
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    PasarHoja = LeerConfig("PasarHoja", "1")
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    tERR.Anotar "acmn"
    Protector = LeerConfig("Protector", "1")
    CargarDuracionTemas = LeerConfig("CargarDuracionTemas", "0")
    MostrarRotulos = LeerConfig("MostrarRotulos", "1")
    RotulosArriba = LeerConfig("RotulosArriba", "0")
    DuracionProtect = LeerConfig("DuracionProtect", "180")
    RankToPeople = LeerConfig("RankToPeople", "1")
    TemasPorCredito = LeerConfig("TemasPorCredito", "1")
    CreditosBilletes = LeerConfig("CreditosBilletes", "10")
    PrecioBase = LeerConfig("PrecioBase", "0,50")
    PrecioBase2 = LeerConfig("PrecioBase2", "10")
    CreditosCuestaTema(0) = LeerConfig("CreditosCuestaTema", "1")
    CreditosCuestaTema(1) = LeerConfig("CreditosCuestaTema2", "0")
    CreditosCuestaTema(2) = LeerConfig("CreditosCuestaTema3", "0")
    
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "0")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "0")
    
    'validar cada X Creditos
    Validar = LeerConfig("Validar", "0")
    ValidarCada = LeerConfig("ValidarCada", "500")
    AvisarAntes = LeerConfig("AvisarAntes", "50")
    tERR.Anotar "acmo"
    MostrarTouch = LeerConfig("MostrarTouch", "0")
    'publicidades
    PUBs.HabilitarPublicidadesMp3Vid = LeerConfig("MostrarPub", "0")
    PUBs.HabilitarPublicidadesVMute = CBool(LeerConfig("MostrarPUBMute", "0"))
    PUBs.SonarPublicidadesCada = LeerConfig("PubliCada", "5")
    PUBs.HabilitarPublicidadesIMG = LeerConfig("MostrarPubIMG", "0")
    
    PUBs.SonarPublicidadesIMGCada = LeerConfig("PubliIMGCada", "10")
    IDIOMA = LeerConfig("Idioma", "Español")
    
    tERR.Anotar "acmp"
    
    
    'las variables ya se cargaron al inicio
    OpImgINI = CargarIMGinicio
    chkAutoReDraw = -AutoReDibuj
    'cargar la teckla que le corresponde a cada uno
    cmbTECLAS(0).ListIndex = FindIndexOfLst(CStr(TeclaDER) + " ", cmbTECLAS(0))
    cmbTECLAS(1).ListIndex = FindIndexOfLst(CStr(TeclaIZQ) + " ", cmbTECLAS(1))
    cmbTECLAS(2).ListIndex = FindIndexOfLst(CStr(TeclaOK) + " ", cmbTECLAS(2))
    cmbTECLAS(3).ListIndex = FindIndexOfLst(CStr(TeclaESC) + " ", cmbTECLAS(3))
    cmbTECLAS(4).ListIndex = FindIndexOfLst(CStr(TeclaNewFicha) + " ", cmbTECLAS(4))
    cmbTECLAS(5).ListIndex = FindIndexOfLst(CStr(TeclaConfig) + " ", cmbTECLAS(5))
    cmbTECLAS(6).ListIndex = FindIndexOfLst(CStr(TeclaPagAd) + " ", cmbTECLAS(6))
    cmbTECLAS(7).ListIndex = FindIndexOfLst(CStr(TeclaPagAt) + " ", cmbTECLAS(7))
    cmbTECLAS(8).ListIndex = FindIndexOfLst(CStr(TeclaCerrarSistema) + " ", cmbTECLAS(8))
    tERR.Anotar "acmq"
    cmbTECLAS(9).ListIndex = FindIndexOfLst(CStr(TeclaShowContador) + " ", cmbTECLAS(9))
    cmbTECLAS(10).ListIndex = FindIndexOfLst(CStr(TeclaPutCeroContador) + " ", cmbTECLAS(10))
    cmbTECLAS(11).ListIndex = FindIndexOfLst(CStr(TeclaFF) + " ", cmbTECLAS(11))
    cmbTECLAS(12).ListIndex = FindIndexOfLst(CStr(TeclaBajaVolumen) + " ", cmbTECLAS(12))
    cmbTECLAS(13).ListIndex = FindIndexOfLst(CStr(TeclaSubeVolumen) + " ", cmbTECLAS(13))
    cmbTECLAS(14).ListIndex = FindIndexOfLst(CStr(TeclaNextMusic) + " ", cmbTECLAS(14))
    cmbTECLAS(15).ListIndex = FindIndexOfLst(CStr(TeclaNewFicha2) + " ", cmbTECLAS(15))
    
    If LeerConfig("ActivarCorreccionSignal", "0") = "1" Then chkCS.Value = 1
    
    chkApagarPC = -ApagarAlCierre
    chkActivarERROR = LeerConfig("ActivarErr", "0")
    chkVerTiempoFaltante = -verTiempoRestante
    chkVerTemasPendientes = -verTemasEnLista
    chkVerCreditos = -verCreditos
    chkVerTotalDiscos = -verTOTdiscos
    chkVerPuestoRank = -verPuesto
    chkVerLista = -verLista
    chkDistorcionarTapas = -DistorcionarTapas
    tERR.Anotar "acmr"
    chkVisEXCL = -Is3pmExclusivo
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
    tERR.Anotar "acms"
    HSvolumen = VolumenIni
    HSVolumen2 = VolumenIni2
    LblVol = "Volumen: " + CStr(VolumenIni)
    lblVol2 = "Volumen2: " + CStr(VolumenIni2)
    txtEsperaTecla = EsperaTecla
    vsEsperaTecla = EsperaTecla
    VsPorcTema = PorcentajeTEMA
    vsSegFade = SegFade
    chkFastINI = -FASTini
    chkVUMeter = -HabilitarVUMetro
    chkVidFullScreen = -vidFullScreen
    chkSalida2 = -Salida2
    chkNoVumVID = -NoVumVID
    chkOutTemasWhenSel = -OutTemasWhenSel
    chkBloquearMusicaElegida = -BloquearMusicaElegida
    vsDiscosH = TapasMostradasH
    vsDiscosV = TapasMostradasV
    TeclaConfOK = "{UP}"
    TeclaConfESC = "{DOWN}"
    chkPasarhoja = -PasarHoja
    tERR.Anotar "acmt"
    If Protector = 0 Then chkNoProtector = True
    If Protector = 1 Then chkProtectOriginal = True
    If Protector = 2 Then chkProtectorCustom = True
    
    txtImgFondo = Trim(LeerConfig("ImgFondo", "NO"))
    If txtImgFondo <> "NO" Then
        chkLoadImgFondo = 1
    Else
        chkLoadImgFondo = 0
    End If
    
    txtImgFondo2 = Trim(LeerConfig("ImgFondo2", "NO"))
    If txtImgFondo2 <> "NO" Then
        chkLoadImgFondo2 = 1
    Else
        chkLoadImgFondo2 = 0
    End If
    tERR.Anotar "acmu"
    chkCargarDuracionTemas = -CargarDuracionTemas
    chkMostrarRotulos = -MostrarRotulos
    chkRotulosArriba = -RotulosArriba
    VSTemasXCredito = TemasPorCredito
    vsCreditosBilletes = CreditosBilletes
    txtPrecioBASE = PrecioBase
    'se pone al cambiar el precioBase
    'txtPrecioBase2 = PrecioBase2
    vsCreditosCuestaTema(0) = CreditosCuestaTema(0)
    vsCreditosCuestaTema(1) = CreditosCuestaTema(1)
    vsCreditosCuestaTema(2) = CreditosCuestaTema(2)
    
    vsCreditosCuestaTemaVIDEO(0) = CreditosCuestaTemaVIDEO(0)
    vsCreditosCuestaTemaVIDEO(1) = CreditosCuestaTemaVIDEO(1)
    vsCreditosCuestaTemaVIDEO(2) = CreditosCuestaTemaVIDEO(2)
    
    TxtUSUARIO = textoUsuario
    chkTouch = -MostrarTouch
    'validar cada X creditos
    chkValidar = -Validar
    vsValidarCada = ValidarCada
    vsAvisarAntes = AvisarAntes
    'publicidad
    ckPUB = -CLng(PUBs.HabilitarPublicidadesMp3Vid)
    chkVidMudos = -CLng(PUBs.HabilitarPublicidadesVMute)
    vsPubliCada = PUBs.SonarPublicidadesCada
    ckPubIMG = -CLng(PUBs.HabilitarPublicidadesIMG)
    vsPubliIMGCada = PUBs.SonarPublicidadesIMGCada
    cmbIDIOMA = IDIOMA
    tERR.Anotar "acmv"
    'mostrar el registro diario de contador
    Dim TE2 As TextStream
    Set TE2 = FSO.OpenTextFile(GPF("rdcday"), ForReading, False)
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
    tERR.Anotar "acmw"
    Command11_Click
    tERR.Anotar "acmx"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aclo"
    Resume Next
End Sub

Private Sub HSvolumen_Change()
    If frmIndex.MP3.IsPlaying(IAA) And CORTAR_TEMA(IAA) = False Then frmIndex.MP3.Volumen(IAA) = HSvolumen
    LblVol = "Volumen: " + Trim(CStr(HSvolumen))
End Sub

Private Sub HSvolumen_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineScroll.Visible = True
    HLP "Volumen del sonido actual."
End Sub

Private Sub HSvolumen_LostFocus()
    LineScroll.Visible = False
End Sub

Private Sub HSVolumen2_Change()
    If frmIndex.MP3.IsPlaying(IAA) And CORTAR_TEMA(IAA) Then frmIndex.MP3.Volumen(IAA) = HSVolumen2
    lblVol2 = "Volumen2: " + Trim(CStr(HSVolumen2))
End Sub

Private Sub HSVolumen2_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineScroll2.Visible = True
    HLP "Volumen del sonido para temas autoreproducidos."
End Sub

Private Sub HSVolumen2_LostFocus()
    LineScroll2.Visible = False
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
    "Arranque rápido. Recomendado para mas de 150 discos"
End Sub

Private Sub OpImgSIS_LostFocus()
    OpImgSIS.ForeColor = vbWhite
End Sub

Private Sub opModo4Teclas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    opModo4Teclas.ForeColor = vbYellow
    HLP "Configuración del teclado que no utiliza las flechas de desplazamiento" + _
        " vertical. El ESC sale del inteiror de los dicos y los mismos botones de" + _
        " desplazamiento sirven en el interior de los discos"
End Sub

Private Sub opModo4Teclas_LostFocus()
    opModo4Teclas.ForeColor = vbWhite
End Sub

Private Sub opModo5Teclas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    opModo5Teclas.ForeColor = vbYellow
    HLP "Configuración del teclado que si utiliza las flechas de desplazamiento" + _
        " vertical. El ESC no se utiliza, los botones de desplazamiento " + _
        "horizontal (Adel, Atras) salen del interior de los dicos y los " + _
        "mismos botones de desplazamiento vertical sirven en el interior de los discos"
End Sub

Private Sub opModo5Teclas_LostFocus()
    opModo5Teclas.ForeColor = vbWhite
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

Private Sub txtClaveAdmin_Change()
    Command31.Default = True
End Sub

Private Sub txtPrecioBASE_Change()
    'MsgBox KeyAscii
    If KeyAscii = 46 Then KeyAscii = 44
    
    'actualziar todo
    vsCreditosCuestaTema_Change 0
    vsCreditosCuestaTema_Change 1
    vsCreditosCuestaTema_Change 2
    
    vsCreditosCuestaTemaVIDEO_Change 0
    vsCreditosCuestaTemaVIDEO_Change 1
    vsCreditosCuestaTemaVIDEO_Change 2
    
    UpP2
End Sub

Private Sub UpP2() 'actualizar el precio 2
    
    Dim CB As Single 'creditos billetes
    CB = CSng(txtCreditosBilletes)
    
    Dim PB As Single 'precio base
    PB = CSng(txtPrecioBASE)
    
    Dim TC As Single '(temas por credito)
    TC = CSng(txtTemasXCredito)
    
    txtPrecioBase2 = CStr(Round((CB * PB) / TC, 2))
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

Private Sub vsCortaMusicaPaga_Change()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPorcTema.BackColor = vbYellow
    HLP "Cortar la musica paga "
End Sub

Private Sub vsCreditosBilletes_Change()
    txtCreditosBilletes = vsCreditosBilletes
    UpP2
End Sub

Private Sub vsCreditosCuestaTema_Change(Index As Integer)
    On Local Error Resume Next
    txtCreditosCuestaTema(Index) = vsCreditosCuestaTema(Index)
    
    txtPrecioM(Index) = FormatCurrency(CSng(txtCreditosCuestaTema(Index)) * _
        CSng(txtPrecioBASE) / CSng(txtTemasXCredito), , , , vbFalse)
    
End Sub

Private Sub vsCreditosCuestaTema_GotFocus(Index As Integer)
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCreditosCuestaTema(Index).BackColor = vbYellow
    HLP "Cantidad de creditos que se necesitan para ejecutar un tema. Si lo configura en dos necesitara" + _
    " dos creditos para poder ejecutar un tema"
End Sub

Private Sub vsCreditosCuestaTema_LostFocus(Index As Integer)
    txtCreditosCuestaTema(Index).BackColor = vbWhite
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_Change(Index As Integer)
    txtCreditosCuestaTemaVIDEO(Index) = vsCreditosCuestaTemaVIDEO(Index)
    On Local Error Resume Next
    txtPrecioV(Index) = FormatCurrency(CSng(txtCreditosCuestaTemaVIDEO(Index)) * CSng(txtPrecioBASE) / CSng(txtTemasXCredito), , , , vbFalse)
    
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_GotFocus(Index As Integer)
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCreditosCuestaTemaVIDEO(Index).BackColor = vbYellow
    HLP "Cantidad de creditos que se necesitan para ejecutar un " + _
    "clip de video musical. Si lo configura en dos necesitara" + _
    " dos creditos para poder ejecutar un clip de video"
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_LostFocus(Index As Integer)
    txtCreditosCuestaTemaVIDEO(Index).BackColor = vbWhite
End Sub

Private Sub VScroll2_Change()

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

Private Sub vsSegFade_Change()
    txtSegFade = vsSegFade
End Sub

Private Sub vsSegFade_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtSegFade.BackColor = vbYellow
    HLP "Segundos que tarda la cancion que esta terminando en 'irse' " + _
        "y la cancion que comienza en llegar al volumen normal"
End Sub

Private Sub vsSegFade_LostFocus()
    txtSegFade.BackColor = vbWhite
End Sub

Private Sub vsTamanoTapaPermitido_Change()
    txtTamanoTapaPermitido = vsTamanoTapaPermitido
End Sub

Private Sub vsTamanoTapaPermitido_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtTamanoTapaPermitido.BackColor = vbYellow
    HLP "Bloquear las imagenes para evitar sobrecargas cuando las " + _
        "imagenes superen los KiloBytes definidos aqui"
End Sub

Private Sub vsTamanoTapaPermitido_LostFocus()
    txtTamanoTapaPermitido.BackColor = vbWhite
End Sub

Private Sub VSTemasXCredito_Change()
    txtTemasXCredito = VSTemasXCredito
    
    'actualziar todo
    vsCreditosCuestaTema_Change 0
    vsCreditosCuestaTema_Change 1
    vsCreditosCuestaTema_Change 2
    
    vsCreditosCuestaTemaVIDEO_Change 0
    vsCreditosCuestaTemaVIDEO_Change 1
    vsCreditosCuestaTemaVIDEO_Change 2
    
    UpP2
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

Private Sub chkVisEXCL_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVisEXCL.ForeColor = vbYellow
    HLP "Cambia el modo de visualizacion a uno diferente. Solo disponible para usuarios " + _
        "de SuperLicencia. Esta visualizacion es exclusiva y cambia radicalmente la " + _
        "presentacion del software" + vbCrLf + "DEBE REINICIAR 3PM PARA VISUALIZAR EL CAMBIO"
End Sub

Private Sub chkVisEXCL_LostFocus()
    chkVisEXCL.ForeColor = vbWhite
End Sub

Private Sub cmbTECLAS_Click(Index As Integer)
    Dim SPL() As String
    SPL = Split(cmbTECLAS(Index), " ")
    txtTeclas(Index) = SPL(0)
End Sub
