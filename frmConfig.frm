VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfig 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuracion de 3PM"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5685
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   270
      TabIndex        =   0
      Top             =   360
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   6059
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   4210688
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Imagen"
      TabPicture(0)   =   "frmConfig.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "chkAutoReDraw"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Teclado"
      TabPicture(1)   =   "frmConfig.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblRIGHT"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblLEFT"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lblTeclaOK"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblTeclaESC"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblTeclaNEWficha"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(9)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Fichas"
      TabPicture(2)   =   "frmConfig.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(1)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label1(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label1(10)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "SCesperaMin"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "txtEsperaMin"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "SCtemasPorFicha"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "txtTemasPorFicha"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "SCmaxFichas"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "txtMaxFichas"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Cortes de luz"
      TabPicture(3)   =   "frmConfig.frx":0496
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(8)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "SLvolumen"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      Begin VB.TextBox txtMaxFichas 
         Height          =   285
         Left            =   -72180
         TabIndex        =   26
         Text            =   "9"
         Top             =   930
         Width           =   375
      End
      Begin VB.VScrollBar SCmaxFichas 
         Height          =   285
         Left            =   -71820
         TabIndex        =   25
         Top             =   930
         Width           =   285
      End
      Begin VB.TextBox txtTemasPorFicha 
         Height          =   285
         Left            =   -71760
         TabIndex        =   24
         Text            =   "1"
         Top             =   1560
         Width           =   375
      End
      Begin VB.VScrollBar SCtemasPorFicha 
         Height          =   285
         Left            =   -71400
         TabIndex        =   23
         Top             =   1560
         Width           =   285
      End
      Begin VB.TextBox txtEsperaMin 
         Height          =   285
         Left            =   -71730
         TabIndex        =   20
         Text            =   "10"
         Top             =   2580
         Width           =   375
      End
      Begin VB.VScrollBar SCesperaMin 
         Height          =   285
         Left            =   -71370
         TabIndex        =   18
         Top             =   2580
         Width           =   285
      End
      Begin VB.CheckBox chkAutoReDraw 
         Caption         =   "AutoReDibujado de pantalla"
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
         Left            =   -74010
         TabIndex        =   17
         Top             =   2850
         Width           =   3165
      End
      Begin VB.Frame Frame2 
         Caption         =   "En cortes de luz. Al reiniciarse el sistema..."
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
         Height          =   2085
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   4665
         Begin VB.OptionButton OpReIniSoloLista 
            Caption         =   "Omitir el tema que se estaba reproduciendo y respetar toda la lista de temas"
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
            Height          =   600
            Left            =   180
            TabIndex        =   9
            Top             =   870
            Width           =   4125
         End
         Begin VB.OptionButton OPreIniFull 
            Caption         =   "Volver a reproducir el tema que se estaba ejecutando y respetar la lista posterior."
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
            Height          =   600
            Left            =   180
            TabIndex        =   8
            Top             =   270
            Width           =   4035
         End
         Begin VB.OptionButton OpReIniNADA 
            Caption         =   "Eliminar toda la lista y no tocar ningun tema. NO RECOMENDADO"
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
            Height          =   600
            Left            =   180
            TabIndex        =   7
            Top             =   1410
            Width           =   4245
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Metodo de carga de imagenes"
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
         Height          =   1995
         Left            =   -74850
         TabIndex        =   1
         Top             =   690
         Width           =   4965
         Begin VB.OptionButton OPimgINICIO 
            Caption         =   $"frmConfig.frx":04B2
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
            Height          =   825
            Left            =   150
            TabIndex        =   3
            Top             =   300
            Width           =   4695
         End
         Begin VB.OptionButton OpImgApedido 
            Caption         =   "Las imagenes se cargan a pedido del usuario cada vez que cambia de pagina. RECOMENDADO para maquinas P II de 64 MB ram o superior"
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
            Height          =   735
            Left            =   210
            TabIndex        =   2
            Top             =   1170
            Width           =   4665
         End
      End
      Begin MSComctlLib.Slider SLvolumen 
         Height          =   450
         Left            =   210
         TabIndex        =   16
         Top             =   2820
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   794
         _Version        =   393216
         LargeChange     =   1000
         SmallChange     =   100
         Min             =   -10000
         Max             =   0
         SelStart        =   -1000
         TickStyle       =   1
         TickFrequency   =   100
         Value           =   -1000
         TextPosition    =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Temas a reproducir por cada ficha:"
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
         Height          =   255
         Index           =   10
         Left            =   -74730
         TabIndex        =   22
         Top             =   1620
         Width           =   3405
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Haga click sobre la ocion que desee modificar"
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
         Index           =   9
         Left            =   -74670
         TabIndex        =   21
         Top             =   630
         Width           =   4485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Minutos"
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
         Height          =   285
         Index           =   3
         Left            =   -71760
         TabIndex        =   19
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Volumen al iniciar el sistema"
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
         Height          =   285
         Index           =   8
         Left            =   300
         TabIndex        =   15
         Top             =   2640
         Width           =   2445
      End
      Begin VB.Label lblTeclaNEWficha 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tecla INSERTAR FICHAS:"
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
         Height          =   285
         Left            =   -74800
         TabIndex        =   14
         Top             =   2940
         Width           =   4500
      End
      Begin VB.Label lblTeclaESC 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tecla CANCELAR:"
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
         Height          =   285
         Left            =   -74800
         TabIndex        =   13
         Top             =   2490
         Width           =   4500
      End
      Begin VB.Label lblTeclaOK 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tecla OK:"
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
         Height          =   285
         Left            =   -74800
         TabIndex        =   12
         Top             =   2040
         Width           =   4500
      End
      Begin VB.Label lblLEFT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tecla desplamiento izquierda:"
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
         Height          =   285
         Left            =   -74800
         TabIndex        =   11
         Top             =   1560
         Width           =   4500
      End
      Begin VB.Label lblRIGHT 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tecla desplamiento derecha:"
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
         Height          =   285
         Left            =   -74800
         TabIndex        =   10
         Top             =   1080
         Width           =   4500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Maximo de fichas permitidos:"
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
         Height          =   285
         Index           =   0
         Left            =   -74760
         TabIndex        =   5
         Top             =   960
         Width           =   2625
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo de espera sin ninguna carga de temas antes de ejecutar automaticamnete un tema:"
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
         Height          =   735
         Index           =   1
         Left            =   -74760
         TabIndex        =   4
         Top             =   2340
         Width           =   4785
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
