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
   Begin VB.PictureBox pVU2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6240
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   34
      Top             =   960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pVU4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   5940
      ScaleHeight     =   975
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   3060
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pVU3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6090
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   32
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picFondo2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   240
      ScaleHeight     =   825
      ScaleWidth      =   6270
      TabIndex        =   25
      Top             =   120
      Width           =   6270
      Begin VB.Line LineRitmo 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   5
         X1              =   4170
         X2              =   4740
         Y1              =   120
         Y2              =   120
      End
      Begin VB.Line LineLETRA 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   5580
         X2              =   6150
         Y1              =   300
         Y2              =   300
      End
      Begin VB.Label lLETRAS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Index           =   0
         Left            =   30
         TabIndex        =   29
         Top             =   390
         UseMnemonic     =   0   'False
         Width           =   270
      End
      Begin VB.Label lRITMO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lRitmo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   28
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label6 
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
         Left            =   7080
         TabIndex        =   27
         Top             =   2700
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label5 
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
         Left            =   10230
         TabIndex        =   26
         Top             =   2940
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lRITMO2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lRitmo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   30
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   1035
      End
      Begin VB.Label lLETRAS2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   495
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   390
         UseMnemonic     =   0   'False
         Width           =   270
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   10350
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Left            =   10350
      Top             =   5010
   End
   Begin VB.PictureBox frDiscos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   420
      ScaleHeight     =   3975
      ScaleWidth      =   5655
      TabIndex        =   21
      Top             =   1050
      Width           =   5655
      Begin VB.PictureBox picFondoDisco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   3315
         Left            =   60
         ScaleHeight     =   3315
         ScaleWidth      =   5505
         TabIndex        =   22
         Top             =   270
         Width           =   5505
         Begin VB.Image imgSALIR 
            Height          =   375
            Left            =   3420
            Top             =   2250
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image imgSELEC 
            Height          =   375
            Left            =   1530
            Top             =   2280
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblNOCREDIT 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "CREDITO INSUFICIENTE"
            BeginProperty Font 
               Name            =   "Trebuchet MS"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   330
            Left            =   1350
            TabIndex        =   41
            Top             =   1770
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.Image cmdTouchAbajo 
            Height          =   360
            Left            =   330
            Top             =   690
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.Image cmdTouchArriba 
            Height          =   360
            Left            =   660
            Top             =   690
            Visible         =   0   'False
            Width           =   300
         End
         Begin VB.Label lblDATA2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "data.txt"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   225
            Left            =   2760
            TabIndex        =   40
            Top             =   2520
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblDATA 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Data.txt"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   2700
            TabIndex        =   39
            Top             =   2340
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblCanciones 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de canciones"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   300
            Index           =   0
            Left            =   1860
            TabIndex        =   37
            Top             =   960
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Image imgFondoDiscoSel 
            Height          =   645
            Left            =   270
            Top             =   1110
            Width           =   825
         End
         Begin VB.Image imgListaSong 
            Height          =   1455
            Left            =   1230
            Top             =   600
            Width           =   2775
         End
         Begin VB.Image imgDiscoSEL 
            Height          =   495
            Left            =   330
            Top             =   1170
            Width           =   675
         End
         Begin VB.Image TapaCD 
            Height          =   465
            Index           =   0
            Left            =   4560
            Stretch         =   -1  'True
            Top             =   300
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image imageFONDO 
            Height          =   660
            Index           =   0
            Left            =   4440
            Stretch         =   -1  'True
            Top             =   180
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label lblDisco 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "complete hoja"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Index           =   0
            Left            =   930
            TabIndex        =   23
            Top             =   2850
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label lblDisco2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Complete al menos la primera hoja de discos cargados"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   435
            Index           =   0
            Left            =   330
            TabIndex        =   24
            Top             =   2820
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label lblDiscoSEL 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Complete al menos la primera hoja de discos cargados"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   435
            Left            =   0
            TabIndex        =   35
            Top             =   30
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label lblDiscoSEL2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Complete al menos la primera hoja de discos cargados"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C0C0C0&
            Height          =   435
            Left            =   60
            TabIndex        =   36
            Top             =   60
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label lblCanciones2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de canciones"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   300
            Index           =   0
            Left            =   1890
            TabIndex        =   38
            Top             =   960
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   1710
         End
      End
   End
   Begin VB.PictureBox picFondo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1590
      Left            =   180
      ScaleHeight     =   1590
      ScaleWidth      =   8790
      TabIndex        =   7
      Top             =   7230
      Width           =   8790
      Begin tbr3pm.txtRolling RollCRED 
         Height          =   1335
         Left            =   870
         TabIndex        =   18
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   2355
      End
      Begin tbr3pm.txtRolling RollSONG 
         Height          =   1155
         Left            =   6210
         TabIndex        =   19
         Top             =   300
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   2037
      End
      Begin tbr3pm.tbrPassImg tbrPassImg1 
         Height          =   1260
         Left            =   3300
         TabIndex        =   11
         Top             =   300
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   2223
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
            Left            =   120
            TabIndex        =   12
            Top             =   150
            Visible         =   0   'False
            Width           =   1845
         End
      End
      Begin VB.Label lblCreditos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Credito $ 15000,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2550
         TabIndex        =   9
         Top             =   150
         Width           =   2235
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
         Left            =   10230
         TabIndex        =   10
         Top             =   2940
         Visible         =   0   'False
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
         Left            =   7080
         TabIndex        =   8
         Top             =   2700
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Credito $ 15000,00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   2490
         TabIndex        =   20
         Top             =   30
         Width           =   2235
      End
   End
   Begin VB.TextBox txtS3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   10800
      TabIndex        =   17
      Top             =   4500
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox pVU1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6180
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   16
      Top             =   1620
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   8010
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   15
      Top             =   5430
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4020
      IntegralHeight  =   0   'False
      ItemData        =   "frmINDEX.frx":08CA
      Left            =   1680
      List            =   "frmINDEX.frx":090A
      TabIndex        =   14
      Top             =   3450
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   7680
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   13
      Top             =   5430
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Frame frModoVideo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   7080
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   3500
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
         Left            =   30
         TabIndex        =   1
         Top             =   0
         Width           =   2625
      End
   End
   Begin VB.Frame frTEMAS 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   7080
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   3500
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
         TabIndex        =   4
         Top             =   0
         Width           =   1245
      End
   End
   Begin VB.Image cmdTouchAbajo2 
      Height          =   375
      Left            =   6240
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdTouchArriba2 
      Height          =   375
      Left            =   5820
      Top             =   4770
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgSelec2 
      Height          =   375
      Left            =   5820
      Top             =   5340
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdPagAt 
      Height          =   615
      Left            =   6690
      Top             =   4350
      Width           =   735
   End
   Begin VB.Image cmdPagAd 
      Height          =   615
      Left            =   7560
      Top             =   4350
      Width           =   855
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
      Left            =   8100
      TabIndex        =   6
      Top             =   330
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.Label lblTEMAS 
      Alignment       =   2  'Center
      BackColor       =   &H00008080&
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
      Height          =   225
      Left            =   7080
      TabIndex        =   5
      Top             =   2670
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label lblModoVideo 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
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
      Height          =   225
      Left            =   7080
      TabIndex        =   2
      Top             =   1050
      Visible         =   0   'False
      Width           =   3495
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VU As tbrSoftVumetro.tbrDrawVUM
Public WithEvents MP3 As tbrPlayer.MainPlayer
Attribute MP3.VB_VarHelpID = -1
Dim YaEsoySaliendoGrat_Cortar(3) As Boolean
Dim LastRetEmpezarSig As Long 'guardo el ultimo valor de empezar siguiente
'para el caso unico de que termine un video y no siga un video y se necesite
'que empieze una publicidad

Dim ModoVideoSelTema As Boolean 'si estoy en video
'saber si estoy eligiendo tema. Sino estoy en disco

Dim TemaElegidoModoVideo As Integer

Dim LastDiscoSel As Long
Dim DiscosEnPagina As Long

Dim VolBajando As Double 'bajando volumen para terminar tema demo
Dim LastpSeconds As Long 'comparador para bajar de a uno el volumen en demos

Dim WithEvents TF As tbrFOCUS.clsFOCUS
Attribute TF.VB_VarHelpID = -1

'me cago en la mierda. Siguen dos canciones al mismo tiempo !!!
Dim IenPlenaCancion(3) As Long 'cada uno de los hilos de ejecucion
'solo uno puede estar activo!
'=0 sin nada
'=1 menor a segFade, comenzando cancion
'=2 en plena cancion despues de 1 y antes de 3
'=3 en los segundos finales bajando el volumen

Dim WithEvents GK As tbrGetKeys
Attribute GK.VB_VarHelpID = -1
Private EstoyEnDisco As Long 'me dice si estoy dentro de un disco en el modo nuevo
Private OkInState1 As Long 'presiones de la tecla ok en el modo SuperSel
'esto para ignorar la primera siempre

Public Function PonerFoco()
    TF.PonerFoco
End Function

Private Function EnQueFilaEstoy() As Long
    'es la fila uno si es la primera
    'la barra invertida devuelve solo la parte entera!!!
    EnQueFilaEstoy = (nDiscoSEL \ TapasMostradasH) + 1
    tERR.Anotar "acaa", nDiscoSEL, TapasMostradasH
End Function

'Private Sub cmdDiscoAd_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acab", KeyCode, Shift
'End Sub

'Private Sub cmdDiscoAt_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acac", KeyCode, Shift
'End Sub

Private Sub cmdPagAd_Click()
    If MostrarTouch Then
        tERR.Anotar "acam", EstoyEnDisco
        If EsVideo And Salida2 = False Then
            Form_KeyDown TeclaDER, 0
        Else
            Form_KeyDown TeclaPagAd, 0
        End If
    End If
End Sub

Private Sub cmdPagAd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchderechaapretado")
    cmdPagAd.Picture = LoadPicture(imF)
End Sub

'Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acad", KeyCode, Shift
'End Sub

Private Sub cmdPagAd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchderechanormal")
    cmdPagAd.Picture = LoadPicture(imF)
End Sub

Private Sub cmdPagAt_Click()
    If MostrarTouch Then
        tERR.Anotar "acak", EstoyEnDisco
        'si tengo videos en la pantalla de la pc no paso pagina, paso solo disco
        If EsVideo And Salida2 = False Then
            Form_KeyDown TeclaIZQ, 0
        Else
            Form_KeyDown TeclaPagAt, 0
        End If
    End If
End Sub

'Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acae", KeyCode, Shift
'End Sub

Private Sub cmdPagAt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchizqapretado")
    cmdPagAt.Picture = LoadPicture(imF)
End Sub

Private Sub cmdPagAt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchizqnormal")
    cmdPagAt.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchAbajo_Click()
    selDiscoI -1
End Sub

Private Sub cmdTouchAbajo2_Click()
    Form_KeyDown TeclaDER, 0
End Sub

Private Sub cmdTouchArriba_Click()
    selDiscoI -2
End Sub

Private Sub cmdTouchArriba_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("tocuharribaelegido")
    cmdTouchArriba.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchArriba_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("tocuharribacomun")
    cmdTouchArriba.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchAbajo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchabajoelegido")
    cmdTouchAbajo.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchAbajo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchabajocomun")
    cmdTouchAbajo.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchArriba2_Click()
    Form_KeyDown TeclaIZQ, 0
End Sub

Private Sub cmdTouchArriba2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("tocuharribaelegido")
    cmdTouchArriba2.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchArriba2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("tocuharribacomun")
    cmdTouchArriba2.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchAbajo2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchabajoelegido")
    cmdTouchAbajo2.Picture = LoadPicture(imF)
End Sub

Private Sub cmdTouchAbajo2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("touchabajocomun")
    cmdTouchAbajo2.Picture = LoadPicture(imF)
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
    
    
    lblPrecios2 = GetPrecios(ShowCreditsMode, " / ")
    
    'acomodo el roll!
    RollCRED.ReplaceIndex 1, GetPrecios(ShowCreditsMode, vbCrLf)
    
    Exit Sub
regERR:
    tERR.Anotar "errACAP"
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acap"
End Sub

Public Sub StartVu(sModo As String) 'empezar a medir sonido

    tERR.Anotar "SV01", sModo
    
    'ver si no hubiera que iniciarlo!
    If HabilitarVUMetro = False Then Exit Sub
    If NoVumVID And EsVideo And (Salida2 = False) And vidFullScreen Then Exit Sub

    Select Case sModo
        Case "grande" 'modo comun a los costados grande
            AnchoBarra = 840
            VU.ModoVumetro = TresColoresEstereo ' TresColoresEstereo
            pVU1.Left = 0
            pVU1.Width = AnchoBarra
            pVU1.Height = frDiscos.Height / 2
            pVU1.Top = picFondo2.Height + frDiscos.Height / 4
            pVU2.Left = Me.Width - (AnchoBarra + 25)
            If EstoyEnModoVideoMiniSelDisco Then pVU2.Left = pVU2.Left - frModoVideo.Width
            
            pVU2.Width = AnchoBarra
            pVU2.Height = pVU1.Height
            pVU2.Top = pVU1.Top
            
            pVU3.Visible = False
            pVU4.Visible = False
        Case "custom"
            'AnchoBarra = pVU1.Width
            VU.ModoVumetro = UnaImagenSobreOtra 'imagenes prendido apagado
            pVU1.Left = 0
            pVU1.Top = frDiscos.Top
            
            If EstoyEnModoVideoMiniSelDisco Then
                pVU3.Left = frModoVideo.Left - pVU3.Width
            Else
                pVU3.Left = Me.Width - pVU2.Width
            End If
            'pVU2.Width = AnchoBarra
            pVU3.Top = pVU1.Top
            'pVU2.Height = pVU1.Height
            pVU2.Top = pVU1.Top
            pVU2.Left = pVU1.Left
            
            pVU4.Top = pVU3.Top
            pVU4.Left = pVU3.Left
            
            pVU3.Visible = True
            pVU4.Visible = True
    End Select
    
    'pVU1.BackColor = vbBlack
    'pVU2.BackColor = vbBlack
    'pVU3.BackColor = vbBlack
    'pVU4.BackColor = vbBlack
    
    'a veces no hay que mostrar!
    If sModo = "grande" Then
        If (EsVideo And Salida2 = False And NoVumVID) Or (HabilitarVUMetro = False) Then
            pVU1.Visible = False
            pVU2.Visible = False
            If VU.IsPlaying Then VU.Terminar
        Else
            pVU1.Visible = True
            pVU2.Visible = True
            pVU1.ZOrder
            pVU2.ZOrder
            If VU.IsPlaying = False Then
                If VU.Empezar = 1 Then
                    tERR.AppendLog "No empieza el vumetro!!"
                End If
            End If
        End If
    Else
        pVU1.Visible = True
        pVU2.Visible = True
        pVU1.ZOrder
        pVU2.ZOrder
        If VU.IsPlaying = False Then
            If VU.Empezar = 1 Then
                tERR.AppendLog "No empieza el vumetro!!"
            End If
        End If
    End If
    
    If HabilitarVUMetro Then VU.NotifyResizeVUM
    
End Sub

Private Sub ProcessKeyCoin(Tecla As Integer, isDown As Long)
    'isDown puede ser
    '0 es up
    '1 es down
    '2 viene de la api que no sabe
    lblNOCREDIT.Visible = False
    '***********************************************************
    'si es 0 o 1 y yo suo los 2 ignorar para que no duplique!!!!
    If GK.IsLisen Then
        If isDown = 0 Or isDown = 1 Then Exit Sub
    End If
    '***********************************************************
    Select Case Tecla
        Case TeclaNewFicha
            If isDown = 1 Then
                'si TeclaOk=KeyDown entonces no lo hace aca
                If FindParam3PM("to") = "kd" Then
                    LTE 1
                    VarCreditos CSng(TemasPorCredito)
                End If
            End If
            
            If isDown = 0 Then
                'si TeclaOk=KeyDown entonces no lo hace aca
                If FindParam3PM("to") = "999999" Then
                    LTE 1
                    VarCreditos CSng(TemasPorCredito)
                End If
            End If
            
            If isDown = 2 Then
                LTE 1
                VarCreditos CSng(TemasPorCredito)
            End If
            
        Case TeclaNewFicha2
            If isDown = 1 Then
                'si TeclaOk2=KeyDown entonces no lo hace aca
                If FindParam3PM("to2") = "kd" Then
                    LTE 2
                    VarCreditos CSng(CreditosBilletes)
                End If
            End If
            
            If isDown = 0 Then
                'si TeclaOk2=KeyDown entonces no lo hace aca
                If FindParam3PM("to2") = "999999" Then
                    LTE 2
                    VarCreditos CSng(CreditosBilletes)
                End If
            End If
            
            If isDown = 2 Then
                LTE 2
                VarCreditos CSng(CreditosBilletes)
            End If
            
    End Select
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
        Case TeclaNewFicha
            ProcessKeyCoin TeclaNewFicha, 1
        Case TeclaNewFicha2
            ProcessKeyCoin TeclaNewFicha2, 1
        Case vbKeyF1
            frmERRORES.Show 1
        Case vbKeyF2
            tERR.AppendLog "USR_PRES_F2"
        Case vbKeyF3
            frmREG2.Show 1
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
                End
            End If
        Case vbKeyF5
            my_MEM.SetMomento "Apreto F5"
            tERR.AppendSinHist "F5: " + vbCrLf + my_MEM.GetFullDetalles
        Case TeclaShowContador
            frmOnlyContador.Show 1
        Case TeclaPutCeroContador
            SumarContadorCreditos -CONTADOR 'esto lo deja en cero
            frmOnlyContador.Show 1
        Case TeclaFF 'avanzar 10 segundos
            If EnableFF Then
                Dim ToSec As Long
                tERR.Anotar "acav", Shift
                If Shift = 1 Then
                    ToSec = (MP3.PositionInSec(3) * 1000) + 10000
                    MP3.SeekTo CStr(ToSec), 3
                Else
                    ToSec = (MP3.PositionInSec(IAA) * 1000) + 10000
                    MP3.SeekTo CStr(ToSec), IAA
                End If
                
            End If
        'subir o bajar volumen
        Case TeclaBajaVolumen
            If frmIndex.MP3.IsPlaying(IAA) Then
                If CORTAR_TEMA(IAA) = False Then 'TEMA PAGO
                    If VolumenIni <= 5 Then
                        frmIndex.MP3.Volumen(IAA) = 0
                    Else
                        frmIndex.MP3.Volumen(IAA) = VolumenIni - 5
                    End If
                    VolumenIni = frmIndex.MP3.Volumen(IAA)
                Else 'TEMA GRATUITO VARIA VOLUMEN 2
                    If VolumenIni2 <= 5 Then
                        frmIndex.MP3.Volumen(IAA) = 0
                    Else
                        frmIndex.MP3.Volumen(IAA) = VolumenIni2 - 5
                    End If
                    VolumenIni2 = frmIndex.MP3.Volumen(IAA)
                End If
            End If
        Case TeclaSubeVolumen
            If frmIndex.MP3.IsPlaying(IAA) Then
                If CORTAR_TEMA(IAA) = False Then 'TEMA PAGO
                    If VolumenIni >= 95 Then
                        frmIndex.MP3.Volumen(IAA) = 100
                    Else
                        frmIndex.MP3.Volumen(IAA) = VolumenIni + 5
                    End If
                    VolumenIni = frmIndex.MP3.Volumen(IAA)
                Else 'TEMA GRATUITO
                    If VolumenIni2 >= 95 Then
                        frmIndex.MP3.Volumen(IAA) = 100
                    Else
                        frmIndex.MP3.Volumen(IAA) = VolumenIni2 + 5
                    End If
                    VolumenIni2 = frmIndex.MP3.Volumen(IAA)
                End If
            End If
        
        Case TeclaPagAd
            'pase lo que pase registrar
            TECLAS_PRES = TECLAS_PRES + "5"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            If EstoyEnDisco = 0 Then
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
                    'ACA LE PUSE >= esra solo=!!!!
                    If PrimeroDeLaPaginaQueSigue >= TOTAL_DISCOS Then
                        PrimeroDeLaPaginaQueSigue = 0
                    End If
                    'supongo que lo puse para que no desseleccione el mismo _
                        que va a seleccionar???
                    If nDiscoSEL <> 0 Then UnSelDisco nDiscoSEL
                    tERR.Anotar "acba", nDiscoSEL
                    DiscosEnPagina = CargarDiscos(PrimeroDeLaPaginaQueSigue, True, 1)
                    'lblTOTdiscos = "Disco " + CStr(PrimeroDeLaPaginaQueSigue + 1) + " de " + CStr(TOTAL_DISCOS)
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
            End If
            'si estoy con el touch lo uso para mover el tema elegido
            If EstoyEnDisco = 1 Then
                selDiscoI -1
            End If
        Case TeclaPagAt
            TECLAS_PRES = TECLAS_PRES + "6"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            If EstoyEnDisco = 0 Then
            
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
                    'lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
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
            End If
            
            'si estoy con el touch lo uso para mover el tema elegido
            If EstoyEnDisco = 1 Then
                selDiscoI -2
            End If
        Case TeclaConfig
             
             frmConfig.Show 1
        Case TeclaIZQ
            If EstoyEnDisco = 0 Then
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
                'lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
FinTeclaZ:
                TECLAS_PRES = TECLAS_PRES + "1"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                lblTECLAS = TECLAS_PRES
            End If
            
            If EstoyEnDisco = 1 Then
                If selDiscoI(-2) = -99 Then 'todo esta elegido!
                    UnSuperSel
                End If
            End If
                
        Case TeclaDER
            'esta tecla es DER en el modo 46 pasandpo de abajo a arriba
            'y siguiendo a la atras ¿? sig en el modo 5
            If EstoyEnDisco = 0 Then
            
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
                
                'lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
                TECLAS_PRES = TECLAS_PRES + "2"
                TECLAS_PRES = Right(TECLAS_PRES, 20)
                lblTECLAS = TECLAS_PRES
                
            End If
            
            If EstoyEnDisco = 1 Then
                If selDiscoI(-1) = -99 Then 'todo esta elegido!
                    UnSuperSel
                End If
            End If
            
        Case TeclaCerrarSistema
            YaCerrar3PM
        Case TeclaESC
            tERR.Anotar "acdo"
            TECLAS_PRES = TECLAS_PRES + "4"
            TECLAS_PRES = Right(TECLAS_PRES, 20)
            lblTECLAS = TECLAS_PRES
            
            If EstoyEnDisco = 0 Then
            
                If ModoVideoSelTema Then 'esta eligiendo canciones dentro del disco
                    AcomodarModoTexto 1
                    ModoVideoSelTema = False 'ya no esta mas!!
                End If
            End If
            
            'solo dejar todo como estaba!
            If EstoyEnDisco = 1 Then UnSuperSel
                
        Case vbKeyF12
            MostrarCursor True
    End Select
    
FinKD:
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    
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
    If RealKeyCode = TeclaNextMusic Then
        tERR.Anotar "acaw", EnableNextMusic
        If EnableNextMusic Then
            EMPEZAR_SIGUIENTE 2
        Else
            'esta en (fade) segundos del inicio
            'EMPEZAR_SIGUIENTE 2
        End If
    End If
            
    If RealKeyCode = TeclaNewFicha Then
        ProcessKeyCoin TeclaNewFicha, 0
    End If
    
    If RealKeyCode = TeclaNewFicha2 Then
        ProcessKeyCoin TeclaNewFicha2, 0
    End If
    
    If RealKeyCode = TeclaOK Then
        TECLAS_PRES = TECLAS_PRES + "3"
        TECLAS_PRES = Right(TECLAS_PRES, 20)
        lblTECLAS = TECLAS_PRES
        
        If EstoyEnDisco = 0 Then
            'si estoy en video
            'saber si estoy eligiendo tema. Si no estoy en disco
            tERR.Anotar "accv", nDiscoGral, nDiscoSEL, ModoVideoSelTema
            If ModoVideoSelTema Then
                'si esta en fullscreen NO EJECUTAR!!!
                'solo si no sale por la segunda salida!!!
                If EsVideo And vidFullScreen And Salida2 = False Then GoTo FinUP 'fin keydown
                'si no dice salir cargar tema
                tERR.Anotar "accw", T(TemaElegidoModoVideo)
                If T(TemaElegidoModoVideo) = "SALIR" Or T(TemaElegidoModoVideo) = "No hay temas" Then
                    'volver a elegir discos
                    AcomodarModoTexto 1
                    ModoVideoSelTema = False
                Else 'else del "SALIR"
                    'ejecutar el tema
                    'ANTES DE VER CUANTOS CREDITOS NECESITA TENGO QUE SABER SI QUIERE EJECUTAR
                    'MP3 O VIDEO!!!!!!
                    Dim temaElegido As String
                    'lstext es una lista oculta  con datos completos
                    temaElegido = txtInLista(MATRIZ_TEMAS(TemaElegidoModoVideo), 0, "#")
                    tERR.Anotar "accx", temaElegido, CREDITOS
                    
                    Dim S36 As Long
                    S36 = TrataEjecutarTema(temaElegido)
                    tERR.Anotar "accx", S36
                    If S36 = 2 Then 'ya esta ejecutando otra cosa!
                        'volver a elegir discos
                        AcomodarModoTexto 1
                        ModoVideoSelTema = False
                    End If
                End If
            
            Else 'ELSE DEL MODOVIDEO SEL TEMA
            
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                tERR.Anotar "acdd"
                'ver si es video debería desplegar los temas del disco elegido
                'en modo de texto
                'pero si estoy viendo el video en salida2 es video sera verdadero
                'pero de todas formas no veo als lista de texto y sigo igual
                'solo si esvideo y necesito el modo texto del video!!!!
                If EsVideo And Salida2 = False Then
                    AcomodarModoTexto 2
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
                    
                Else 'ELSE DEL ESVIDEO AND SALIDA2
                    'If lblDISCO(nDiscoSEL) = "_Los mas escuchados" Then GoTo TOP10Show
                    tERR.Anotar "acdm", lblDISCO(nDiscoSEL), nDiscoSEL, nDiscoGral
                    SuperSel nDiscoSEL
                End If
            End If
        End If
        
        If EstoyEnDisco = 1 Then
            OkInState1 = OkInState1 + 1 'la primera no va!
            If OkInState1 > 1 Then EjecutarDeTouch
        End If
    End If
    
FinUP:
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    
    Exit Sub


'TOP10Show:
'    'ACA ENTRA AL FEO RANKING MEJORAR!!
'    'XXXXX
'    tERR.Anotar "acdq"
'    FRMTOP10.Show 1
'    Exit Sub


FallaKD:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdr"
    Resume Next

End Sub

Private Sub AcomodarModoTexto(lModo As Long)
    'acomodar la parte de los textos segun sea necesario
    'los modos de ingreso pueden ser:
    '1: Lista de discos para elegir (=true)
    '2: Se entro a un disco (modivideoseltema=true)
    
    'cada uno tendra variaciones segun si esta en modo touch o no
    '*********************************
    tERR.Anotar "acdp", nDiscoSEL, nDiscoGral
    
    '*********************************
    'hay cosas universales (la parte de arriba)
    If MostrarTouch Then
        
    Else
        
    End If

    ' ... y otras que dependen de lo que se muestra
    
    If lModo = 1 Then
        If MostrarTouch Then
            'muestro las 2 flechas verticales y el "selecconar" arriba cerca del lblmodovideo
            imgSelec2.Top = picFondo2.Top + picFondo2.Height
            imgSelec2.Left = lblModoVideo.Left + (lblModoVideo.Width / 2 - imgSelec2.Width / 2)
            cmdTouchArriba2.Top = imgSelec2.Top + imgSelec2.Height + 60
            cmdTouchAbajo2.Top = cmdTouchArriba2.Top
            cmdTouchAbajo2.Left = lblModoVideo.Left + _
                (lblModoVideo.Width / 2 - cmdTouchAbajo2.Width) - 60
            cmdTouchArriba2.Left = lblModoVideo.Left + (lblModoVideo.Width / 2) + 60
            lblModoVideo.Top = cmdTouchArriba2.Top + cmdTouchArriba2.Height
        Else
            lblModoVideo.Top = picFondo2.Top + picFondo2.Height
        End If
        
        frModoVideo.Top = lblModoVideo.Top + lblModoVideo.Height
        frModoVideo.Height = frDiscos.Height - 200 - (lblModoVideo.Top + lblModoVideo.Height) + picFondo2.Height
        
        frTEMAS.Visible = False
        lblTEMAS.Visible = False
        imgSelec2.Visible = MostrarTouch
        cmdTouchArriba2.Visible = MostrarTouch
        cmdTouchAbajo2.Visible = MostrarTouch
        
        UnSelTema 0 'desmarca para que cuando se cargue sea en formato original
    End If
    
    If lModo = 2 Then
        'cosas indistintas
        lblModoVideo.Top = picFondo2.Top + picFondo2.Height
        frModoVideo.Height = frDiscos.Height / 5
        frModoVideo.Top = lblModoVideo.Top + lblModoVideo.Height
        
        If MostrarTouch Then
            'muestro las 2 flechas verticales y el "selecconar" arriba cerca del lblTEMAS
            imgSelec2.Top = frModoVideo.Top + frModoVideo.Height + 60
            imgSelec2.Left = lblModoVideo.Left + (lblModoVideo.Width / 2 - imgSelec2.Width / 2)
            cmdTouchArriba2.Top = imgSelec2.Top + imgSelec2.Height + 60
            cmdTouchAbajo2.Top = cmdTouchArriba2.Top
            cmdTouchAbajo2.Left = lblModoVideo.Left + _
                (lblModoVideo.Width / 2 - cmdTouchAbajo2.Width) - 60
                
            cmdTouchArriba2.Left = lblModoVideo.Left + (lblModoVideo.Width / 2) + 60
            
            lblTEMAS.Top = cmdTouchArriba2.Top + cmdTouchArriba2.Height
            lblTEMAS.Left = lblModoVideo.Left
            frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
            frTEMAS.Height = frDiscos.Height - 200 - (lblTEMAS.Height + lblTEMAS.Top) + picFondo2.Height
        Else
            frModoVideo.Height = frDiscos.Height / 5
            lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
            lblTEMAS.Left = lblModoVideo.Left
            frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
            frTEMAS.Height = frDiscos.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
        End If
        
        OrdenarListaModoVideo 'asegurarme que el disco elegido se ve en la lista
        'este ultimo depende de frModoVideo.Height por lo que sirve para cualquier caso
    
        imgSelec2.Visible = MostrarTouch
        cmdTouchArriba2.Visible = MostrarTouch
        cmdTouchAbajo2.Visible = MostrarTouch
    
        lblTEMAS.Visible = True
        frTEMAS.Visible = True
    End If
    
End Sub


Private Sub EjecutarDeTouch()

    On Local Error GoTo errDTouch
    tERR.Anotar "caaa"
    Dim Fg As Long
    Fg = selDiscoI(-3) 'es el numero de cancion!
    tERR.Anotar "caab", Fg
    Dim S37 As Long
    S37 = TrataEjecutarTema(lblCanciones(Fg).Tag)
    tERR.Anotar "caac", S37
    If S37 = 1 Then 'si no alcanza el credito avisar!
        lblNOCREDIT.Visible = True
        Exit Sub
    End If
    
    'o algo se ejecuta o va a la lista seguro
    If BloquearMusicaElegida Then
        lblCanciones(Fg).Visible = False
        lblCanciones(Fg).Tag = "" 'para que no lo elija de nuevo
    Else 'de alguna forma tengo que decirle que se eligio!
        'ya se corre al sigiuente solo
        lblCanciones(Fg).Font.Italic = True
    End If
    
    'elegir el que sigue!
    If selDiscoI(-1) = -99 Then 'todo esta elegido!
        UnSuperSel
    End If
    
    If OutTemasWhenSel Or (S37 = 3 And Salida2 = False) Then UnSuperSel
    
    Exit Sub
    
errDTouch:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdu5"
    Resume Next
End Sub

Public Function GetIntervalS3() As Long
    GetIntervalS3 = S3.GetInterval
End Function

Public Sub SetIntervalS3(NewIntervalS3 As Long)
    S3.SetInterval NewIntervalS3
End Sub

Private Sub Form_Load()
    On Error GoTo NoLoadIndex
    
    'no puedo hacer referencia a ningun objeto de frmIndex por que lo cargaria antes de tiempo
    imF = ExtraData.GetImagePath("vumetroprendido")
    'temporalmente uso pVu1 pero puede ser cualquiera es solo por que no se cuanto tiene de ancho la imagen segun el skin
    pVU1.AutoSize = True
    pVU1.Picture = LoadPicture(imF)
    AnchoBarra = pVU1.Width
    pVU1.Picture = LoadPicture
    pVU1.AutoSize = False
    
    tERR.Anotar "sVU01"
    Me.BackColor = vbBlack
    'dejo todo definido en el vumetro
    
    'no se activa escuchar por el puerto si no esta configurado
    If LeerConfig("UsarS3", "0") = "1" Then
        Set S3 = New tbrSKS3.clsTbrSKS3
        '*************************
        S3.HwndMsg = txtS3.hwnd
        S3.Prender
        
        S3.SetInterval CLng(LeerConfig("FrecTecladoTBR", "50"))
        '*************************
    End If
        
    Set VU = New tbrSoftVumetro.tbrDrawVUM
    Dim UAT As String
    
    UAT = LeerConfig("UseAPITecla", "0")
    'la declaro para que pueda saberse el Islisen que es necesario !!!!
    Set GK = New tbrGetKeys
    If UAT <> "0" Then
        'ver que letras necesito
        Dim TMP44 As String
        TMP44 = CStr(TeclaNewFicha) + " " + CStr(TeclaNewFicha2)
        GK.Startlisen TMP44
    End If
    
    EstoyEnDisco = 0
    
    If HabilitarVUMetro Then
        If VU.DispositivosCant = 0 Then
            tERR.AppendLog "SinPLACA!!!"
            HabilitarVUMetro = False 'lo inhabilito!
            'YaCerrar3PM
            'Exit Sub
        Else
            VU.DefinePictureBox pVU1
            VU.DefinePictureBox2 pVU2
            VU.DefinePictureBox3 pVU3
            VU.DefinePictureBox4 pVU4
            
            imF = ExtraData.GetImagePath("vumetroprendido")
            VU.DefineImage 1, imF, True
            VU.DefineImage 3, imF, True
            imF = ExtraData.GetImagePath("vumetroapagado")
            VU.DefineImage 2, imF, True
            VU.DefineImage 4, imF, True
            
            pVU1.ZOrder
            pVU2.ZOrder
            pVU3.ZOrder
            pVU4.ZOrder
            VU.CantCuadros = 20
            VU.CantPic = 10
            VU.ColorBase = vbRed
        End If
    End If
    
    tERR.Anotar "cMM"
    Set MP3 = New tbrPlayer.MainPlayer
    
    tERR.Anotar "Ix001"
    Set TF = New tbrFOCUS.clsFOCUS
    TF.IntervalTimer = 5000
    TF.Iniciar Me.hwnd
    tERR.Anotar "Ix002"
    On Error GoTo MiErr
        
    If MostrarTouch Then
        'imagenes del touch screen
        imF = ExtraData.GetImagePath("touchizqnormal")
        cmdPagAt.Picture = LoadPicture(imF)
        imF = ExtraData.GetImagePath("touchderechanormal")
        cmdPagAd.Picture = LoadPicture(imF)
        imF = ExtraData.GetImagePath("botonokcomun")
        imgSELEC.Picture = LoadPicture(imF)
        
        imF = ExtraData.GetImagePath("botonsalirnormal")
        imgSALIR.Picture = LoadPicture(imF)
        
        imgSelec2.Picture = imgSELEC.Picture
        imF = ExtraData.GetImagePath("tocuharribacomun")
        cmdTouchArriba.Picture = LoadPicture(imF)
        imF = ExtraData.GetImagePath("touchabajocomun")
        cmdTouchAbajo.Picture = LoadPicture(imF)
        cmdTouchArriba2.Picture = cmdTouchArriba.Picture
        cmdTouchAbajo2.Picture = cmdTouchAbajo.Picture
    End If
    
    'ver si es superlicencia y usa otra tapa predeterminada
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("tddp322")) Then
            imF = GPF("tddp322")
        Else
            imF = ExtraData.GetImagePath("tapapredeterminada")
        End If
    Else
        imF = ExtraData.GetImagePath("tapapredeterminada")
    End If
    
    tbrPassImg1.Picture imF
    tERR.Anotar "acem", SYSfolder
    
    '*****************************
    '*****************************
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    AjustarFRM Me, 12000 'solo una vez despues sale todo a proporcion!
    BaseVista 'por unica vez cosas que no cambian
    UpdateVista 'acomodar todo segun variables SIEMPRE DESPUES DE AJUSTAR EL TAMAÑO DE LAS COSAS!
    '*****************************
    '*****************************
    
    'imagenes no cargadas, ver si hay algo configurado para el fondo
    imF = ExtraData.GetImagePath("FondoDeLasTapas")
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("iischu")) Then
            picFondoDisco.PaintPicture LoadPicture(GPF("iischu")), 0, 0, picFondoDisco.Width, picFondoDisco.Height
        Else
            picFondoDisco.PaintPicture LoadPicture(imF), 0, 0, picFondoDisco.Width, picFondoDisco.Height
        End If
    Else
        picFondoDisco.PaintPicture LoadPicture(imF), 0, 0, picFondoDisco.Width, picFondoDisco.Height
    End If

    tERR.Anotar "acek", imF
    
    RegistroDiario 'anota la fecha, hora y numero del contador
    
    tERR.Anotar "acet", HabilitarVUMetro
    
    'primero defino las separaciones y tamaños de los discos!
    
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
        AnchoTapaDisco = (picFondoDisco.Width / TapasMostradasH)
        AltoTapaDisco = (picFondoDisco.Height / TapasMostradasV)
    Else
        'el alto de estos incluye tambien el lbldisco
        AnchoTapaDisco = (picFondoDisco.Width * 0.8 / TapasMostradasH)
        AltoTapaDisco = (picFondoDisco.Height * 0.8 / TapasMostradasV)
        'ver cual es mayor para no permitir mucha distorsion
        'lo que se ajuste se agranda del espacio entrediscos
        EspacioEntreDiscosV = (picFondoDisco.Height * 0.2 / (TapasMostradasV + 1))
        EspacioEntreDiscosH = (picFondoDisco.Width * 0.2 / (TapasMostradasH + 1))
    End If
    
    'acomodo el disco cero con sus tamaños
    'NUEVO MAYO 07. Le quito los margenes segun corresponda!!!!
    'en este caso las variables que contienen los margenes solo guardan el porcentaje en enteros 0-100!
    Dim MargDer As Single, MargIzq As Single, MargSup As Single, MargInf As Single
    Dim IND As Long
    IND = ExtraData.GetIndexImage("marcodiscocomun")
    MargSup = ExtraData.GetFinalMargenSuperiorTra(IND) * AltoTapaDisco / 100
    MargInf = ExtraData.GetFinalMargenInferiorTra(IND) * AltoTapaDisco / 100
    MargDer = ExtraData.GetFinalMargenDerechoTra(IND) * AnchoTapaDisco / 100
    MargIzq = ExtraData.GetFinalMargenIzquierdoTra(IND) * AnchoTapaDisco / 100
    
    
    tERR.Anotar "acew", MostrarRotulos
    
    TapaCD(0).Width = AnchoTapaDisco - (MargDer + MargIzq)
    TapaCD(0).Height = AltoTapaDisco - (MargSup + MargInf + MargInf)
    
    If MostrarRotulos Then
        lblDISCO(0).Height = MargInf 'AltoTapaDisco * 0.19 '80%disco, 20% lbldisco
        lblDISCO(0).Width = AnchoTapaDisco - MargDer - MargIzq 'XXXX habria que ver si hay que sacar los margenes laterales
    Else
        lblDISCO(0).Visible = False
        lblDisco2(0).Visible = False
    End If
    'centrar!!
    Dim IniCentrarH As Long
    IniCentrarH = EspacioEntreDiscosH
    Dim IniCentrarV As Long
    IniCentrarV = EspacioEntreDiscosV
    lblDISCO(0).Left = IniCentrarH + MargDer
    TapaCD(0).Left = IniCentrarH + MargDer
    'ver si los rotulos van arriba o abajo
    tERR.Anotar "acex", RotulosArriba
    If RotulosArriba Then
        lblDISCO(0).Top = IniCentrarV
        TapaCD(0).Top = IniCentrarV + MargSup
    Else
        tERR.Anotar "000-0269"
        TapaCD(0).Top = IniCentrarV + MargSup
        lblDISCO(0).Top = IniCentrarV + AltoTapaDisco - (2 * MargInf) 'TapaCD(0).Top + TapaCD(0).Height - MargInf '+ 150
    End If
    
    tERR.Anotar "TCD(0).TOP", TapaCD(c).Top
    tERR.Anotar "LBL(0).TOP", lblDISCO(c).Top
    
    imF = ExtraData.GetImagePath("marcodiscocomun")
    
    imageFONDO(0).Picture = LoadPicture(imF)
    imageFONDO(0).Visible = True
    imageFONDO(0).Top = TapaCD(0).Top - 2 * MargSup 'IniCentrarV 'TapaCD(0).Top - 150
    imageFONDO(0).Left = IniCentrarH 'TapaCD(0).Left - 200
    imageFONDO(0).Width = AnchoTapaDisco 'TapaCD(0).Width + 400
    imageFONDO(0).Height = AltoTapaDisco + MargSup ' TapaCD(0).Height + lblDISCO(0).Height + 200
       
    TapaCD(0).ZOrder
    imageFONDO(0).ZOrder
    lblDisco2(0).ZOrder
    lblDISCO(0).ZOrder
 
    Dim CantDiscos As Long
    CantDiscos = TapasMostradasH * TapasMostradasV
    tERR.Anotar "acey", CantDiscos
    'cargar la cantidad de tapas correspondientes a una pagina!
    c = 0
    Do While c < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        tERR.Anotar "acez", c
        c = c + 1
        Load TapaCD(c)
        Load lblDISCO(c)
        Load lblDisco2(c)
        Load imageFONDO(c)
        'ya toman el tamaño del original
        
        Dim LineaTopActual As Long
        If c >= TapasMostradasH Then
            LineaTopActual = (AltoTapaDisco * (c / TapasMostradasH)) + (EspacioEntreDiscosV * ((c / TapasMostradasH) + 1))
                        'imageFONDO(c - TapasMostradasH).Top + _
                         imageFONDO(c - TapasMostradasH).Height _
                         EspacioEntreDiscosV
        Else
            LineaTopActual = EspacioEntreDiscosV
        End If
        tERR.Anotar "LTA(" + CStr(c) + ")", LineaTopActual
        If c / TapasMostradasH = c \ TapasMostradasH Then
            'es una tapa al principio de linea!!!!
            lblDISCO(c).Left = IniCentrarH + MargDer
            TapaCD(c).Left = TapaCD(0).Left
            TapaCD(c).Top = LineaTopActual + MargSup
            tERR.Anotar "TCD(" + CStr(c) + ").TOP", TapaCD(c).Top
            If RotulosArriba Then
                lblDISCO(c).Top = LineaTopActual
                tERR.Anotar "LBL(" + CStr(c) + ").TOP", lblDISCO(c).Top
                TapaCD(c).Visible = True
                imageFONDO(c).Visible = True
                If MostrarRotulos Then
'                   TapaCD(c).Top =lblDISCO(c).Top + lblDISCO(c).Height + 50
                    lblDISCO(c).Visible = True
                    lblDisco2(c).Visible = True
                Else
'                   TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + 50
                End If
            Else
'                If MostrarRotulos Then
'                    TapaCD(c).Top = lblDISCO(c - TapasMostradasH).Top + lblDISCO(c - TapasMostradasH).Height + EspacioEntreDiscosV
'                Else
'                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
'                End If
                lblDISCO(c).Top = LineaTopActual + AltoTapaDisco - (2 * MargInf) 'TapaCD(c).Top + TapaCD(c).Height - MargInf '+ 150
                tERR.Anotar "LBL(" + CStr(c) + ").TOP", lblDISCO(c).Top
                TapaCD(c).Visible = True
                imageFONDO(c).Visible = True
                If MostrarRotulos Then
                    lblDISCO(c).Visible = True
                    lblDisco2(c).Visible = True
                End If
            End If
        Else 'el c-1 tiene el mismo top, es cualquiera de una linea que no sea el pri de la izq
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                lblDISCO(c).Left = lblDISCO(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH + MargDer
                lblDISCO(c).Top = lblDISCO(c - 1).Top
                tERR.Anotar "LBL(" + CStr(c) + ").TOP", lblDISCO(c).Top
                TapaCD(c).Left = lblDISCO(c).Left
                TapaCD(c).Top = TapaCD(c - 1).Top
                tERR.Anotar "TCD(" + CStr(c) + ").TOP", TapaCD(c).Top
                TapaCD(c).Visible = True
                imageFONDO(c).Visible = True
            Else
                TapaCD(c).Left = TapaCD(c - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                TapaCD(c).Top = TapaCD(c - 1).Top
                
                tERR.Anotar "TCD(" + CStr(c) + ").TOP", TapaCD(c).Top
                lblDISCO(c).Left = TapaCD(c).Left
                lblDISCO(c).Top = lblDISCO(c - 1).Top
                tERR.Anotar "LBL(" + CStr(c) + ").TOP", lblDISCO(c).Top
                TapaCD(c).Visible = True
                imageFONDO(c).Visible = True
            End If
            If MostrarRotulos Then
                lblDISCO(c).Visible = True
                lblDisco2(c).Visible = True
            End If
        End If
        
        imF = ExtraData.GetImagePath("marcodiscocomun")
        imageFONDO(c).Picture = LoadPicture(imF)
        imageFONDO(c).Top = TapaCD(c).Top - 2 * MargSup 'TapaCD(c).Top - 150
        imageFONDO(c).Left = TapaCD(c).Left - MargIzq  ' TapaCD(c).Left - 200
        imageFONDO(c).Width = AnchoTapaDisco 'TapaCD(c).Width + MargDer + MargIzq 'TapaCD(c).Width + 400
        imageFONDO(c).Height = AltoTapaDisco + MargSup 'TapaCD(c).Height + MargSup + MargInf 'TapaCD(c).Height + lblDISCO(c).Height + 200
        
        TapaCD(c).ZOrder
        imageFONDO(c).ZOrder
        lblDisco2(c).ZOrder
        lblDISCO(c).ZOrder
    Loop
    'tERR.AppendLog "LISTO TAPAS"
    tERR.Anotar "acfa"
    SetKeyState vbKeyScrollLock, True
    lblV = "versión " + Trim(CStr(App.Major)) + "." + Trim(CStr(App.Minor)) + "." + Trim(CStr(App.Revision))
    'lblTiempoRestante = "Falta: " + "00:00"
    'ocultar las etiquetas
    tERR.Anotar "acfa2", lblV.Caption
    Me.AutoRedraw = True
    Me.Left = Screen.Width / 2 - Me.Width / 2
    Me.Top = Screen.Height / 2 - Me.Height / 2
    
    'mostrar creditos
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
    Origenes = LeerArch1Linea(GPF("origs"))
    PartOrigenes = Split(Origenes, "*")

    Dim AAA As Long
    For AAA = 65 To 90
        If AAA > 65 Then
            Load lLETRAS(AAA - 65)
            Load lLETRAS2(AAA - 65)
            lLETRAS(AAA - 65).Visible = True
            lLETRAS2(AAA - 65).Visible = True
        End If
        
        lLETRAS(AAA - 65).Caption = Chr(AAA)
        lLETRAS2(AAA - 65).Caption = lLETRAS(AAA - 65).Caption
        
        If AAA > 65 Then
            lLETRAS(AAA - 65).Left = lLETRAS(AAA - 66).Left + lLETRAS(AAA - 66).Width + 60
        End If
        
        lLETRAS2(AAA - 65).Left = lLETRAS(AAA - 65).Left + 15
        lLETRAS2(AAA - 65).Top = lLETRAS(AAA - 65).Top + 15
    Next AAA
    
    For AAA = 0 To UBound(PartOrigenes)
        If AAA > 0 Then
            Load lRITMO(AAA)
            Load lRITMO2(AAA)
            lRITMO(AAA).Left = lRITMO(AAA - 1).Left + lRITMO(AAA - 1).Width + 160
            lRITMO(AAA).Visible = True
            lRITMO2(AAA).Visible = True
        End If
        
        lRITMO(AAA).Caption = FSO.GetBaseName(PartOrigenes(AAA))
        lRITMO2(AAA).Caption = lRITMO(AAA).Caption
        lRITMO2(AAA).Left = lRITMO(AAA).Left + 15
        lRITMO2(AAA).Top = lRITMO(AAA).Top + 15
        
        tERR.Anotar "acfc", PartOrigenes(AAA)
        'ver los discos del origene elegido
        MtxTmpOrigenes() = ObtenerDir(PartOrigenes(AAA))
        'acumular a la matriz general
        SumarMatriz MATRIZ_DISCOS, MtxTmpOrigenes
    Next AAA
    '
    
    '=============================================================================
    '=============================================================================
    Dim MD As Long
    Randomize
    MD = CLng(Rnd * 49)
    
    tERR.Anotar "001-0063"
    If K.LICENCIA <= CGratuita And UBound(MATRIZ_DISCOS) > MD Then
        'limite de discos
        tERR.Anotar "001-0064"
        MsgBox "Esta es una version demo y no se pueden cargar muchos " + _
        " discos." + vbCrLf + _
        "Para conseguir la versión sin límite de discos y con el manual " + _
        "completo envie un e-mail a tbrsoft@hotmail.com o a " + _
        "tbrsoft@cpcipc.org."
        tERR.Anotar "001-0065"
        'cortar la matriz
        ReDim Preserve MATRIZ_DISCOS(MD)
    End If
    '=============================================================================
    '=============================================================================
    
    'puedo revisar cada disco para saber si tiene MM!!!
    Dim CantMM As Long, MMs() As String, dDI As String
    Dim IsQuitar As String 'lista de los indices a quitar
    IsQuitar = ""
    For AAA = 0 To UBound(MATRIZ_DISCOS)
        'obtengo la lista de arhivos
        dDI = txtInLista(MATRIZ_DISCOS(AAA), 0, ",")
        If dDI = "_RANK_" Then 'este ni siquiera existe en el disco
            CantMM = 10
        Else
            MMs = ObtenerArchMM(dDI)
            tERR.Anotar "caam", AAA, dDI
            'veo que tenga al menos 1!
            CantMM = UBound(MMs)
        End If
        If CantMM = 0 Then
            'MsgBox "El disco " + txtInLista(MATRIZ_DISCOS(AAA), 0, ",") + _
                " no tiene contenido multimedia!"
            tERR.Anotar "caak", AAA, dDI
            'si se quita aqui en un for que depende del ubound voy a generar errores!!!
            'QuitaIndiceMatriz MATRIZ_DISCOS, AAA
            IsQuitar = IsQuitar + CStr(AAA) + " "
        End If
    Next AAA
    
    If Len(IsQuitar) > 1 Then
        tERR.Anotar "caao", IsQuitar
        Dim ListaQuitar() As String
        ListaQuitar = Split(IsQuitar)
        For AAA = 0 To UBound(ListaQuitar) - 1 'el ultimo indice es "" por que solo hay un espacio vacio!
            QuitaIndiceMatriz MATRIZ_DISCOS, CLng(ListaQuitar(AAA))
        Next AAA
    End If
    'ya se sumop y esta listo para cargarse ordenados los discos dentro de cada origen
    tERR.Anotar "caan", AAA
    MostrarDiscosMTX
    'MATRIZ_DISCOS() = ObtenerDir(AP + "discos")
    
    BeginRoll
    
    Dim CarpActual As String
    Dim pathTema As String, DuracionTema As String, nombreTEMA As String
    'mostrar proceso
    ReDim Preserve MATRIZ_TOTAL(150, 30)
    
    'ret devuelve la cantidadd de discos cargados
    tERR.Anotar "acfd"
    DiscosEnPagina = CargarDiscos(0, True, 1)
    
    tERR.Anotar "acfe", DiscosEnPagina
    
    'lblTOTdiscos = "Discos: " + Trim(CStr(UBound(MATRIZ_DISCOS)))
    tERR.Anotar "acff", ReINI
    'si quedaron temas pendientes cargarlos
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            tLST.ListaAbrirDeDisco GPF("casc1001")
            EMPEZAR_SIGUIENTE 3
        Case "NADA"
            'no hacer nada
            'borrar la lista
            'borrra los temas 'y los creditos?
            If FSO.FileExists(GPF("casc1001")) Then FSO.DeleteFile GPF("casc1001"), True
            Timer1.Interval = 10000
    End Select
    
    Unload frmINI
    
    'ver si hay validacion por creditos
    Validar = LeerConfig("Validar", "0")
    tERR.Anotar "acfh", Validar
    If Validar Then
        'ver si existe el archivo Creditos Validar
        
        If FSO.FileExists(GPF("radliv")) Then
            'leer el archivo de creditos vaildados
            CreditosValidar = CLng(LeerArch1Linea(GPF("radliv")))
            tERR.Anotar "acfi", CreditosValidar
            'CodigoParaClaveActual busca el archivo con el numero que corresponde validar en este periodo de control
        Else
            tERR.Anotar "acfj"
            EscribirArch1Linea GPF("radliv"), "0"
            CreditosValidar = 0
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el máximo y si hay que avisar
        
        ValidarCada = LeerConfig("ValidarCada", "3000")
        AvisarAntes = LeerConfig("AvisarAntes", "500")
        tERR.Anotar "acfj", CreditosValidar, ValidarCada, AvisarAntes
        If (CreditosValidar > ValidarCada - AvisarAntes) Then
            'solicitar una clave
            'se podra saltear solo si todavia no llego al limite
            'uso el frmClave que tiene la variable publica ClaveIngresada
            Dim ClaveCorrespondiente As String
            ClaveCorrespondiente = NumToTec(ClaveParaValidar(CodigoParaClaveActual))
            
            tERR.Anotar "acfl"
            frmCLAVE.Show 1
            tERR.Anotar "acfm", UCase(ClaveIngresada), UCase(ClaveCorrespondiente)
            If TexToTec(UCase(ClaveIngresada)) <> UCase(ClaveCorrespondiente) Then
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
                EscribirArch1Linea GPF("radliv"), "0"
                'empezar un nuevo periodo
                CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            End If
        End If
        tERR.Anotar "acfo", ValidarCada, CodigoParaClaveActual
        lblValidar = "Val=" + CStr(ValidarCada) + "-Qued=" + CStr(ValidarCada - CreditosValidar) + "Actual=" + CStr(CreditosValidar) + " Codigo: " + CodigoParaClaveActual
    End If
    tERR.Anotar "acfj2", PUBs.HabilitarPublicidadesIMG, PUBs.SonarPublicidadesIMGCada
'    'caso especial Eduardo rodirguez
'    If ClaveAdmin = "ERO77701192FF" Or ClaveAdmin = "MARC777" Then
'        RollCRED.ReplaceIndex 2, "3PM"
'    End If

    'ver que onda con la publicidad de imagenes
    tbrPassImg1.ActivarPUBS = PUBs.HabilitarPublicidadesIMG
    tbrPassImg1.IntervalBetwenIMGs = PUBs.SonarPublicidadesIMGCada
    tbrPassImg1.ClearList
    'empiezan en 1 ambos!!!
    Dim AA As Long
    For AA = 1 To PUBs.TotalPUBsIMG
        tbrPassImg1.AddArchivoIMG (PUBs.ArchsPubsIMG(AA))
    Next
    
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    
    tbrPassImg1.IniciarPASS
    
    'si no tiene el foco ponerlo!!!
    If TF.GetState <> 1 Then TF.PonerFoco
    
    'lo prendo por mas que no haya protecto configurado por que lo uso para salir de los
    'discos tambien!
    frmIndex.Timer3.Interval = 3000
    
    VerSiTocaVMute
    
    Exit Sub
NoLoadIndex:
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdu"
    Resume Next
End Sub

Private Sub BeginRoll()
    '--------
    RollCRED.SetInterval 40
    RollCRED.SetVarColor 7
    RollCRED.MaxlargoRenglon = 30
    RollCRED.TextoACola "disfrute su música", vbYellow   '0 es
    RollCRED.TextoACola "desafios digitales", vbGreen '1 es lista de precios
    RollCRED.TextoACola "tbrSoft", vbRed              '2 es texto de SL al publico
    RollCRED.TextoACola "tbrSoft", vbBlue             '3 es texto gratis al publico
    If TOTAL_DISCOS < 2 Then
        RollCRED.TextoACola "NO HAY DISCOS. " + vbCrLf + _
                            "Presione 'C' para ingresar" + vbCrLf + _
                            "a la configuracion y utilize" + vbCrLf + _
                            "el asistente para cargar " + vbCrLf + _
                            "multimedia al sistema.", vbBlue
    End If
    
    If K.LICENCIA <= CGratuita Then
        RollCRED.TextoACola "PRESIONE F3" + vbCrLf + "PARA USAR EL SOFTWARE" + vbCrLf + "SIN RESTRICCIONES", vbGreen
    End If
    
    RollCRED.INI
    
    RollSONG.SetInterval 30
    RollSONG.SetVarColor 5
    RollSONG.MaxlargoRenglon = 30
    RollSONG.TextoACola "Sin reproducción", vbGreen 'cancion que se esta reproduciendo + rank
    RollSONG.TextoACola "no hay proximas canciones", vbBlue 'la proxima cancion
    RollSONG.TextoACola "no hay proximas canciones", vbRed 'algun elemento del ranking
    If TOTAL_DISCOS < 2 Then
        RollSONG.TextoACola "NO HAY DISCOS. " + vbCrLf + _
                            "Presione 'C' para ingresar" + vbCrLf + _
                            "a la configuracion y utilize" + vbCrLf + _
                            "el asistente para cargar " + vbCrLf + _
                            "multimedia al sistema.", vbBlue
    End If
    If K.LICENCIA <= CGratuita Then
        RollSONG.TextoACola "PRESIONE F3" + vbCrLf + "PARA USAR EL SOFTWARE" + vbCrLf + "SIN RESTRICCIONES", vbGreen
    End If
    RollSONG.INI
    
    tERR.Anotar "acep", K.LICENCIA
    If K.LICENCIA <= aSinCargar Then
        RollCRED.ReplaceIndex 3, "Este espacio sera suyo " + vbCrLf + _
                                 "cuando adquiera la " + vbCrLf + _
                                 "version full de 3PM"
    Else
        RollCRED.ReplaceIndex 3, textoUsuario
    End If
    
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    '-----------------
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("tslpri112")) Then
            tERR.Anotar "aceq"
            Set TE = FSO.OpenTextFile(GPF("tslpri112"), ForReading, False)
            Dim NewT As String
            NewT = TE.ReadAll
            RollCRED.ReplaceIndex 2, NewT
            TE.Close
        Else
            tERR.Anotar "acer"
            RollCRED.ReplaceIndex 2, "Software desarrollado" + vbCrLf + _
                                     "por tbrSoft " + vbCrLf + _
                                     "www.tbrsoft.com" + vbCrLf + _
                                     "info@tbrsoft.com" + vbCrLf + _
                                     "tbrsoft@cpcipc.org."
        End If
    Else
        tERR.Anotar "aces"
        RollCRED.ReplaceIndex 2, "Software desarrollado" + vbCrLf + _
                                     "por tbrSoft " + vbCrLf + _
                                     "www.tbrsoft.com" + vbCrLf + _
                                     "info@tbrsoft.com" + vbCrLf + _
                                     "tbrsoft@cpcipc.org."
    End If
    '-----------------
End Sub

Public Sub SelDisco(nDisco As Long)
    
    On Error GoTo MiErr
    
    'version 7 con fondo cheto
    imF = ExtraData.GetImagePath("marcodiscoelegido")
    imageFONDO(nDisco).Picture = LoadPicture(imF)
    'lblDisco(nDisco).ForeColor = vbWhite
    tERR.Anotar "acfp", nDisco, nDiscoSEL, nDiscoGral
    
    nDiscoSEL = nDisco
        
    Dim AAA As Long
    
    Dim FolRit As String
    Dim FolSel As String
    LineRitmo.Visible = False
    Dim LeftRitmoSel As Long
    For AAA = 0 To UBound(PartOrigenes)
        'ver que ritmo esta
        FolSel = UCase(FSO.GetBaseName(FSO.GetParentFolderName(txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ","))))
        FolRit = UCase(lRITMO(AAA).Caption)
        If FolSel = FolRit Then
            lRITMO(AAA).ForeColor = vbYellow
            LineRitmo.X1 = lRITMO(AAA).Left
            LineRitmo.X2 = lRITMO(AAA).Left + lRITMO(AAA).Width
            LineRitmo.Y1 = lRITMO(AAA).Top
            LineRitmo.Y2 = LineRitmo.Y1
            LineRitmo.Visible = True
            LeftRitmoSel = lRITMO(AAA).Left
        Else
            lRITMO(AAA).ForeColor = vbWhite
        End If
    Next AAA
    
    LineLETRA.Visible = False
    For AAA = 65 To 90
        
        If AAA > 65 Then
            lLETRAS(AAA - 65).Left = lLETRAS(AAA - 66).Left + lLETRAS(AAA - 66).Width + 60
        Else 'es el primero ponerlo debajo del ritmo
            lLETRAS(AAA - 65).Left = 60 'LeftRitmoSel
        End If
        
        lLETRAS2(AAA - 65).Left = lLETRAS(AAA - 65).Left + 15
        lLETRAS2(AAA - 65).Top = lLETRAS(AAA - 65).Top + 15
        
        If UCase(Left(lblDISCO(nDisco), 1)) = UCase(lLETRAS(AAA - 65).Caption) Then
            lLETRAS(AAA - 65).ForeColor = vbRed
            LineLETRA.X1 = lLETRAS(AAA - 65).Left
            LineLETRA.X2 = lLETRAS(AAA - 65).Left + lLETRAS(AAA - 65).Width
            LineLETRA.Y1 = lLETRAS(AAA - 65).Top
            LineLETRA.Y2 = LineLETRA.Y1
            LineLETRA.Visible = True
        Else
            lLETRAS(AAA - 65).ForeColor = vbYellow
        End If
    Next
    
    'seleccionar de la lista de solo video
    tERR.Anotar "acfq", nDisco, nDiscoSEL, nDiscoGral
    L(nDiscoGral).ForeColor = vbWhite
    L(nDiscoGral).BackColor = vbBlack
    LastDiscoSel = nDiscoGral 'para saber cual desactivar en unsel
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
    
    imF = ExtraData.GetImagePath("marcodiscocomun")
    imageFONDO(nDisco).Picture = LoadPicture(imF)
    'lblDisco(nDisco).ForeColor = vbBlack
    
    tERR.Anotar "acft", LastDiscoSel, EsVideo
    L(LastDiscoSel).ForeColor = vbBlack
    L(LastDiscoSel).BackColor = vbWhite
    If EsVideo Then OrdenarListaModoVideo
        
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdw"
    Resume Next

End Sub

Public Function CargarDiscos(numDiscoIniciar As Long, SelPrimero As Boolean, DeQueFila As Long, _
                             Optional ElegirDirecto As Long = -1) As Long
        
    On Local Error GoTo NoCRG
    
    'indicando en que disco se inicia carga ese y los seis (o lo que corresponde) _
        que le sigen
    'DeQueFila dice si es primero o último de cual fila!!!
    'devuelve el número de discos cargados
    Dim mCargarDiscos As Long
    mCargarDiscos = 0
    Dim TotPags As Long
    TotPags = (TOTAL_DISCOS - 1) \ (TapasMostradasH * TapasMostradasV)
    tERR.Anotar "acfu", numDiscoIniciar, SelPrimero, DeQueFila, TotPags
    
    'ACA SE DECIA EN QUE PAGINA SE ESTABA
    'lblPag = "Pagina " + CStr(Round(numDiscoIniciar / (TapasMostradasH * TapasMostradasV) + 1, 0)) + " de " + CStr(TotPags + 1)
    
    'tomar el disco que va a quedar seleccionado
    'como numero de disco en el indice general
    
    If ElegirDirecto < 0 Then
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

    End If
    
    tERR.Anotar "acfv", nDiscoGral, nDiscoSEL, TOTAL_DISCOS
    'esconder todos los discos
    Dim NDR As Long 'numero de tapa de disco real del 0 al 5 (total de discos-1)
    
    'no hacer esto al pedo si ya estan cargadas
    Dim NDI As Long '=numdiscoiniciar de la pagina
    Dim c As Integer
    c = 1
    
    NDI = numDiscoIniciar
    tERR.Anotar "acfw", NDI, SelPrimero
    
    'si no se cargaron al inicio!!
    tERR.Anotar "acgb", NDR, TapasMostradasH, TapasMostradasV
    Do While NDR < ((TapasMostradasH * TapasMostradasV))
        TapaCD(NDR).Visible = False
        lblDISCO(NDR).Visible = False
        lblDisco2(NDR).Visible = False
        imageFONDO(NDR).Visible = False
        NDR = NDR + 1
    Loop
    Dim ArchTapa As String
    NDR = 0
    tERR.Anotar "acgc", NDI, numDiscoIniciar, TapasMostradasH, TapasMostradasV
    
    Do While NDI < numDiscoIniciar + ((TapasMostradasH * TapasMostradasV))
        'ver si existe si hay disco con este n°
        'el '=' es de la 6.5
        If NDI <= UBound(MATRIZ_DISCOS) Then
            mCargarDiscos = mCargarDiscos + 1
            'ver si ya estan cargadas o se deben cargar
            tERR.Anotar "acgd", mCargarDiscos, NDI
            
            'ver si hay tapa
            ArchTapa = txtInLista(MATRIZ_DISCOS(NDI), 0, ",")
            
            If ArchTapa = "_RANK_" Then GoTo TAPADEF
            
            tERR.Anotar "acge", ArchTapa
            If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
            ArchTapa = ArchTapa + "tapa.jpg"
            If FSO.FileExists(ArchTapa) Then
                'si la tapa es demasiado grande
                If FileLen(ArchTapa) > TamanoTapaPermitido * 1024 Then
                    tERR.Anotar "acgf2", NDR, ArchTapa, CStr(FileLen(ArchTapa))
                    GoTo TAPADEF
                End If
                tERR.Anotar "acgf", NDR
                TapaCD(NDR).Picture = LoadPicture(ArchTapa)
            Else
TAPADEF:
                'ver si es superlicencia y usa otra tapa predeterminada
                'ver si es el rank o no
                Dim F6 As String
                If ArchTapa = "_RANK_" Then
                    F6 = "tddp323"
                    If K.LICENCIA = HSuperLicencia Then
                        If FSO.FileExists(GPF(F6)) Then
                            imF = GPF(F6)
                        Else
                            imF = ExtraData.GetImagePath("taparanking")
                        End If
                    Else
                        imF = ExtraData.GetImagePath("taparanking")
                    End If
                Else
                    F6 = "tddp322"
                    If K.LICENCIA = HSuperLicencia Then
                        If FSO.FileExists(GPF(F6)) Then
                            imF = GPF(F6)
                        Else
                            imF = ExtraData.GetImagePath("tapapredeterminada")
                        End If
                    Else
                        imF = ExtraData.GetImagePath("tapapredeterminada")
                    End If
                End If
                
                
                TapaCD(NDR).Picture = LoadPicture(imF)

            End If
            TapaCD(NDR).Visible = True
            imageFONDO(NDR).Visible = True

            'poner nombre al disco
            'antes en la 6.3 era NDI+1 !!
            lblDISCO(NDR) = txtInLista(MATRIZ_DISCOS(NDI), 1, ",")
            If MostrarRotulos Then
                lblDISCO(NDR).Visible = True
                lblDisco2(NDR).Visible = True
            End If
        End If
        tERR.Anotar "acgh", NDI, NDR
        NDI = NDI + 1
        NDR = NDR + 1
    Loop
    CargarDiscos = mCargarDiscos
    
    If ElegirDirecto >= 0 Then
        nDiscoGral = numDiscoIniciar + ElegirDirecto
        UnSelDisco nDiscoSEL 'deberia funcionar ?????
        SelDisco ElegirDirecto 'deberia funcionar ?????
    Else
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
    End If
    Exit Function
    
NoCRG:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdx"
    Resume Next

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    TF.Detener

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
End Sub

Private Sub GK_LlegoTecla(nTecla As Byte)
    'Dim VT As Long
    Select Case nTecla
        Case TeclaNewFicha
            ProcessKeyCoin TeclaNewFicha, 2
            'VT = CLng(Mid(List1.List(18), 5, Len(List1.List(18)) - 4))
            'List1.List(18) = "NF1:" + CStr(VT + 1)
        Case TeclaNewFicha2
            ProcessKeyCoin TeclaNewFicha2, 2
            'VT = CLng(Mid(List1.List(19), 5, Len(List1.List(19)) - 4))
            'List1.List(19) = "NF2:" + CStr(VT + 1)
    End Select
End Sub

Private Sub imageFONDO_Click(Index As Integer)

'*************************'*************************'*************************
    On Error GoTo MiErr
    'nunca hay que pasar hojas
    nDiscoGral = nDiscoGral + (Index - nDiscoSEL)
    tERR.Anotar "acgx", Index, nDiscoGral, TOTAL_DISCOS, nDiscoSEL
    'nDiscoGral = Index 'si se cargan todas las imágenes al inicio index=nDiscoGral
    If nDiscoGral + 1 > TOTAL_DISCOS Then
'        MsgBox "No existe el disco elegido!!. " + vbCrLf + _
'            "Carge discos desde el ADMINISTRADOR DE DISCOS en la " + vbCrLf + _
'            "página de configuracion (presionando la tecla 'C')"
        Exit Sub
    End If

    UnSelDisco nDiscoSEL
    Dim PagNum As Long
    PagNum = nDiscoGral \ (TapasMostradasH * TapasMostradasV)
    tERR.Anotar "acgy", PagNum
    nDiscoSEL = Index - (PagNum * (TapasMostradasH * TapasMostradasV))
    SelDisco nDiscoSEL
    'lblTOTdiscos = "Disco " + CStr(nDiscoGral + 1) + " de " + CStr(TOTAL_DISCOS)
    'tocar la tecla de entrar a disco
    SuperSel Index ', nDiscoGral
    Exit Sub

MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acea"
    Resume Next
'*************************'*************************'*************************

End Sub

Private Sub SuperSel(ByVal Index As Integer)
    
    'elegir el disco normalmente
    SelDisco CLng(Index)
    
    EstoyEnDisco = 2 'no estoy en ningun lado!
    Dim M As Long
    
    'ver cuales eran visibles!!!
    'para saber cuales hay que mostrar esto es exclusivamente para las ultimas páginas
    'no queda otra
    For M = 0 To (TapasMostradasH * TapasMostradasV) - 1
        If TapaCD(M).Visible Then
            TapaCD(M).Tag = "1" 'bandera de que hay que mostrar!
        Else
            TapaCD(M).Tag = "0" 'bandera de dejar escondida!!
        End If
        'lo uso solo en tapa cd, cuando esta no esta mostrada ni su label ni su fondo lo estan
        TapaCD(M).Visible = False
        imageFONDO(M).Visible = False
        lblDISCO(M).Visible = False
        lblDisco2(M).Visible = False
    Next M

    'imgDiscoSEL.Picture = imageFONDO(nDiscoGral).Picture
    imgDiscoSEL.Visible = False
    imgDiscoSEL.Stretch = True
    imgDiscoSEL.Picture = TapaCD(Index).Picture
    imgDiscoSEL.Width = (picFondoDisco.Width / 4)
    imgDiscoSEL.Height = (picFondoDisco.Height / 3)
    imgDiscoSEL.Top = picFondoDisco.Height / 2 - imgDiscoSEL.Height / 2
    imgDiscoSEL.Left = 500 'picFondoDisco.Width / 4 - imgDiscoSEL.Width / 2
    imgDiscoSEL.Visible = True
    
    lblDiscoSEL.Visible = False
    lblDiscoSEL2.Visible = False
    
    lblDiscoSEL.Caption = lblDISCO(Index).Caption
    lblDiscoSEL.Font.Size = lblDISCO(Index).Font.Size
    lblDiscoSEL.Top = imgDiscoSEL.Top + imgDiscoSEL.Height
    lblDiscoSEL.Left = imgDiscoSEL.Left
    lblDiscoSEL.Width = imgDiscoSEL.Width
    lblDiscoSEL.Height = 900
    
    lblDiscoSEL2.Caption = lblDiscoSEL.Caption
    lblDiscoSEL2.Font.Size = lblDiscoSEL.Font.Size
    lblDiscoSEL2.Top = lblDiscoSEL.Top + 15
    lblDiscoSEL2.Left = lblDiscoSEL.Left + 15
    lblDiscoSEL2.Width = lblDiscoSEL.Width
    lblDiscoSEL2.Height = lblDiscoSEL.Height
    
    lblDiscoSEL.Visible = True
    lblDiscoSEL2.Visible = True
    
    imgFondoDiscoSel.Visible = False
    imgFondoDiscoSel.Stretch = True
    imgFondoDiscoSel.Picture = imageFONDO(Index).Picture
    imgFondoDiscoSel.Width = (picFondoDisco.Width / 4) + 200
    imgFondoDiscoSel.Height = imgDiscoSEL.Height + lblDiscoSEL.Height '+ 400
    imgFondoDiscoSel.Top = imgDiscoSEL.Top - 100
    imgFondoDiscoSel.Left = imgDiscoSEL.Left - 200
    imgFondoDiscoSel.Visible = True
    
    imgDiscoSEL.ZOrder
    imgFondoDiscoSel.ZOrder
    lblDiscoSEL2.ZOrder
    lblDiscoSEL.ZOrder
    
    imgListaSong.Visible = False
    imgListaSong.Stretch = True
    
    imF = ExtraData.GetImagePath("MarcoFondodelosdiscos")
    imgListaSong.Picture = LoadPicture(imF)
    Dim IND As Long
    IND = ExtraData.GetIndexImage("MarcoChicoIndicadores")
    Dim MargDer As Long, MargIzq As Long, MargSup As Long, MargInf As Long
    MargSup = imgListaSong.Height * ExtraData.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = imgListaSong.Height * ExtraData.GetFinalMargenInferiorTra(IND) / 100
    MargDer = imgListaSong.Width * ExtraData.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = imgListaSong.Width * ExtraData.GetFinalMargenIzquierdoTra(IND) / 100
    
    
    imgListaSong.Top = 150
    If MostrarTouch Then
        imgListaSong.Height = (picFondoDisco.Height) - imgSELEC.Height - 300 - lblNOCREDIT.Height
    Else
        imgListaSong.Height = (picFondoDisco.Height) - 300
    End If
    
    imgListaSong.Left = imgFondoDiscoSel.Left + imgFondoDiscoSel.Width + 200 ' picFondoDisco.Width / 2
    imgListaSong.Width = picFondoDisco.Width - imgListaSong.Left - 300
    imgListaSong.Visible = True
    
    'ya esta agrandado
    lblDATA.Font.Size = lblDiscoSEL2.Font.Size
    lblDATA.Width = imgFondoDiscoSel.Width - 200
    lblDATA.Height = picFondoDisco.Height - (imgFondoDiscoSel.Top + imgFondoDiscoSel.Height)
    lblDATA.Left = imgFondoDiscoSel.Left + 100
    lblDATA.Top = imgFondoDiscoSel.Top + imgFondoDiscoSel.Height + 300
    
    lblDATA2.Font.Size = lblDATA.Font.Size
    lblDATA2.Width = lblDATA.Width
    lblDATA2.Height = lblDATA.Height
    lblDATA2.Left = lblDATA.Left + 15
    lblDATA2.Top = lblDATA.Top + 15
    
    lblDATA.ZOrder
    
    'ahora cargar las canciones*******************************************************
    '*********************************************************************************
    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
    ReDim Preserve MATRIZ_TEMAS(0)
    If UbicDiscoActual = "_RANK_" Then
        MATRIZ_TEMAS = ObtenerRankComoMM(50)
    Else
        MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
        'de la forma!
        'D:\musica\Cuartetazo\Alma Fuerte-En vivo Obras 2001\02 - Almafuerte.mp3#02 - Almafuerte.mp3
    End If
    tERR.Anotar "caah", UBound(MATRIZ_TEMAS)
    'usar esto y no una variable para saber de discos vacios
    If UBound(MATRIZ_TEMAS) = 0 Then
        lblCanciones(0).Caption = "NO HAY CANCIONES EN ESTE DISCO!"
        tERR.AppendLog "No hay temas en el disco: " + UbicDiscoActual + ".acpu"
        Exit Sub
    End If
    
    Dim DataTXT As String
    If UbicDiscoActual = "_RANK_" Then
        DataTXT = "Estos son los mas escuchados !"
    Else
        Dim ArchDaTa As String
        ArchDaTa = UbicDiscoActual + "data.txt"
        If FSO.FileExists(ArchDaTa) Then
            Dim a As TextStream
            Set a = FSO.OpenTextFile(ArchDaTa, ForReading, False)
            DataTXT = a.ReadAll
        Else
            DataTXT = "No hay datos adicionales de este disco"
        End If
    End If
    
    lblDATA.Caption = DataTXT:    lblDATA2.Caption = lblDATA.Caption
    lblDATA.Visible = True:       lblDATA2.Visible = True
    
    Dim c As Integer, nombreTemas As String
    Dim pathTema As String
    c = 1
    Dim AltoRenglon As Long
    AltoRenglon = lblCanciones(0).Height + 30
    tERR.Anotar "caai", AltoRenglon
    Dim EXT As String

    Do While c <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(c), 0, "#")
        nombreTemas = txtInLista(MATRIZ_TEMAS(c), 1, "#")
        EXT = LCase(txtInLista(pathTema, 1, "."))
        
        'quitar el molesto .mp3 o lo que fuera
        Select Case LCase(EXT)
            Case "mp3"
                EXT = "" 'se sobreentiende que todo es mp3" (mp3-Musica)"
'            Case "mp4"
'                EXT = " (mp4-Musica)"
            Case "wma"
                EXT = " (wma-Musica)"
            Case "mpeg", "mpg", "avi", "wmv"
                EXT = " (" + LCase(EXT) + "-Video)"
            Case "vob"
                EXT = " (DVD!)"
            Case "dat"
                EXT = " (VCD-Video)"
        End Select
        nombreTemas = FSO.GetBaseName(nombreTemas) + EXT
        Load lblCanciones(c)
        Load lblCanciones2(c)
        
        lblCanciones(c).Caption = nombreTemas
        tERR.Anotar "caaj", c, nombreTemas
        lblCanciones(c).Tag = pathTema
        lblCanciones(c).Top = MargSup + (c * AltoRenglon)
        lblCanciones2(c).Top = lblCanciones(c).Top + 15
        lblCanciones2(c).Left = lblCanciones(c).Left + 15
        'tiene autosize
        'ver que no se muestren mas canciones de las que entren
        
        If lblCanciones(c).Top > (imgListaSong.Top + imgListaSong.Height _
                - AltoRenglon * 2 - MargInf) Then
            
            Exit Do
        End If
        
        c = c + 1 'ver que el proximo entre
    Loop
    
    Dim TotalSong As Long
    TotalSong = c - 1
    'en adelante se usa como referencia el ubound asi que lo corto directamente asi!
    ReDim Preserve MATRIZ_TEMAS(TotalSong)
    
    If CargarDuracionTemas Then
        'ahora cargar las duaciones
        Dim NoCargoDuracion As Long
        NoCargoDuracion = 0
        c = 1
        Dim MP3tmp As New MP3Info
        Do While c <= UBound(MATRIZ_TEMAS)
            pathTema = lblCanciones(c).Tag
            'si es mp3 usar el rápido, si no usar el viejo
            'XXXX no se si podra leer la duracion del mp4 igual que el mp3
            If UCase(Right(pathTema, 3)) = "MP3" Then '''Or UCase(Right(pathTema, 3)) = "MP4" Then
                MP3tmp.FileName = pathTema
                DuracionTema = MP3tmp.DurationSTR
            Else
                'en caso de que sea video el clsMp3 no anda!!
                'mostrar duracion VIEJO FORMATO
                DuracionTema = frmIndex.MP3.QuickLargoDeTema(pathTema)
                If DuracionTema = "N/S" Then
                    NoCargoDuracion = NoCargoDuracion + 1
                    If NoCargoDuracion > 3 Then
                        'hay algun problema y no cargo mas
'                        lstTIME.Visible = False
'                        lstTEMAS.Left = 50
'                        lstTEMAS.Width = lblNoEjecuta.Left - 50
                    End If
                End If
            End If
            lblCanciones(c).Caption = lblCanciones(c).Caption + " (" + DuracionTema + ")"
            c = c + 1
        Loop
        Set MP3tmp = Nothing
    End If

    'revisar especificamente que no haya nada mas largo que lo que se puede
    c = 1
    Do While c <= UBound(MATRIZ_TEMAS)
    
        'si o si dejar un margen
        If lblCanciones(c).Width > (imgListaSong.Width * 0.9) Then
            Dim D As Long
            For D = 1 To 35 'con estas pasadas debe quedar ok
                'que nunca de error!!!!
                If Len(lblCanciones(c).Caption) > 10 Then
                    lblCanciones(c).Caption = _
                        Mid(lblCanciones(c).Caption, 1, Len(lblCanciones(c).Caption) - 10) + "..."
                Else
                    Exit For
                End If
                'ver si con eso alcanza
                If lblCanciones(c).Width < (imgListaSong.Width * 0.9) Then Exit For
            Next D
        End If
        
        c = c + 1
    Loop


    c = 1

    Do While c <= UBound(MATRIZ_TEMAS)
        lblCanciones(c).Left = imgListaSong.Left + (imgListaSong.Width / 2 - lblCanciones(c).Width / 2)
        lblCanciones2(c).Left = lblCanciones(c).Left + 15
        lblCanciones(c).Visible = True
        lblCanciones2(c).Visible = True
        lblCanciones2(c).ZOrder 'lo necesito paar poder hacerle click
        lblCanciones(c).ZOrder
        c = c + 1
    Loop
    
    lblNOCREDIT.Left = imgListaSong.Left + (imgListaSong.Width / 2 - lblNOCREDIT.Width / 2)
    
    If MostrarTouch Then
        cmdTouchAbajo.Top = 120 'imgListaSong.Top + cmdTouchArriba.Height + 120
        cmdTouchAbajo.Left = (imgListaSong.Left / 2 - cmdTouchAbajo.Width) ' - 120
        
        cmdTouchArriba.Top = 120 'imgListaSong.Top + 120
        cmdTouchArriba.Left = (imgListaSong.Left / 2) '+ 120
        
        cmdTouchArriba.Visible = True
        cmdTouchAbajo.Visible = True
        
        imgSELEC.Left = imgListaSong.Left + (imgListaSong.Width / 3 - imgSELEC.Width)
        imgSELEC.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        imgSALIR.Left = imgListaSong.Left + (imgListaSong.Width / 1.5)
        imgSALIR.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        lblNOCREDIT.Top = imgSELEC.Height + imgSELEC.Top + 60
        
        imgSELEC.Visible = True
        imgSALIR.Visible = True
    Else
        lblNOCREDIT.Top = imgListaSong.Height + imgListaSong.Top - lblNOCREDIT.Height - 120
    End If
    
    
    EstoyEnDisco = 1
    OkInState1 = 0
    selDiscoI 1
End Sub

Private Function selDiscoI(i As Integer) As Long
    'elegir un disco de la lista
    'como aqui pueden venir con el mouse (que aun no pone en cero el contador de los botones)
    SecSinTecla = 0
    
    lblNOCREDIT.Visible = False
    tERR.Anotar "sdi", i
    Dim TMPi As Long 'para saber siempre que se eligio originalmente
    TMPi = i
    'elegir disco en el index
    'si solo quiero el que sigue pongo -1 o -2 para el anterior
    '-3 solo para saber cual esta elegido por ejemplo para reproducirlo
    'devuelve -99 si no hay nada mas para elegir
    Dim c As Long
    Dim sSEL As Long
    sSEL = -1 'bandera de que nada esta elegido
    If i < 0 Then
        'necesito saber cual esta elegido
        For c = 1 To UBound(MATRIZ_TEMAS)
            If lblCanciones(c).BackStyle = 1 Then
                sSEL = c
                Exit For
            End If
        Next c
    Else 'ya sabe lo que quiere
        sSEL = i
    End If
    tERR.Anotar "sdi2", i, sSEL
    'el que sigue
    If i = -1 Then sSEL = sSEL + 1
    'el anterior
    If i = -2 Then sSEL = sSEL - 1
    'ver que no se pase
    tERR.Anotar "sdi3", i, sSEL, UBound(MATRIZ_TEMAS)
    'el limite para ambos casos estaba en 1 y funiocnaba ok
    'pero en disco de una sola cancion anda ok on el cero que parece que es el que va
    If sSEL < 1 Then sSEL = UBound(MATRIZ_TEMAS)
    If sSEL > UBound(MATRIZ_TEMAS) Then sSEL = 1
    tERR.Anotar "sdi4", i, sSEL, lblDiscoSEL
    i = sSEL
    
    'ver si el que voy a elegir se puede elegir
    Dim CO As Long
    CO = 0
    Do
        If lblCanciones(i).Tag = "" Then 'lo pongo asi cuando una cancion se elije
            If TMPi = -1 Then sSEL = sSEL + 1
            If TMPi = -2 Then sSEL = sSEL - 1
        Else
            i = sSEL
            Exit Do 'ya encontre!
        End If
        i = sSEL
        CO = CO + 1
        'si dio toda la vuelta me voy!
        If CO >= lblCanciones.UBound Then
            selDiscoI = -99
            Exit Function
        End If
    Loop
    tERR.Anotar "sdi5", i, sSEL, UBound(MATRIZ_TEMAS)
    For c = 1 To UBound(MATRIZ_TEMAS)
        lblCanciones(c).BackColor = vbBlack
        lblCanciones2(c).BackColor = lblCanciones(c).BackColor
        If c = i Then
            lblCanciones(c).BackStyle = 1
            lblCanciones2(c).BackStyle = 1
        Else
            lblCanciones(c).BackStyle = 0
            lblCanciones2(c).BackStyle = 0
        End If
    Next c
    
    selDiscoI = i
    
End Function

Private Sub UnSuperSel()
    tERR.Anotar "sdi6"
    EstoyEnDisco = 1
    lblNOCREDIT.Visible = False
    imgSELEC.Visible = False
    imgSALIR.Visible = False
    Dim M As Long
        
    For M = 0 To (TapasMostradasH * TapasMostradasV) - 1
        If TapaCD(M).Tag = "1" Then
            tERR.Anotar "sdi7", M
            TapaCD(M).Visible = True
            imageFONDO(M).Visible = True
            lblDISCO(M).Visible = True
            lblDisco2(M).Visible = True
        End If
    Next M
    
    'imgDiscoSEL.Picture = imageFONDO(nDiscoGral).Picture
    imgDiscoSEL.Visible = False
    lblDiscoSEL.Visible = False
    lblDiscoSEL2.Visible = False
    imgFondoDiscoSel.Visible = False
    imgListaSong.Visible = False
    lblDATA.Visible = False
    lblDATA2.Visible = False
    
    'descargar todos los objetos cargados
    On Local Error Resume Next
    For M = 1 To 30 'no debo permitir que se cargue mas de 30
        Unload lblCanciones(M)
        Unload lblCanciones2(M)
    Next M
    
    If MostrarTouch Then
        cmdTouchArriba.Visible = False
        cmdTouchAbajo.Visible = False
    End If
    
    EstoyEnDisco = 0

End Sub

Private Sub imgSELEC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonokelegido")
    imgSELEC.Picture = LoadPicture(imF)
End Sub

Private Sub imgSELEC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonokcomun")
    imgSELEC.Picture = LoadPicture(imF)
    EjecutarDeTouch
End Sub

Private Sub imgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonsalirapretado")
    imgSALIR.Picture = LoadPicture(imF)
End Sub

Private Sub imgSalir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonsalirnormal")
    imgSALIR.Picture = LoadPicture(imF)
    UnSuperSel
End Sub

Private Sub imgSELEC2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonokelegido")
    imgSelec2.Picture = LoadPicture(imF)
End Sub

Private Sub imgSELEC2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imF = ExtraData.GetImagePath("botonokcomun")
    imgSelec2.Picture = LoadPicture(imF)
    Form_KeyUp TeclaOK, 0
End Sub

Private Sub lblCanciones_Change(Index As Integer)
    lblCanciones2(Index).Caption = lblCanciones(Index).Caption
End Sub

Private Sub lblCanciones_Click(Index As Integer)
    If selDiscoI(Index) = -99 Then 'todo esta elegido!
        UnSuperSel
    End If
End Sub

Private Sub lblCreditos_Change()
    Label2.Caption = lblCreditos.Caption
    Label2.Top = lblCreditos.Top + 15
    Label2.Left = lblCreditos.Left + 15
    Label2.ZOrder 1
End Sub

Private Sub lblCreditos_DblClick()
    List1.Visible = Not (List1.Visible)
    List1.ZOrder
End Sub

Private Sub lblDisco_Change(Index As Integer)
    lblDisco2(Index).Caption = lblDISCO(Index).Caption
    lblDisco2(Index).Left = lblDISCO(Index).Left + 15
    lblDisco2(Index).Top = lblDISCO(Index).Top + 15
    lblDisco2(Index).Width = lblDISCO(Index).Width
    lblDisco2(Index).Height = lblDISCO(Index).Height
End Sub

Private Sub lLETRAS_Click(Index As Integer)
    If EstoyEnDisco = 0 Then
        Dim CC As Long
        For CC = 0 To UBound(PartOrigenes)
            If lRITMO(CC).ForeColor = vbYellow Then
                SelPagina lRITMO(CC).Caption, lLETRAS(Index).Caption
            End If
        Next CC
    End If
End Sub

Private Sub lRITMO_Click(Index As Integer)
    If EstoyEnDisco = 0 Then SelPagina lRITMO(Index).Caption
End Sub

Private Sub MP3_BeginPlay(iAlias As Long)
    
    'los video mudos no se tocan
    If iAlias = 3 Then Exit Sub
    'si es la primera cancion no se detecta en el empezar siguiente
    If EsVideo Then
        MP3.DoStop 3
    End If
    EnableFF = False
    EnableNextMusic = False
    On Error GoTo MiErr
    tERR.Anotar "acgq", MP3.FileName(iAlias)
    Dim Tapa As String
    Tapa = FSO.GetParentFolderName(MP3.FileName(iAlias)) + "\tapa.jpg"
    
    tERR.Anotar "acgr", TotalTema(iAlias)
'    If TotalTema(iAlias) > 0 And MP3.IsPlaying(iAlias) Then
'        lblTiempoRestante = "TOTAL: " + MP3.Falta(iAlias)
'    Else
'        lblTiempoRestante = "Falta: " + "00:00"
'    End If
    
    VolBajando = MP3.Volumen(iAlias)
    tERR.Anotar "acgs", VolBajando
    
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdy"
    Resume Next
    
End Sub

Private Sub MP3_EndPlay(iAlias As Long)
    IenPlenaCancion(iAlias) = 0
    List1.List(iAlias) = ".. PLAY" + CStr(iAlias) + ":END"
    
    On Error GoTo MiErr
    
    'si es un video mudo entonces sigo con el otro
    If iAlias = 3 Then
        'parece que no se cerrar bien!!
        'o que apareciera en playing y por lo tanto no empezar la que sigue
        MP3.DoClose 3
        
        VerSiTocaVMute
        Exit Sub
    End If
    
    'estaba en played
    '-------------------------------------
    List1.List(8) = "STOP:" + CStr(iAlias)
    tERR.Anotar "acgv6", LastRetEmpezarSig, iAlias, CORTAR_TEMA(iAlias)
    'MP3.DoStop iAlias 'este desencadena un EndPlay !!!!!!!!!!!
    MP3.DoClose iAlias
    
    If LastRetEmpezarSig <> 4 Then 'no sigue un video
    
        'si es video y lo que sigue no es video esconder el picvideo
        'que puede molestar a la publicidad de la salida de tv
    
        'ademas que no haya publicidad en video mudoooo!!!
        If PUBs.HabilitarPublicidadesVMute = False Then
            frmVIDEO.picVideo.Visible = False
        Else
            'si no sigue un video ver si esta reproduciendo
            'y ademas es visible el "3"
            VerSiTocaVMute
            'parece que como el picVideo del frmvideo
            'tiene las imagenes de otro video no agarra el nuevo!
            'VerSiTocaVMute
        End If
    End If
    '-------------------------------------
    
    tERR.Anotar "acgt", PasarHoja, vidFullScreen, HabilitarVUMetro
    
    'antes al finalizar se desacomodaba todo a lo normal total el tema que segui se
    'acomodaba, ahora que las canciones empiezan antes de que termine esto molesta
    'todo se puso en un procedimiento ByeTema que se llama en otro momento
    
    'CMP cambio a multipista
    'si el tipo uso ff a los 15 segundos se paso de largo el segundo 10 exacto
    'y por lo tanto no se lanzo en ese momento
    'que es cuando empieza una cancion
    If EMPEZAR_SIGUIENTE(5) <> 4 Then
        'sigue algo que no es video!
        VerSiTocaVMute
    End If
    
    'si no hay tema a continuacion y termino un video no se acomodaba
    'en empezarsigioente ya esvideo se puso en false!
    UpdateVista
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdz"
    Resume Next
End Sub

Private Sub ByeTema()
    UpdateVista
End Sub

Private Sub MP3_mmError(txtMasHist As String)
    tERR.AppendLog txtMasHist
End Sub

Private Sub ShowPaso(Inis33 As String, iAlias As Long, SP As Long)
    List1.List(iAlias) = Inis33 + " PLAY" + CStr(iAlias) + ":" + CStr(SP) + _
            ":" + CStr(TotalTema(iAlias)) + "(" + _
            CStr(MP3.HastaTema(iAlias)) + "):" + CStr(MP3.Volumen(iAlias)) + _
            " CUT:" + CStr(CORTAR_TEMA(iAlias)) + _
            " ToySal:" + CStr(YaEsoySaliendoGrat_Cortar(iAlias))
End Sub

Private Sub MP3_Played(SecondsPlayed As Long, iAlias As Long, MS As Long)

    'cualquier cosa se corrige despues!
    EnableFF = True:    EnableNextMusic = True
    
    tERR.Anotar "acgv0", MS, iAlias, ThisFade, SegFade
    
    ShowPaso "==", iAlias, MS
    
    List1.List(4) = "IAA:" + CStr(IAA)
    List1.List(5) = "IAANext:" + CStr(IAANext)
    
    If iAlias = 3 Then Exit Sub
    
    On Error GoTo MiErr
    
    Dim NV As Long 'para nuevos voluemnes si se tienen  que cambiar
    
    'los primeros X segundos van en FadeIn sea el momento que sea
    If SecondsPlayed <= ThisFade Then
        '**********************************************
        IenPlenaCancion(iAlias) = 1 'indica que esta empezando
        '**********************************************
        YaEsoySaliendoGrat_Cortar(iAlias) = False
        
        List1.List(6) = "ININ:" + CStr(iAlias)
        EnableFF = False:        EnableNextMusic = False
        
        ShowPaso "++", iAlias, MS
    
        tERR.Anotar "acgv2", CORTAR_TEMA(iAlias), VolumenIni, VolumenIni2
        
        If CORTAR_TEMA(iAlias) Then
            NV = CLng(VolumenIni2 * ((MS / 1000) * (1 / ThisFade)))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
            MP3.Volumen(iAlias) = NV
        Else
            NV = CLng(VolumenIni * ((MS / 1000) * (1 / ThisFade)))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
            MP3.Volumen(iAlias) = NV
        End If
        
        GoTo SIGUE55
    End If
    'solo una vez pasa a ser activo el que era IAANext
    tERR.Anotar "acgv3", IAA, IAANext
    
    'como el timer pasa mas de una vez por segundo solo hacer el cambio una vez
    Dim F As Long, F2 As Single
    F = (TotalTema(iAlias) - SecondsPlayed)
    'para bajar el volumen mas gradualmente
    F2 = (TotalTema(iAlias) - (MS / 1000))
    tERR.Anotar "acgv4B", IenPlenaCancion(iAlias)
    'cuerpo de la canción
    If (SecondsPlayed > ThisFade) And (F > ThisFade) Then
        'son demasiadas veces por segundo que haga esto
        'de todas formas no afecta demasiado el tiempo de procesador
        
        If iAlias <> IAA Then
            Dim TMP As Long: TMP = IAANext: IAANext = IAA: IAA = TMP
        End If
        '**********************************************
        IenPlenaCancion(iAlias) = 2 'en plena cancion
        '**********************************************
    
        'si corto una cancion y thisfade era 1 si lo paso a 5 ahora va a entrar de
        'nuevo en el que secondsplayed es menor todavia
        'entonces lo actualizo asi
        If (SecondsPlayed > SegFade) And (ThisFade <> SegFade) Then
            ThisFade = SegFade
        End If
    
        'dejo listo para que empieze otro solo despues de termino bien terminado el
        'anterior! Esto es cuando el nuevo paso un rato!
            
        'seria un crimen que haya otra cancion aca!!!!!!!!!!!
        'me aseguro que se corten las canciones que ya llegaron a su final
        Dim J As Long
        For J = 0 To 2
            If iAlias <> J And MP3.IsPlaying(J) Then
                'avisar que hago cagadas!!!
                tERR.Anotar "SPUCDL0", iAlias, J, ThisFade, SecondsPlayed
                tERR.Anotar "SPUCDL1", TotalTema(iAlias), TotalTema(J)
                tERR.Anotar "SPUCDL2", MP3.FaltaInSec(iAlias), MP3.FaltaInSec(J)
                tERR.Anotar "SPUCDL3", MP3.EsGratis(iAlias), MP3.EsGratis(J)
                tERR.Anotar "SPUCDL4", MP3.Volumen(iAlias), MP3.Volumen(J)
                tERR.Anotar "SPUCDL5", IenPlenaCancion(iAlias), IenPlenaCancion(J)
                tERR.AppendLog "SPUCDL" 'se paso nuna cancion de largo
                
                'analizar cual debo matar y cual se queda como IAA
                'ACA ESTOY SEGURO DE QUE J <> IALIAS
                'El que tenga mas segundos recorridos estimo que es el que se esta muriendo
                If MP3.PositionInSec(J) > MP3.PositionInSec(iAlias) Then
                    MP3.DoStop J 'mato uno ...
                    IAA = iAlias 'dejo como unico al otro
                Else
                    MP3.DoStop iAlias
                    IAA = J
                End If
                'acomodar iaanext
                If IAA = 0 Then
                    IAANext = 1
                Else
                    IAANext = 0
                End If
            End If
            
        Next J
        'los videos mudos no tiene que ver con esto
        'If iAlias <> 3 And MP3.IsPlaying(3) Then MP3.DoStop 3
        '**********************************************
        LastRetEmpezarSig = -99 'solo como bandera para que termine solo _
            una vez esta cancion
    End If
    
    'este es el modo automático de finalizacion de las canciones
    'al llegar a los ultimos X segundos se va bajando hasta terminar
    
    tERR.Anotar "acgv4", F, ThisFade
    
    If F <= ThisFade Then
        tERR.Anotar "acgv4A", LastRetEmpezarSig, IenPlenaCancion(iAlias)
        ShowPaso "--", iAlias, MS
        '**********************************************
        IenPlenaCancion(iAlias) = 3 'terminando cancion
        '**********************************************
        EnableFF = False:        EnableNextMusic = False
        'ir abriendo el que sigue!!!
        'PUEDE ENTRAR MAS DE UNA VEZ ACA YA QUE EL TIMER ES MAS DE UNA VEZ POR SEGUNDO
        If LastRetEmpezarSig = -99 Then
        
            'en algunas putas pcs antes de terminar el empezar siguiente ya llega de nuevo aca!!!
            'adelanto la bandera a otro valor!!!
            LastRetEmpezarSig = -98
            
            List1.List(6) = "OPNEXT:" + CStr(iAlias)
            'aqui lo que estaba en el end play para desacomodarlo!
            
            ByeTema
            
            Dim lRet As Long
            lRet = EMPEZAR_SIGUIENTE(1)
            tERR.Anotar "acgv5", lRet, Salida2, PUBs.HabilitarPublicidadesVMute
            'este lastret... me sirve como bandera para que solo entre una vez
            LastRetEmpezarSig = lRet
            
            If lRet = 4 Then
                'sigue un video!!!!
                If Salida2 Then 'si sale en la tv corto la publicidad si hubiera
                    If PUBs.HabilitarPublicidadesVMute Then
                        MP3.DoStop 3
                    End If
                End If
            End If
            
            GoTo SIGUE55
        End If
        'ver si el tema se acorto para pasar al siguiente con la "B"!!!!!!
        
        If CORTAR_TEMA(iAlias) Then
            NV = CLng(VolumenIni2 * (F2 * (1 / ThisFade)))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
            MP3.Volumen(iAlias) = NV
        Else
            List1.List(7) = "CUTIN:" + CStr(iAlias)
            NV = CLng(VolumenIni * (F2 * (1 / ThisFade)))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
            MP3.Volumen(iAlias) = NV
        End If
        GoTo SIGUE55
    End If
    
SIGUE55:
    Dim PorcejEcutado As Long
    'esto pasa cada un segundo (si o si una vez por segundo)
    PorcejEcutado = CLng(SecondsPlayed / TotalTema(iAlias) * 100)
    
    frmIndex.List1.List(11) = "PorcEjec=" + CStr(PorcejEcutado)
    frmIndex.List1.List(12) = "PorcTema=" + CStr(PorcentajeTEMA)
    
    tERR.Anotar "acgw", MS, PorcejEcutado, PorcentajeTEMA, TotalTema(iAlias)
    
    'lblTiempoRestante = "Falta: " + MP3.Falta(iAlias)
    
    'temas de autoplay
    If CORTAR_TEMA(iAlias) Then
        If PorcejEcutado > PorcentajeTEMA Then
            If YaEsoySaliendoGrat_Cortar(iAlias) = False Then
                YaEsoySaliendoGrat_Cortar(iAlias) = True
                EMPEZAR_SIGUIENTE 2
            End If
        End If
    Else
        '===== sin licecnia ==================
        If K.LICENCIA <= CGratuita Then
            If SecondsPlayed > 46 Then
                If YaEsoySaliendoGrat_Cortar(iAlias) = False Then
                    YaEsoySaliendoGrat_Cortar(iAlias) = True
                    EMPEZAR_SIGUIENTE 2
                End If
            End If
        End If
        '=====================================
    End If
    
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdz"
    Resume Next
End Sub

Private Sub tbrPassImg1_ChangeImg()
    On Error GoTo MiErr
    'si se esta pasando un video no dar bola!!!
    tERR.Anotar "acgz", MP3.isPlayingAny, EsVideo
    If MP3.isPlayingAny And EsVideo And Salida2 Then
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

Private Sub TF_PerdioFoco(hwndFoco As Long)
    TF.PonerFoco
End Sub

Private Sub Timer1_Timer()
    On Error GoTo MiErr
    
    'controla el tiempo sin uso (sin ejecucion de temas)
    If MP3.IsPlaying(0) Or MP3.IsPlaying(1) Then Exit Sub
    'controla el tiempo sin uso (sin ejecucion de temas)
    SecSinUso = SecSinUso + (Timer1.Interval / 1000)

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
        If FSO.FileExists(GPF("rd3_444")) = False Then
            FSO.CreateTextFile GPF("rd3_444"), True
            'me voy al azar ya que no hay para elegirdel rank
            tERR.Anotar "achc.NORANK"
            GoTo MataReloj
        End If
        Set TE = FSO.OpenTextFile(GPF("rd3_444"), ForReading, False)
        Dim TT As String
        'antes de entra ver si el archivo no tiene nada
        If TE.AtEndOfStream Then
            tERR.Anotar "achd.NORANK"
            TE.Close
            GoTo MataReloj 'SE VAAAAA
        End If
        
        Dim CuentaVueltasBuscandoAzar As Long
        CuentaVueltasBuscandoAzar = 0
    
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
                    CORTAR_TEMA(IAANext) = True 'este tema se eligio al azar no va entero
                    SecSinUso = 0
                    TE.Close
                    EjecutarTema TemaAzar, False
                    Exit Sub
                Else
                    'NI BOSTA, NO PASA MUSICA QUE HAGA EL RANKING !!!
                    'NO HAY RANK!
                    tERR.Anotar "achm9", "NoRankAzar"
                End If
                Exit Do
            End If
         Loop
        tERR.Anotar "achm.REAZAR"

        On Local Error Resume Next
        TE.Close
    End If 'SecSinUso >= EsperaMinutos
Exit Sub

MataReloj:
    'mato este reloj
    Timer1.Interval = 0
    Exit Sub

MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acec"
    Resume Next
    
End Sub

Private Sub Timer3_Timer()
    On Error GoTo MiErr
    If Protector = 0 Then Exit Sub 'SE QUEDA PARA SALIR DE LOS DISCOS'Timer3.Interval = 0
    'para el reloj del protector. Lo ha inhabilitado
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + 3
    'dragones: destino de fuego
    'no protector en video
    If EsVideo Then SecSinTecla = 0
    tERR.Anotar "achn", SecSinTecla, EsperaTecla
    ' a los 7 segundos sale del disco!
    If SecSinTecla > 7 And EsVideo = False Then UnSuperSel
    If SecSinTecla > EsperaTecla And EsVideo = False Then frmProtect.Show 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aced"
    Resume Next

End Sub

Public Function TemasEnRank(MasDeXVotos) As Long
    On Error GoTo MiErr
    'indica cuantos temas hay en el ranking
    tERR.Anotar "acho", MasDeXVotos
    If FSO.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "achp"
        FSO.CreateTextFile GPF("rd3_444"), True
        TemasEnRank = 0
        Exit Function
    End If
    Set TE = FSO.OpenTextFile(GPF("rd3_444"), ForReading, False)
    
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

Public Sub BaseVista()
    'cosas que estaban en el update vista pero nunca cambian!
    'esto nunca cambia
    lblModoVideo.Left = Screen.Width - frModoVideo.Width
    lblTEMAS.Left = Screen.Width - frModoVideo.Width
    frModoVideo.Left = Screen.Width - frModoVideo.Width
    frTEMAS.Left = Screen.Width - frTEMAS.Width
    
    'barra inferior
    picFondo.Top = Screen.Height - picFondo.Height
    
    'barra superior
    picFondo2.Top = 0
    
    'contenedor de los discos
    frDiscos.Top = picFondo2.Height
    frDiscos.Height = picFondo.Top - picFondo2.Height
    
End Sub

Public Sub UpdateVista()
    frDiscos.Visible = False
    'hay una serie de variables que informan sobre como se debe ver
    
    'esVideo: se esta reproduciendo un video
    'salida2: cuando haya un video se vera en frmVideo (TV)
    'vidFullScreen: cuando se reproduzca un video sera a pantalla completa
    'NoVumVID no usa el vumetro en los videos
    'MostrarTouch: mostrar los botones de touch screen
    'EstoyEnModoVideoMiniSelDisco estoy eligiendo discos en modo texto
    'ModoVideoSelTema presione enter en el modo de texto viendo un video _
        y debe desplegarse la lista de canciones en modo texto
    'novumvid: no mostrar vumetros para videos en el monitor
    'habilitarVumetro: dejar espacio para el vumetro

    '**********************************************************
    'tomo como referencia el frDiscos que me define casi tod0
    
    If HabilitarVUMetro Then
        frDiscos.Width = Me.Width - (AnchoBarra * 2) - 50
        frDiscos.Left = AnchoBarra + 30
    Else
        frDiscos.Left = 0
        frDiscos.Width = Me.Width - 50
    End If
    picFondo2.Left = frDiscos.Left
    picFondo2.Width = frDiscos.Width 'screen.Width
    
    'ver si le saco el pedazo de la lista
    If EstoyEnModoVideoMiniSelDisco Then frDiscos.Width = frDiscos.Width - frModoVideo.Width
        
    
    imF = ExtraData.GetImagePath("MarcoFondodelosdiscos")
    frDiscos.PaintPicture LoadPicture(imF), 0, 0, frDiscos.Width, frDiscos.Height
    
    'listo frDiscos, ahora si lo tomo como referencia
    '**********************************************************
        
    frModoVideo.Visible = EstoyEnModoVideoMiniSelDisco
    lblModoVideo.Visible = EstoyEnModoVideoMiniSelDisco
    
    'si entre a la lista de discos hago lugar para eso
    If EstoyEnModoVideoMiniSelDisco Then AcomodarModoTexto 1
    
    tERR.Anotar "aceu", MostrarTouch, EstoyEnModoVideoMiniSelDisco
    
    If MostrarTouch Then
        cmdPagAd.Left = Screen.Width - cmdPagAd.Width
        cmdPagAt.Left = 0
        'cuando pasa al modo de texto de los discos queda muy feo si no se ocupa todo el ancho
        'ya que se borran los botones de touch laterales
        If EstoyEnModoVideoMiniSelDisco Then
            picFondo.Width = Me.Width
            picFondo.Left = 0
        Else
            picFondo.Width = Screen.Width - (cmdPagAt.Width * 2)
            picFondo.Left = cmdPagAt.Width
        End If
                 
        cmdPagAt.Top = picFondo.Top + ((picFondo.Height / 2) - (cmdPagAt.Height / 2))
        cmdPagAd.Top = picFondo.Top + ((picFondo.Height / 2) - (cmdPagAd.Height / 2))
    Else
        picFondo.Width = Me.Width
        picFondo.Left = 0
    End If
    'saco los touch de los costados para evitar confusiones en modo de texto
    If EstoyEnModoVideoMiniSelDisco Then
        cmdPagAt.Visible = False
        cmdPagAd.Visible = False
    Else
        cmdPagAt.Visible = MostrarTouch
        cmdPagAd.Visible = MostrarTouch
    End If
    
    imF = ExtraData.GetImagePath("MarcoChicoIndicadores")
    tERR.Anotar "aceu2", imF
    'picFondo.Picture = LoadPicture(imF)
    picFondo.PaintPicture LoadPicture(imF), 0, 0, picFondo.Width, picFondo.Height

    'dentro del picfondo hay que reacomodar
    lblCreditos.Left = picFondo.Width / 2 - lblCreditos.Width / 2
    tbrPassImg1.Left = picFondo.Width / 2 - tbrPassImg1.Width / 2
    lblCreditos.Top = 30
    'que se reacomode la sombra de los creditos, es necesario cuando se cambia el tamaño de picFondo!!!
    lblCreditos_Change
    
    'usar los datos del skin para saber como colocarlos
    Dim MargDer As Long, MargIzq As Long, MargSup As Long, MargInf As Long
    Dim IND As Long
    IND = ExtraData.GetIndexImage("MarcoChicoIndicadores")
    tERR.Anotar "aceu3", IND
    MargSup = picFondo.Height * ExtraData.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = picFondo.Height * ExtraData.GetFinalMargenInferiorTra(IND) / 100
    MargDer = picFondo.Width * ExtraData.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = picFondo.Width * ExtraData.GetFinalMargenIzquierdoTra(IND) / 100
    
    If MargSup = 0 Then MargSup = 60
    If MargInf = 0 Then MargInf = 60
    
    RollCRED.Width = tbrPassImg1.Left - MargDer  'tbrPassImg1.Left * 0.75
    RollSONG.Width = (picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) - MargIzq '(picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) * 0.75
    RollSONG.Height = picFondo.Height - MargSup - MargInf   'picFondo.Height * 0.7
    RollCRED.Height = picFondo.Height - MargSup - MargInf   ' picFondo.Height * 0.7
    RollCRED.Left = MargDer '(tbrPassImg1.Left / 2) - (RollCRED.Width / 2)
    RollSONG.Left = (tbrPassImg1.Left + tbrPassImg1.Width) '(tbrPassImg1.Left + tbrPassImg1.Width) + _
                    ((picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) / 2) - _
                    (RollSONG.Width / 2)
    RollSONG.Top = MargSup  'picFondo.Height / 2 - RollSONG.Height / 2
    RollCRED.Top = MargSup  'picFondo.Height / 2 - RollCRED.Height / 2
    
    'al momento de definir el skin se define un porcentaje que ocuapa cada uno de los 4 margenes
    
    
    IND = ExtraData.GetIndexImage("MarcoFondoDeLosDiscos")
    tERR.Anotar "aceu4", imF
    MargSup = frDiscos.Height * ExtraData.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = frDiscos.Height * ExtraData.GetFinalMargenInferiorTra(IND) / 100
    MargDer = frDiscos.Width * ExtraData.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = frDiscos.Width * ExtraData.GetFinalMargenIzquierdoTra(IND) / 100
    
    picFondoDisco.Top = MargSup
    picFondoDisco.Left = MargDer
    picFondoDisco.Height = frDiscos.Height - MargSup - MargInf
    picFondoDisco.Width = frDiscos.Width - MargDer - MargIzq
    
    tERR.Anotar "aceu5", Salida2
    
    picFondo.Visible = True 'imagen de fondo de los indicadores en modo simple
    'dejarlo porque se esconde en modo de video con vumetro en monitor
    If Salida2 = False Then
        If vidFullScreen Then
            'en este caso picvideo no sigue a frDiscos
            picVideo(0).Top = 0
            picVideo(0).Height = Me.Height
            If HabilitarVUMetro Then
                If NoVumVID Then
                    picVideo(0).Left = 0
                    picVideo(0).Width = Me.Width
                Else
                    picVideo(0).Left = AnchoBarra + 30
                    picVideo(0).Width = Me.Width - (AnchoBarra * 2) - 50
                    'QUE NEGRADA!!!! pero funciona
                    If EsVideo Then 'se esta ejecutando algo y se podrian ver de fondo
                        picFondo.Visible = False 'imagen de fondo de los indicadores en modo simple
                        cmdPagAt.Visible = False
                        cmdPagAd.Visible = False
                    End If
                End If
            Else
                picVideo(0).Left = 0
                picVideo(0).Width = Me.Width
            End If
                
        Else 'el video sale en minimo y la listas de textos tiene un lugar
            picVideo(0).Top = frDiscos.Top + picFondoDisco.Top
            picVideo(0).Left = frDiscos.Left + picFondoDisco.Left
            picVideo(0).Height = picFondoDisco.Height  'frDiscos.Height 'este ya sabe si es exclusivo y su alto correspondiente !
            picVideo(0).Width = picFondoDisco.Width ' frDiscos.Width
        End If
    End If
    
    'si no tuviera que iniciar sale solo de alli
    StartVu "custom"        '"grande"
    
    'se hacen visibles o invisibles desde donde corresponda
    picVideo(1).Top = picVideo(0).Top
    picVideo(1).Left = picVideo(0).Left
    picVideo(1).Width = picVideo(0).Width
    picVideo(1).Height = picVideo(0).Height
    tERR.Anotar "aceu6", HabilitarVUMetro
    
    If EsVideo = False Then
        picVideo(0).Visible = False
        picVideo(1).Visible = False
    End If
    
    frDiscos.Visible = True
End Sub

Private Sub txtS3_Change()
    If txtS3 = "" Then Exit Sub
    
    Dim P As String
    P = txtS3
    
    Dim SP() As String
    SP = Split(P, ":")
    
    If SP(0) = "sD" Then
        Select Case CLng(SP(1))
            Case TeclaIZQx2: SendKeys Chr(TeclaIZQ)
            Case TeclaDERx2: SendKeys Chr(TeclaDER)
            Case TeclaPagAdx2: SendKeys Chr(TeclaPagAd)
            Case TeclaPagAtx2: SendKeys Chr(TeclaPagAt)
            Case TeclaOKx2: SendKeys Chr(TeclaOK)
            Case TeclaESCx2: SendKeys Chr(TeclaESC)
            Case TeclaConfigx2: SendKeys Chr(TeclaConfig)
            Case TeclaCerrarSistemax2: SendKeys Chr(TeclaCerrarSistema)
            Case TeclaShowContadorx2: SendKeys Chr(TeclaShowContador)
            Case TeclaPutCeroContadorx2: SendKeys Chr(TeclaPutCeroContador)
            Case TeclaFFx2: SendKeys Chr(TeclaFF)
            Case TeclaBajaVolumenx2: SendKeys Chr(TeclaBajaVolumen)
            Case TeclaSubeVolumenx2: SendKeys Chr(TeclaSubeVolumen)
            Case TeclaNextMusicx2: SendKeys Chr(TeclaNextMusic)
            Case TeclaNewFichax2: Form_KeyUp TeclaNewFicha, 0  'especial directo
            Case TeclaNewFicha2x2: Form_KeyUp TeclaNewFicha2, 0
        End Select
    End If
    
    'vaciarlos !!!
    txtS3 = ""
    
End Sub

