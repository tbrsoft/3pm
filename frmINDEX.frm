VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
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
      Index           =   2
      Left            =   9000
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   59
      Top             =   3900
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox picFondo 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1620
      Left            =   120
      ScaleHeight     =   1620
      ScaleWidth      =   6600
      TabIndex        =   3
      Top             =   7350
      Width           =   6600
      Begin tbr3pm.txtRolling RollCRED 
         Height          =   1335
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   2355
      End
      Begin tbr3pm.txtRolling RollSONG 
         Height          =   1155
         Left            =   4740
         TabIndex        =   15
         Top             =   360
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   2037
      End
      Begin tbr3pm.tbrPassImg tbrPassImg1 
         Height          =   1260
         Left            =   2580
         TabIndex        =   7
         Top             =   330
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
            TabIndex        =   8
            Top             =   180
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
         Left            =   2580
         TabIndex        =   5
         Top             =   90
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
         TabIndex        =   6
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
         TabIndex        =   4
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
         Left            =   3270
         TabIndex        =   16
         Top             =   30
         Width           =   2235
      End
   End
   Begin VB.PictureBox picFondoPacha 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   150
      ScaleHeight     =   555
      ScaleWidth      =   5025
      TabIndex        =   56
      Top             =   6720
      Visible         =   0   'False
      Width           =   5025
      Begin tbrFaroButton.fBoton btOKPacha 
         Height          =   600
         Left            =   1020
         TabIndex        =   57
         Top             =   0
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   1058
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Escuchar cancion"
         fEnabled        =   -1  'True
         fFontN          =   "Verdana"
         fFontS          =   10
         fECol           =   5452834
      End
      Begin VB.Label lCredPacha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Credito $ 15000,00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   405
         TabIndex        =   58
         Top             =   60
         Width           =   2835
      End
      Begin VB.Image t3 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   4110
         Top             =   60
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Image t1 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   60
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
   End
   Begin VB.TextBox tUltra 
      Height          =   285
      Left            =   10950
      TabIndex        =   55
      Top             =   2670
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox tBT 
      Height          =   285
      Left            =   10530
      TabIndex        =   40
      Top             =   2670
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox TUsb 
      Height          =   285
      Left            =   11550
      TabIndex        =   39
      Top             =   2670
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.PictureBox Fondoxxx 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   90
      Picture         =   "frmINDEX.frx":08CA
      ScaleHeight     =   1005
      ScaleWidth      =   1455
      TabIndex        =   38
      Top             =   4350
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox picKAR 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1905
      Left            =   10140
      ScaleHeight     =   1905
      ScaleWidth      =   1815
      TabIndex        =   28
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
      Begin VB.Label lblWAIT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "WAIT"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   690
         Left            =   270
         TabIndex        =   33
         Top             =   0
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.Shape shKAR 
         BackColor       =   &H0080FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H0000C0C0&
         Height          =   255
         Left            =   720
         Shape           =   3  'Circle
         Top             =   1350
         Width           =   255
      End
      Begin VB.Label LF1 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   600
         Left            =   60
         TabIndex        =   30
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label lblTimeK 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   660
         Left            =   30
         TabIndex        =   29
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblTimeK2 
         BackStyle       =   0  'Transparent
         Caption         =   "00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   540
         Left            =   0
         TabIndex        =   32
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label LF2 
         BackStyle       =   0  'Transparent
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   600
         Left            =   60
         TabIndex        =   31
         Top             =   1380
         Width           =   825
      End
   End
   Begin VB.PictureBox pVU2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   6990
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   27
      Top             =   330
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox pVU4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   6480
      ScaleHeight     =   825
      ScaleWidth      =   480
      TabIndex        =   26
      Top             =   90
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pVU3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   7650
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   25
      Top             =   300
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picFondo2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   765
      Left            =   9150
      ScaleHeight     =   765
      ScaleWidth      =   1890
      TabIndex        =   18
      Top             =   150
      Width           =   1890
      Begin VB.Line LineRitmo 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   5
         X1              =   1350
         X2              =   1920
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line LineLETRA 
         BorderColor     =   &H000000FF&
         BorderWidth     =   3
         X1              =   1080
         X2              =   1650
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label lLETRAS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   585
         Index           =   0
         Left            =   30
         TabIndex        =   22
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   330
      End
      Begin VB.Label lRITMO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lRitmo"
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
         Left            =   60
         TabIndex        =   21
         Top             =   -30
         UseMnemonic     =   0   'False
         Width           =   600
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   2940
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lRITMO2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lRitmo"
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
         Left            =   180
         TabIndex        =   23
         Top             =   -30
         UseMnemonic     =   0   'False
         Width           =   600
      End
      Begin VB.Label lLETRAS2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   585
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   210
         UseMnemonic     =   0   'False
         Width           =   330
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   10000
      Left            =   11370
      Top             =   3270
   End
   Begin VB.Timer Timer1 
      Left            =   11370
      Top             =   3720
   End
   Begin VB.PictureBox frDiscos 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   60
      ScaleHeight     =   4185
      ScaleWidth      =   6225
      TabIndex        =   17
      Top             =   60
      Width           =   6225
      Begin VB.PictureBox picFondoDisco 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   4005
         Left            =   90
         ScaleHeight     =   4005
         ScaleWidth      =   6075
         TabIndex        =   41
         Top             =   90
         Width           =   6075
         Begin tbrFaroButton.fBoton btBUYDisco 
            Height          =   705
            Left            =   120
            TabIndex        =   42
            Top             =   1110
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1244
            fFColor         =   16777215
            fBColor         =   14737632
            fCapt           =   "Comprar Disco"
            fEnabled        =   -1  'True
            fFontN          =   "Verdana"
            fFontS          =   10
            fECol           =   5452834
         End
         Begin tbrFaroButton.fBoton btBuyCancion 
            Height          =   705
            Left            =   90
            TabIndex        =   43
            Top             =   1590
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1244
            fFColor         =   16777215
            fBColor         =   14737632
            fCapt           =   "Comprar Cancion"
            fEnabled        =   -1  'True
            fFontN          =   "Verdana"
            fFontS          =   10
            fECol           =   5452834
         End
         Begin tbrFaroButton.fBoton btSalir 
            Height          =   705
            Left            =   90
            TabIndex        =   54
            Top             =   1980
            Visible         =   0   'False
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   1244
            fFColor         =   16777215
            fBColor         =   14737632
            fCapt           =   "Salir"
            fEnabled        =   -1  'True
            fFontN          =   "Verdana"
            fFontS          =   10
            fECol           =   5452834
         End
         Begin VB.Label lblInfoPreview 
            BackColor       =   &H00000000&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   495
            Left            =   3750
            TabIndex        =   60
            Top             =   2790
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Image imgExtraObjeto 
            Height          =   495
            Left            =   4590
            Top             =   2820
            Width           =   705
         End
         Begin VB.Image picExtraObjeto 
            Height          =   765
            Left            =   3720
            Stretch         =   -1  'True
            Top             =   2700
            Width           =   1635
         End
         Begin VB.Image ImgSelecVIP 
            Height          =   375
            Left            =   2910
            Top             =   2100
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblCanciones2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de canciones"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   345
            Index           =   0
            Left            =   2790
            TabIndex        =   53
            Top             =   1050
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2160
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
            Height          =   255
            Left            =   630
            TabIndex        =   52
            Top             =   3510
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   750
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
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   3510
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   690
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
            Left            =   1320
            TabIndex        =   50
            Top             =   3540
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
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
            Left            =   1920
            TabIndex        =   49
            Top             =   3570
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Image imageFONDO 
            Height          =   660
            Index           =   0
            Left            =   5130
            Stretch         =   -1  'True
            Top             =   150
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Image TapaCD 
            Height          =   465
            Index           =   0
            Left            =   5250
            Stretch         =   -1  'True
            Top             =   240
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Image imgDiscoSEL 
            Height          =   495
            Left            =   210
            Top             =   2850
            Width           =   675
         End
         Begin VB.Image imgListaSong 
            Height          =   1995
            Left            =   2490
            Top             =   60
            Width           =   2565
         End
         Begin VB.Image imgFondoDiscoSel 
            Height          =   645
            Left            =   120
            Top             =   2820
            Width           =   825
         End
         Begin VB.Label lblCanciones 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Lista de canciones"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   0
            Left            =   2730
            TabIndex        =   48
            Top             =   750
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.Label lblDATA 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Data.txt"
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
            Height          =   225
            Left            =   60
            TabIndex        =   47
            Top             =   3720
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblDATA2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "data.txt"
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
            Height          =   225
            Left            =   660
            TabIndex        =   46
            Top             =   3660
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Image cmdTouchArriba 
            Height          =   360
            Left            =   660
            Top             =   90
            Visible         =   0   'False
            Width           =   540
         End
         Begin VB.Image cmdTouchAbajo 
            Height          =   360
            Left            =   60
            Top             =   60
            Visible         =   0   'False
            Width           =   510
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
            Left            =   2490
            TabIndex        =   45
            Top             =   1740
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.Image imgSELEC 
            Height          =   375
            Left            =   2490
            Top             =   2100
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Image imgSALIR 
            Height          =   375
            Left            =   3300
            Top             =   2100
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label lblXY1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            Caption         =   "2/34"
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
            Left            =   5190
            TabIndex        =   44
            Top             =   1260
            UseMnemonic     =   0   'False
            Visible         =   0   'False
            Width           =   360
         End
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
      Left            =   9480
      TabIndex        =   13
      Top             =   3420
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.PictureBox pVU1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   600
      Left            =   7350
      ScaleHeight     =   600
      ScaleWidth      =   300
      TabIndex        =   12
      Top             =   330
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   8700
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   3900
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4290
      IntegralHeight  =   0   'False
      ItemData        =   "frmINDEX.frx":10B7C
      Left            =   6960
      List            =   "frmINDEX.frx":10BBF
      TabIndex        =   10
      Top             =   4560
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.PictureBox picVideo 
      BackColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   8430
      ScaleHeight     =   255
      ScaleWidth      =   210
      TabIndex        =   9
      Top             =   3900
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox frTEMAS 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   6270
      ScaleHeight     =   285
      ScaleWidth      =   2895
      TabIndex        =   36
      Top             =   2730
      Width           =   2895
      Begin VB.Label T 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del TEMA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   37
         Top             =   30
         Width           =   1710
      End
   End
   Begin VB.PictureBox frModoVideo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   6300
      ScaleHeight     =   1125
      ScaleWidth      =   2895
      TabIndex        =   34
      Top             =   1260
      Width           =   2895
      Begin VB.Label L 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del artista - nombre del disco"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   35
         Top             =   0
         Width           =   3720
      End
   End
   Begin VB.Image imgTapaRankBUP 
      Height          =   660
      Left            =   9240
      Stretch         =   -1  'True
      Top             =   2370
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgTapaDefBUP 
      Height          =   660
      Left            =   9210
      Stretch         =   -1  'True
      Top             =   1620
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgUNSELBUP 
      Height          =   660
      Left            =   9240
      Stretch         =   -1  'True
      Top             =   870
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image imgSELBUP 
      Height          =   660
      Left            =   9150
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Image cmdTouchAbajo2 
      Height          =   375
      Left            =   7860
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdTouchArriba2 
      Height          =   375
      Left            =   7380
      Top             =   3720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image imgSelec2 
      Height          =   375
      Left            =   6900
      Top             =   3690
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Image cmdPagAt 
      Height          =   615
      Left            =   7710
      Top             =   3060
      Width           =   735
   End
   Begin VB.Image cmdPagAd 
      Height          =   615
      Left            =   8640
      Top             =   3120
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
      Left            =   7050
      TabIndex        =   2
      Top             =   90
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
      Left            =   6285
      TabIndex        =   1
      Top             =   2460
      Visible         =   0   'False
      Width           =   2895
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
      Left            =   6300
      TabIndex        =   0
      Top             =   990
      Visible         =   0   'False
      Width           =   2895
   End
End
Attribute VB_Name = "frmIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VU As tbrSoftVumetro.tbrDrawVUM
Public WithEvents MP3 As tbrPlayer02.MainPlayer
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

'Dim WithEvents TF As tbrFOCUS.clsFOCUS

'me cago en la mierda. Siguen dos canciones al mismo tiempo !!!
Dim IenPlenaCancion(3) As Long 'cada uno de los hilos de ejecucion
'solo uno puede estar activo!
'=0 sin nada
'=1 menor a segFade, comenzando cancion
'=2 en plena cancion despues de 1 y antes de 3
'=3 en los segundos finales bajando el volumen

Dim WithEvents GK As tbrGetKeys
Attribute GK.VB_VarHelpID = -1

'upManu
Private EstoyEnDisco As Long 'me dice si estoy dentro de un disco en el modo nuevo
'=0 es en portadas
'=1 es en lista de canciones
'=2 estoy pasando de mun lado a otro

Public OkInState1 As Long 'presiones de la tecla ok en el modo SuperSel
'esto para ignorar la primera siempre
Private VengoDeCarrito As Boolean
'Public Function PonerFoco()
'    TF.PonerFoco
'End Function

'cuenta el tiempo que esta apertada la tecla del carrito antes de hacer el keyUp
Private TimePressTeclaCart As Single
Private TimePressTeclaOK As Single

Private Function EnQueFilaEstoy() As Long
    'es la fila uno si es la primera
    'la barra invertida devuelve solo la parte entera!!!
    EnQueFilaEstoy = (nDiscoSEL \ TapasMostradasH) + 1
    tERR.Anotar "acaa", nDiscoSEL, TapasMostradasH
End Function

Private Sub btBuyCancion_Click()
'    If PachaMode = 10000 Then
        tERR.Anotar "eaaa"
        Carrito.AddFile lblCanciones(selDiscoI(-3)).Tag, False, PerfilActual
        
        Timer3.Enabled = False
        VengoDeCarrito = True
        frmCarrito.Show 1
        tERR.Anotar "eaab"
        Timer3.Enabled = True
'    End If
'
'    If PachaMode = 11000 Then
'        Carrito.AddFile lblCanciones(selDiscoI(-3)).Tag
'
'        Timer3.Enabled = False
'        VengoDeCarrito = True
'        frmCarritoPacha.Show 1
'        Timer3.Enabled = True
'    End If
End Sub

Private Sub btBUYDisco_Click()
'    If PachaMode = 10000 Then
        Carrito.AddFolder txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ","), PerfilActual
        
        Timer3.Enabled = False
        VengoDeCarrito = True
        frmCarrito.Show 1
        Timer3.Enabled = True
'    End If
'
'    If PachaMode = 11000 Then
'        Carrito.AddFile lblCanciones(selDiscoI(-3)).Tag
'
'        Timer3.Enabled = False
'        VengoDeCarrito = True
'        frmCarritoPacha.Show 1
'        Timer3.Enabled = True
'    End If
'
End Sub

Private Sub btSalir_Click()
    If PachaMode = 11000 Then UnSuperSel
End Sub

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
        If EsVideo And Salida2 = False Then 'si estoy pasando un video en el monitor
            Form_KeyDown TeclaDER, 0
        Else
            'si estoy en modo 5 teclas y tengo habilitado el touch screen
            'los botones de los costados deberían pasar página
            EsModo5PeroLabura46 = (EsVideo And (Salida2 = False) And (IsMod46Teclas = 5))
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                PasarPaginaAdelante
            Else
                Form_KeyDown TeclaPagAd, 0
            End If
        End If
    End If
End Sub

Private Sub cmdPagAd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchderechaapretado")
    cmdPagAd.Picture = LoadPicture(IMF)
End Sub

'Private Sub cmdPagAd_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acad", KeyCode, Shift
'End Sub

Private Sub cmdPagAd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchderechanormal")
    cmdPagAd.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdPagAt_Click()
    If MostrarTouch Then
        tERR.Anotar "acak", EstoyEnDisco
        'si tengo videos en la pantalla de la pc no paso pagina, paso solo disco
        If EsVideo And Salida2 = False Then
            Form_KeyDown TeclaIZQ, 0
        Else
            'si estoy en modo 5 teclas y tengo habilitado el touch screen
            'los botones de los costados deberían pasar página
            EsModo5PeroLabura46 = (EsVideo And (Salida2 = False) And (IsMod46Teclas = 5))
            If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                PasarPaginaAtras
            Else
                Form_KeyDown TeclaPagAt, 0
            End If
        End If
    End If
End Sub

'Private Sub cmdPagAt_KeyDown(KeyCode As Integer, Shift As Integer)
'    Form_KeyDown KeyCode, Shift
'    tERR.Anotar "acae", KeyCode, Shift
'End Sub

Private Sub cmdPagAt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchizqapretado")
    cmdPagAt.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdPagAt_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchizqnormal")
    cmdPagAt.Picture = LoadPicture(IMF)
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
    IMF = ExtraData.getDef.getImagePath("tocuharribaelegido")
    cmdTouchArriba.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchArriba_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("tocuharribacomun")
    cmdTouchArriba.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchAbajo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchabajoelegido")
    cmdTouchAbajo.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchAbajo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchabajocomun")
    cmdTouchAbajo.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchArriba2_Click()
    Form_KeyDown TeclaIZQ, 0
End Sub

Private Sub cmdTouchArriba2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("tocuharribaelegido")
    cmdTouchArriba2.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchArriba2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("tocuharribacomun")
    cmdTouchArriba2.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchAbajo2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchabajoelegido")
    cmdTouchAbajo2.Picture = LoadPicture(IMF)
End Sub

Private Sub cmdTouchAbajo2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("touchabajocomun")
    cmdTouchAbajo2.Picture = LoadPicture(IMF)
End Sub


Private Sub Form_Activate()
    On Error GoTo regERR
    
    'si esta usando interfase pasar los mensajes aqui
    If (LCs3 = "1") Then s3.HwndMsg = txtS3.HWND
    
    If TengoBluetooth Then
        BTM.UseEventMSG tBT.HWND
    End If
    
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
    
    If TengoUSB Then UB.UseEventMSG TUsb.HWND
    
    Exit Sub
regERR:
    If Err.Number = 5 Then
        tERR.AppendSinHist "SetFOC"
    Else
        tERR.Anotar "errACAP"
        tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acap"
    End If
End Sub

Public Sub StartVu(sModo As String) 'empezar a medir sonido

    tERR.Anotar "SV01", sModo, HabilitarVUMetro
    
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
            pVU1.Top = 90 'frDiscos.Top
            pVU1.Height = frDiscos.Height + picFondo2.Height
            If EstoyEnModoVideoMiniSelDisco Then
                pVU3.Left = frModoVideo.Left - pVU3.Width
            Else
                pVU3.Left = Me.Width - pVU2.Width
            End If
            'pVU2.Width = AnchoBarra
            pVU3.Top = pVU1.Top
            
            pVU2.Top = pVU1.Top
            pVU2.Left = pVU1.Left
            
            pVU4.Top = pVU3.Top
            pVU4.Left = pVU3.Left
            
            pVU2.Height = pVU1.Height
            pVU3.Height = pVU1.Height
            pVU4.Height = pVU1.Height
            
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
    picFondo.ZOrder 'para que tape el vumetro en casos especiales
    If HabilitarVUMetro Then VU.NotifyResizeVUM
    
End Sub

Private Sub ProcessKeyCoin(Tecla As Integer, isDown As Long)
    'isDown puede ser
    '0 es up
    '1 es down
    '2 viene de la api que no sabe
    
    'este dentro o fuera de un disco no se debe mostrar
    'si estoy en la lista de discos le aviso si hay musica vip
    'ademas tego en cuenta que sea un disco de musica, los de wallpapers
    If EstoyEnDisco = 1 And PerfilActual = 1 Then
        lblNOCREDIT.Caption = getStrMusicaVIP 'deja vacio si no esta activado el vip o pone lo que va
    Else
        lblNOCREDIT.Caption = ""
    End If
    '***********************************************************
    'si es 0 o 1 y yo uso los 2 ignorar para que no duplique!!!!
    If GK.IsLisen Then
        If isDown = 0 Or isDown = 1 Then Exit Sub
    End If
    '***********************************************************
    
    If Tecla = TeclaNewFicha Or Tecla = TeclaNewFicha2 Then
        If CREDITOS > MaximoFichas Then
            LedEvent "ActionLedMuchoCredito"
            Exit Sub
        Else
            'apagar el fichero electronico
            LedEvent "ActionLedPocoCredito"
        End If
    End If
    
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
    
    '*******************************************************
    LastTecla = CLng(KeyCode)
    'pase lo que pase registrar
    Select Case KeyCode
        Case TeclaIZQ: TECLAS_PRES = TECLAS_PRES + "1"
        Case TeclaDER: TECLAS_PRES = TECLAS_PRES + "2"
        Case TeclaOK
            TECLAS_PRES = TECLAS_PRES + "3"
            'cuento cuanto tiempo se aprieta para entrar al tema VIP
            If TimePressTeclaOK = -1 Then TimePressTeclaOK = Timer
        Case TeclaESC: TECLAS_PRES = TECLAS_PRES + "4"
        Case TeclaPagAd: TECLAS_PRES = TECLAS_PRES + "5"
        Case TeclaPagAt: TECLAS_PRES = TECLAS_PRES + "6"
    End Select
    
    TECLAS_PRES = Right(TECLAS_PRES, 20)
    lblTECLAS = TECLAS_PRES
    
    'EJECUTAR ALGO SI CORRESPONDE
    VerClaves TECLAS_PRES
    '*******************************************************
    
    '*******************************************************
    'y si no es una ficha la que se esta cargando
    'aqui se regsitran las presiones de las teclas elegidas
    Dim PagNum As Long
    
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    'ver si es o no numpad
    If IsKeyPad(Me.HWND) Then
        'la falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        RealKeyCode = KeyCode
        If KeyCode = 13 Then
            RealKeyCode = 108
            'tambien controlo el tiempo si el tipo configuro la tecla enter del teclado numerico
            If TeclaOK = 108 Then
                If TimePressTeclaOK = -1 Then TimePressTeclaOK = Timer
            End If
        End If
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
        
    EsModo5PeroLabura46 = (EsVideo And (Salida2 = False) And (IsMod46Teclas = 5))
    
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
            Timer3.Enabled = False
            MostrarCursor True
            frmREG2.Show 1
            MostrarCursor False
            Timer3.Enabled = True
        Case vbKeyF4
            If Shift = 4 Then
                Unload Me
                End
            End If
        Case vbKeyF5
            my_MEM.SetMomento "Apreto F5"
            tERR.AppendSinHist "F5: " + vbCrLf + my_MEM.GetFullDetalles
        Case vbKeyF9 'mostrar la ayuda rápida
            Timer3.Enabled = False
            MostrarCursor True
            frmQuikHelp.Show 1
            MostrarCursor False
            Timer3.Enabled = True
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
                    If EsKar Then
                        ToSec = (MP3.PositionInSec(2) * 1000) + 10000
                        MP3.SeekTo CStr(ToSec), 2
                    Else
                        ToSec = (MP3.PositionInSec(IAA) * 1000) + 10000
                        MP3.SeekTo CStr(ToSec), IAA
                    End If

                End If
                
            End If
        'subir o bajar volumen
        Case TeclaBajaVolumen
            If MP3.IsPlaying(IAA) Then
                If CORTAR_TEMA(IAA) = False Then 'TEMA PAGO
                    If VolumenIni <= 5 Then
                        MP3.Volumen(IAA) = 0
                    Else
                        MP3.Volumen(IAA) = VolumenIni - 5
                    End If
                    VolumenIni = MP3.Volumen(IAA)
                Else 'TEMA GRATUITO VARIA VOLUMEN 2
                    If VolumenIni2 <= 5 Then
                        MP3.Volumen(IAA) = 0
                    Else
                        MP3.Volumen(IAA) = VolumenIni2 - 5
                    End If
                    VolumenIni2 = MP3.Volumen(IAA)
                End If
            End If
        Case TeclaSubeVolumen
            If MP3.IsPlaying(IAA) Then
                If CORTAR_TEMA(IAA) = False Then 'TEMA PAGO
                    If VolumenIni >= 95 Then
                        MP3.Volumen(IAA) = 100
                    Else
                        MP3.Volumen(IAA) = VolumenIni + 5
                    End If
                    VolumenIni = MP3.Volumen(IAA)
                Else 'TEMA GRATUITO
                    If VolumenIni2 >= 95 Then
                        MP3.Volumen(IAA) = 100
                    Else
                        MP3.Volumen(IAA) = VolumenIni2 + 5
                    End If
                    VolumenIni2 = MP3.Volumen(IAA)
                End If
            End If
        
        Case TeclaPagAd
            
            If EstoyEnDisco = 0 Then
                'es para abajo en el modo 5 y pagina adelante de el modo 46
                
                If EsModo5PeroLabura46 Then
                    'esto confirma que es modo 5
                    tERR.Anotar "acax"
                    Form_KeyDown TeclaDER, 0
                End If
                
                If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                    PasarPaginaAdelante
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
                If PachaMode = 10000 Then selDiscoI -1
                If PachaMode = 11000 Then btBuyCancion_Click
            End If
        
        Case TeclaCarrito
            'cuanto cuanto lo tiene apretado para ver si entro cargando lo elegido al carrito o no
            'aqui empiezo a contar el tiempo
            If VendoMusica Then
                If TimePressTeclaCart = -1 Then TimePressTeclaCart = Timer
            End If
        Case TeclaPagAt
            If EstoyEnDisco = 0 Then
                If EsModo5PeroLabura46 Then
                    tERR.Anotar "acbd"
                    'esto confirma que es modo 5
                    Form_KeyDown TeclaIZQ, 0
                End If
                If IsMod46Teclas = 46 Then 'EN ESTE CASO NO ES LO MISMO 'Or EsModo5PeroLabura46 Then
                    PasarPaginaAtras
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
                If PachaMode = 10000 Then selDiscoI -2
                If PachaMode = 11000 Then btBUYDisco_Click
            End If
        Case TeclaConfig
             Timer3.Enabled = False
             frmConfig.Show 1
             Timer3.Enabled = True
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
                
                
            End If
            
            If EstoyEnDisco = 1 Then
                If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                    UnSuperSel
                Else
                    If selDiscoI(-2) = -99 Then 'todo esta elegido!
                        UnSuperSel
                    End If
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
                
                
            End If
            
            If EstoyEnDisco = 1 Then
                If IsMod46Teclas = 5 And EsModo5PeroLabura46 = False Then
                    UnSuperSel
                Else
                    If selDiscoI(-1) = -99 Then 'todo esta elegido!
                        UnSuperSel
                    End If
                End If
            End If
            
        Case TeclaCerrarSistema
            YaCerrar3PM
        Case TeclaESC
            tERR.Anotar "acdo"
            'si estoy fuera de los discos solo me importa si estoy en modo video dentro de un disco
            If EstoyEnDisco = 0 Then
                If ModoVideoSelTema Then 'esta eligiendo canciones dentro del disco
                    AcomodarModoTexto 1
                    ModoVideoSelTema = False 'ya no esta mas!!
                Else
                    'en cualquier otro caso el escape pasa de un ritmo a otro
                    goNextRitmo
                End If
            End If
            
            'si estoy dentro de un disco salgo
            If EstoyEnDisco = 1 Then
                UnSuperSel
            End If
                
        Case vbKeyF12
            MostrarCursor True
    End Select
    
FinKD:
    
    'SecSinTecla = 0
    'ahora (set08) hay una opcion de protector que mueve los discos y lo hace mandando
    'un key_Dow, teclaDer, 0 que pone el segundero en cero. Ya tengo uno en el keyUp
    'por lo tanto este no es necesario
    
    Exit Sub
    
FallaKD:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acas"
    Resume Next

End Sub

Private Sub PasarPaginaAtras()
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
End Sub

Private Sub PasarPaginaAdelante()
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
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Local Error GoTo FallaKD
    
    tERR.Anotar "acds", KeyCode, Shift
    'la verdadera tecla debe mostrar si es una tecla del teclado numerico
    Dim RealKeyCode As Integer
    
    If IsKeyPad(Me.HWND) Then
        'lasa falla reconocida por microsoft es de la tecla enter
        'sea cual sea sale el keycode 13 por mas que sea la del keypad
        'que es el 108
        If KeyCode = 13 Then
            RealKeyCode = 108
        End If
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
    
    'si estoy en fullscreen entonces hay cosas que no se pueden hacer!!!
    'por ejemplo el carrito noooooo
    If EsVideo And (Salida2 = False) And vidFullScreen Then
        'ni bosta de carrito!
        TimePressTeclaCart = 0
        'no puede elegir canciones VIP
        TimePressTeclaOK = 0
    Else
        If RealKeyCode = TeclaCarrito And VendoMusica Then 'solo si activo la venta de musica
            
            TimePressTeclaCart = Timer - TimePressTeclaCart
            
            'puede entrar agregando o no al carrito
            'si deja apretado 2 segundos la tecla de carrito entra sin agregar!
            If TimePressTeclaCart <= 1.5 Then
                'ver si esta en la lista de discos o en las carpetas
                If EstoyEnDisco = 1 Then
                    Dim cancionD As String
                    cancionD = lblCanciones(selDiscoI(-3)).Tag
                    Carrito.AddFile cancionD, False, PerfilActual
                Else
                    'si el tipo desea que se autodetecten los perfiles debo ver el perfil de este
                    'en general el perfil se calcula al ingresar!!!
                    Dim lstDisco() As String, esteDisco As String
                    Dim PerfilEncontrado As Long
                    esteDisco = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
                    
                    If VentaExtras Then
                        PerfilEncontrado = 1 'para que busque el perfil
                        lstDisco = ObtenerArchMM(esteDisco, False, PerfilEncontrado)
                        PerfilActual = PerfilEncontrado
                    End If
                    
                    Carrito.AddFolder esteDisco, PerfilActual
                End If
            End If
            'entro al carrito haya agregado o no
            'mostrar para que vea, corrija, revise y copie o siga
            frmIndex.Timer3.Enabled = False
            VengoDeCarrito = True
            frmCarrito.Show 1
            frmIndex.Timer3.Enabled = True
            TimePressTeclaCart = -1
            GoTo FinUP
        End If
    End If
    'el 108 es el enter del numerico y anda para el keyDown pewro yo no quiero alli!!!
    'asi que como me llega un 13 lo tomo tambien cuando se pida un 108 como respuesta!!
    If EsVideo And (Salida2 = False) And vidFullScreen Then
        'ni bosta de tecla ok!
        GoTo FinUP
    End If
    
    If RealKeyCode = TeclaCancionVIP Then
        RealKeyCode = TeclaOK
        TimePressTeclaOK = 99 'hago como que esta apretado hace mucho
    End If
    
    If (RealKeyCode = TeclaOK) Or (TeclaOK = 108 And RealKeyCode = 13) Then
        'al salir del carrito quiere entrar al disco
        If VengoDeCarrito Then
            VengoDeCarrito = False
            GoTo FinUP
        End If
        
        If EstoyEnDisco = 0 Then
            'si estoy en video
            'saber si estoy eligiendo tema. Si no estoy en disco
            tERR.Anotar "accv", nDiscoGral, nDiscoSEL, ModoVideoSelTema
            If ModoVideoSelTema Then 'estoy con las canciones desplegadas en modo texto de algun disco
                'si esta en fullscreen NO EJECUTAR!!!
                'solo si no sale por la segunda salida!!!
                If EsVideo And vidFullScreen And Salida2 = False Then GoTo FinUP 'fin keydown
                'si no dice salir cargar tema
                tERR.Anotar "accw", T(TemaElegidoModoVideo)
                If T(TemaElegidoModoVideo) = TR.Trad("SALIR%99%") Or _
                    T(TemaElegidoModoVideo) = TR.Trad("No hay temas%98%Es un disco sin canciones%99%") Then
                    
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
                    
                    'ver si se apreto 2 o mas segundos (o si apreto la tecla de vip directo (99))
                    If TimePressTeclaOK <> 99 Then
                        TimePressTeclaOK = Timer - TimePressTeclaOK
                    End If
                    Dim isVip As Boolean 'ver si lo dejo mucho apretado es un VIP
                    isVip = (TimePressTeclaOK > 1.5) And PrecNowVIP > 0
                    TimePressTeclaOK = -1 'listo ya lo use
                    Dim S36 As Long
                    S36 = TrataEjecutarTema(temaElegido, isVip) 'no mando perfil actual por que en la lista de modo texto es si o si perfil=1=3pm base
                    tERR.Anotar "accx", S36
                    If S36 = 2 Then 'ya esta ejecutando otra cosa!
                        'volver a elegir discos
                        AcomodarModoTexto 1
                        ModoVideoSelTema = False
                    End If
                End If
            
            Else 'ELSE DEL MODOVIDEO SEL TEMA
                
                TimePressTeclaOK = -1 'no se usa aca
                
                'ver si hay que mostrar el frm
                'o estamos en MODO VIDEO
                tERR.Anotar "acdd"
                'ver si es video debería desplegar los temas del disco elegido
                'en modo de texto
                'pero si estoy viendo el video en salida2 es video sera verdadero
                'pero de todas formas no veo als lista de texto y sigo igual
                'solo si esvideo y necesito el modo texto del video!!!!
                If EsVideo And (Salida2 = False) Then
                    AcomodarModoTexto 2
                    'cargar los temas multimedia en t()
                    'es una matriz global
                    'en la 6.3 era nDiscoGral+1!!!
                    tERR.Anotar "acde", MATRIZ_DISCOS(nDiscoGral)
                    UbicDiscoActual = txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ",")
                    'encontrar todos los archivos *.mp3, *.avi, *.mpg, *.mpeg, etc
                    tERR.Anotar "acdf", UbicDiscoActual
                    ReDim Preserve MATRIZ_TEMAS(0)
                    'OM- el tipo entro a un disco en modo texto (hay un video en el monitor), _
                        si el origen o el disco ya indica un perfil usarlo y si no automático
                    MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual)
                    'forzado el perfil es 3pm, quiozas mas adelante en el modo texto haya previews de wallpapers y otros
                    PerfilActual = 1
                    tERR.Anotar "acdg", UBound(MATRIZ_TEMAS)
                    If UBound(MATRIZ_TEMAS) = 0 Then
                        T(0) = TR.Trad("No hay temas%98%Es un disco sin canciones%99%")
                        SelTema 0
                        ModoVideoSelTema = True
                        tERR.Anotar "acdh", nDiscoSEL, nDiscoGral
                        Exit Sub
                    End If
                    tERR.Anotar "acdi"
                    T(0) = TR.Trad("SALIR%99%")
                    '----------------------------
                    'a daniel cruz le da un error como si se volviera a cargar algo que esta cargado
                    'por lo tanto tengo que poner un manejador de error aqui, unico lugar en que se carga esto
                    For Each LLL In frmIndex.T 't(x) son los labels con las canciones en modo video
                        If LLL.Index > 0 Then Unload LLL
                    Next
                    '----------------------------
                    tERR.Anotar "acdj", UBound(MATRIZ_TEMAS)
                    For AA = 1 To UBound(MATRIZ_TEMAS)
                        tERR.Anotar "acdk", AA, MATRIZ_TEMAS(AA)
                        Load T(AA)
                        T(AA) = fso.GetBaseName(txtInLista(MATRIZ_TEMAS(AA), 1, "#"))
                        T(AA).Top = T(AA - 1).Top + T(AA - 1).Height
                        T(AA).Left = T(AA - 1).Left
                        T(AA).Visible = True
                    Next
                    tERR.Anotar "acdl", nDiscoSEL, nDiscoGral
                    TemaElegidoModoVideo = 0
                    SelTema 0
                    ModoVideoSelTema = True
                    
                Else 'ELSE DEL ESVIDEO AND SALIDA2
                    'ndiscosel puede valer 99999 ver
                    If nDiscoSEL <> 99999 Then
                        tERR.Anotar "acdm", nDiscoSEL, nDiscoGral
                        SuperSel nDiscoSEL
                    Else
                        tERR.Anotar "acdm3"
                    End If
                End If
            End If
        End If
        
        If EstoyEnDisco = 1 Then 'estoy dentro de un disco
            If TimePressTeclaOK <> 99 Then
                TimePressTeclaOK = Timer - TimePressTeclaOK
            End If
            Dim isVip2 As Boolean 'ver si lo dejo mucho apretado es un VIP
            isVip2 = (TimePressTeclaOK > 1.5) And PrecNowVIP > 0
            TimePressTeclaOK = -1 'listo ya lo use
            
            OkInState1 = OkInState1 + 1 'la primera no va!
            If OkInState1 > 1 Then EjecutarDeTouch isVip2
        End If
    End If
    
FinUP:
    VerClaves TECLAS_PRES
    SecSinTecla = 0
    
    Exit Sub

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
    'y otras que dependen de lo que se muestra
    
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
        
        frModoVideo.Height = frDiscos.Height - lblModoVideo.Height - 150 ' Me.Height - (lblModoVideo.Top + lblModoVideo.Height + _
                                          picFondo.Height + picFondo2.Height) - 333
                                          
        'frModoVideo.Height = frDiscos.Height - 200 - _
                            (lblModoVideo.Top + lblModoVideo.Height) + _
                            picFondo.Height
                            '+ picFondo2.Height
                            
        Degrade frModoVideo, 1
        
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
        Degrade frModoVideo, 1
        
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
            Degrade frTEMAS, 1
        Else
            frModoVideo.Height = frDiscos.Height / 5
            lblTEMAS.Top = frModoVideo.Top + frModoVideo.Height + 50
            lblTEMAS.Left = lblModoVideo.Left
            frTEMAS.Top = lblTEMAS.Top + lblTEMAS.Height
            frTEMAS.Height = frDiscos.Height - lblModoVideo.Height - frModoVideo.Height - lblTEMAS.Height - 75
            Degrade frTEMAS, 1
            Degrade frModoVideo, 1
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

'busca en la lista del disco abierto el elemento y lo mada al tratarEjTema que hace otros controles
Private Sub EjecutarDeTouch(Optional ToVIP As Boolean = False)
    'opcionalmente puedo avisar si es un tema VIP
    
    On Local Error GoTo errDTouch
    tERR.Anotar "caaa"
    
    'si esta programado sin musica que no haga nada
    If NOMUSIC Then 'no es fonola
        If ShowDemoMusic Then 'pasa las canciones como demo 20 segundos
            If CREDITOS < CreditForTestMusic Then Exit Sub 'si se configuro exigir creditos para pasar muestras (igual son gratis)
            If MaxListaTestMusic > 0 Then 'si es cero permite todo
                If tLST.GetLastIndex >= MaxListaTestMusic Then Exit Sub 'hay mas canciones en lista que las permitidas
            End If
        Else
            Exit Sub
        End If
    End If
    
    Dim Fg As Long
    Fg = selDiscoI(-3) 'es el numero de cancion!
    tERR.Anotar "caab", Fg
    
    'ya no hay nada mas(11/09/07)
    If Fg = -99 Then
        UnSuperSel
        Exit Sub 'no salia antes
    End If
    
    Dim S37 As Long
    S37 = TrataEjecutarTema(lblCanciones(Fg).Tag, ToVIP, PerfilActual) 'mando ademas el perfil ya _
            que los rigtones y los MP3s comparten extencion y no son lo mismo
            
    tERR.Anotar "caac", S37
    If S37 = 1 Then 'si no alcanza el credito avisar!
        lblNOCREDIT.Caption = TR.Trad("CREDITO INSUFICIENTE%99%")
        Exit Sub
    End If
    
    If S37 = 4 Or S37 = 5 Or S37 = 6 Or S37 = 7 Then
        'me voy por que lo que esta abajo depende de que sea musica!
        Exit Sub
    End If
    
    'o algo se ejecuta o va a la lista seguro
    If BloquearMusicaElegida Then
        tERR.Anotar "caac2", Fg
        lblCanciones(Fg).Visible = False
        lblCanciones(Fg).Tag = "" 'para que no lo elija de nuevo
    Else 'de alguna forma tengo que decirle que se eligio!
        tERR.Anotar "caac3", Fg
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
    GetIntervalS3 = s3.GetInterval
End Function

Public Sub SetIntervalS3(NewIntervalS3 As Long)
    s3.SetInterval NewIntervalS3
End Sub

Private Sub esperar(n As Single)
    n = Timer + n
    Do While Timer < n
        DoEvents
    Loop
End Sub

Private Sub Form_Load()
    On Error GoTo NoLoadIndex
    
    Traducir 'Agregado por el complemento traductor
    
    'escribir el hwnd que escucho lo que me dice desde otros exes
    'EscribirArch1Linea2 AP + "hd.wan", getWAN(tUltra.HWND)
    
    'si no se cierra bien se empiezan a acumular
'    'lanzar el asistant si existe!
'    If fso.FileExists(AP + "tius.exe") Then
'        Shell AP + "tius.exe", vbNormalFocus
'    End If
    
    tERR.Anotar "some" 'que empieze a registrar
    
    Degrade frTEMAS, 1
    Degrade frModoVideo, 1
    
    'no puedo hacer referencia a ningun objeto de frmIndex por que lo cargaria antes de tiempo
    IMF = ExtraData.getDef.getImagePath("vumetroprendido")
    'temporalmente uso pVu1 pero puede ser cualquiera es solo por que no se cuanto tiene de ancho la imagen segun el skin
    pVU1.AutoSize = True
    pVU1.Picture = LoadPicture(IMF)
    AnchoBarra = pVU1.Width
    pVU1.Picture = LoadPicture
    pVU1.AutoSize = False
    
    lblNOCREDIT.Caption = ""
    lblNOCREDIT.Visible = True
    
    tERR.Anotar "sVU01"
    Me.BackColor = vbBlack
    '************************************
    'aqui estaba lo de la licencia!!!
    '*************************************
    
    TengoBluetooth = CBool(CLng(LeerConfig("TengoBluetooth", "0")))
    TengoUSB = CBool(CLng(LeerConfig("TengoUSB", "1")))
    TengoCD = CBool(CLng(LeerConfig("TengoCD", "0")))
    
    Me.Caption = "QTIO232087402198347"
    tERR.Anotar "eaaq", TengoBluetooth, TengoUSB, TengoCD
    
    If TengoUSB Then
        Set UB = New tbrDRIVES.clsDRIVES
        UB.SoloDispositivosUSB = True
        UB.Iniciar Me
    End If
    
    If TengoBluetooth Then
        'indica en el modulo que se usa la referencia al objeto BTManager
        'tbrBtActivex.UsarBluetooth
        tERR.Anotar "eaar22"
        Set BTM = tbrBtActivex.btManager
        tbrBtActivex.SetWindowMsg Me.HWND
        BTM.Initialize
    End If
    
    'ver si hay que buscar cds
    If TengoCD Then
        'mm91
        tERR.Anotar "eaar22A"
        Set CDR = New tbrCD
        tERR.Anotar "eaar22b"
        CDR.DetectarUnidades
        tERR.Anotar "eaar23", CDR.Cantidad
                
        'inicar lo de andres P
        Dim H As Long
        H = CDR.Iniciar
        tERR.Anotar "bgah", H
        Select Case H
            Case 1, 2 'no hay grabadoras!, 'hay mas de una grabadora!
                tERR.AppendLog "bagi"
                TengoCD = False
                'xxxx deberia hacer algo mas pulenta para que el cliente sepa que paso!
                'Exit Sub
            Case Is > 2 'algo no se registro ok por que no puede cargar la dll tbrBurner
                TengoCD = False
                'MsgBox "No se pudo iniciar Modulo de Grabacion de CD." + vbCrLf + _
                    "Asegúrese de tener instalado NET Framework y Nero 7 o superior"
        End Select
    End If
    tERR.Anotar "eaas"
    
    Set VU = New tbrSoftVumetro.tbrDrawVUM
    Dim UAT As String
    
    UAT = LeerConfig("UseAPITecla", "0")
    tERR.Anotar "eaat", UAT
    'la declaro para que pueda saberse el Islisen que es necesario !!!!
    Set GK = New tbrGetKeys
    If UAT <> "0" Then
        'ver que letras necesito
        Dim TMP44 As String
        TMP44 = CStr(TeclaNewFicha) + " " + CStr(TeclaNewFicha2)
        GK.Startlisen TMP44
    End If
    
    tERR.Anotar "eaau", HabilitarVUMetro
    
    EstoyEnDisco = 0
    
    If HabilitarVUMetro Then
        If VU.DispositivosCant = 0 Then
            tERR.AppendLog "Sin PLACA para vumetro!!!"
            HabilitarVUMetro = False 'lo inhabilito!
            'YaCerrar3PM
            'Exit Sub
        Else
            VU.DefinePictureBox pVU1
            VU.DefinePictureBox2 pVU2
            VU.DefinePictureBox3 pVU3
            VU.DefinePictureBox4 pVU4
            
            IMF = ExtraData.getDef.getImagePath("vumetroprendido")
            VU.DefineImage 1, IMF, True
            VU.DefineImage 3, IMF, True
            IMF = ExtraData.getDef.getImagePath("vumetroapagado")
            VU.DefineImage 2, IMF, True
            VU.DefineImage 4, IMF, True
            
            pVU1.ZOrder
            pVU2.ZOrder
            pVU3.ZOrder
            pVU4.ZOrder
            VU.CantCuadros = 20
            VU.CantPic = 10
            VU.ColorBase = vbRed
        End If
    End If
    
    picFondo.ZOrder 'para que tape el vumetro en casos especiales
    
    tERR.Anotar "cMM"
    Set MP3 = New tbrPlayer02.MainPlayer
    
    If ActivarERR Then
        Dim n As String
        n = CStr(Day(Date)) + "." + CStr(Month(Date)) + "." + CStr(Year(Date)) + _
            "." + CStr(Hour(time)) + "." + CStr(Minute(time)) + "." + CStr(Second(time))
        
        
        MP3.ActivaFulLog AP + "REG_MM" + CStr(n) + ".W15"
    End If
    
    'ver si quiere setear todo lo predeterminado en videos
    Dim VDV As String, GDD As String
    VDV = LeerConfig("ValidarDriverVideo", "1")
    tERR.Anotar "Ix100", VDV
    
    GDD = MP3.GetDefaultDevice("MPEGVideo")
    tERR.Anotar "Ix101", GDD
    
    GDD = MP3.GetDefaultDevice("avivideo")
    tERR.Anotar "Ix102", GDD
    
    If VDV <> "0" Then
        MP3.SetDriversVideo
    End If
        
    MP3.DefinePathLogs AP + "regM1.log", AP + "regM2.log"
    
    tERR.Anotar "Ix001"
    'Set TF = New tbrFOCUS.clsFOCUS
    'TF.IntervalTimer = 5000
    'TF.Iniciar Me.Hwnd
    tERR.Anotar "Ix002"
    On Error GoTo MiErr
            
    If MostrarTouch Then
        'imagenes del touch screen
        IMF = ExtraData.getDef.getImagePath("touchizqnormal")
        cmdPagAt.Picture = LoadPicture(IMF)
        IMF = ExtraData.getDef.getImagePath("touchderechanormal")
        cmdPagAd.Picture = LoadPicture(IMF)
        IMF = ExtraData.getDef.getImagePath("botonokcomun")
        imgSELEC.Picture = LoadPicture(IMF)
        IMF = ExtraData.getDef.getImagePath("botonokcomunvip")
        If fso.FileExists(IMF) Then
            ImgSelecVIP.Picture = LoadPicture(IMF)
        Else
            ImgSelecVIP.Picture = LoadPicture("") 'puede que sea skin viejo sin boton VIP
        End If
        
        IMF = ExtraData.getDef.getImagePath("botonsalirnormal")
        imgSALIR.Picture = LoadPicture(IMF)
        
        imgSelec2.Picture = imgSELEC.Picture
        IMF = ExtraData.getDef.getImagePath("tocuharribacomun")
        cmdTouchArriba.Picture = LoadPicture(IMF)
        IMF = ExtraData.getDef.getImagePath("touchabajocomun")
        cmdTouchAbajo.Picture = LoadPicture(IMF)
        cmdTouchArriba2.Picture = cmdTouchArriba.Picture
        cmdTouchAbajo2.Picture = cmdTouchAbajo.Picture
    End If
    
    'ver si es superlicencia y usa otra tapa predeterminada
    IMF = GetTpPred
    tERR.Anotar "Ix002-e", IMF
    
    tbrPassImg1.Picture IMF
    tERR.Anotar "acem", SYSfolder
    
    '*****************************
    '*****************************
    Me.Height = Screen.Height
    Me.Width = Screen.Width
    AjustarFRM Me, 12000, 9000 'solo una vez despues sale todo a proporcion!
    BaseVista 'por unica vez cosas que no cambian
    UpdateVista 'acomodar todo segun variables SIEMPRE DESPUES DE AJUSTAR EL TAMAÑO DE LAS COSAS!
    '*****************************
    '*****************************
    
    'imagenes no cargadas, ver si hay algo configurado para el fondo
    IMF = ExtraData.getDef.getImagePath("FondoDeLasTapas")
    If K.sabseee("3pm") = Supsabseee Then
        If fso.FileExists(GPF("iischu")) Then
            picFondoDisco.PaintPicture LoadPicture(GPF("iischu")), 0, 0, picFondoDisco.Width, picFondoDisco.Height
        Else
            picFondoDisco.PaintPicture LoadPicture(IMF), 0, 0, picFondoDisco.Width, picFondoDisco.Height
        End If
    Else
        picFondoDisco.PaintPicture LoadPicture(IMF), 0, 0, picFondoDisco.Width, picFondoDisco.Height
    End If

    tERR.Anotar "acek", IMF
    
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
    'cargar imagenes para no cargar tantas veces cada vez que se mueve
    
    IMF = ExtraData.getDef.getImagePath("marcodiscocomun")
    imgUNSELBUP.Picture = LoadPicture(IMF)
    
    IND = ExtraData.getDef.GetIndexImage("marcodiscocomun")
    MargSup = ExtraData.getDef.GetFinalMargenSuperiorTra(IND) * AltoTapaDisco / 100
    MargInf = ExtraData.getDef.GetFinalMargenInferiorTra(IND) * AltoTapaDisco / 100
    MargDer = ExtraData.getDef.GetFinalMargenDerechoTra(IND) * AnchoTapaDisco / 100
    MargIzq = ExtraData.getDef.GetFinalMargenIzquierdoTra(IND) * AnchoTapaDisco / 100
    
    IMF = ExtraData.getDef.getImagePath("marcodiscoelegido")
    imgSELBUP.Picture = LoadPicture(IMF)
    
    
    IMF = ExtraData.getDef.getImagePath("taparanking")
    imgTapaRankBUP.Picture = LoadPicture(IMF)
    
    IMF = ExtraData.getDef.getImagePath("tapapredeterminada")
    imgTapaDefBUP.Picture = LoadPicture(IMF)
    
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
    
    tERR.Anotar "TCD(0).TOP", TapaCD(C).Top
    tERR.Anotar "LBL(0).TOP", lblDISCO(C).Top
    
    imageFONDO(0).Picture = imgUNSELBUP.Picture
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
    C = 0
    Do While C < CantDiscos - 1 'si la primera hoja incompleta se carga completa!!
        tERR.Anotar "acez", C
        C = C + 1
        Load TapaCD(C)
        Load lblDISCO(C)
        Load lblDisco2(C)
        Load imageFONDO(C)
        'ya toman el tamaño del original
        
        Dim LineaTopActual As Long
        If C >= TapasMostradasH Then
            LineaTopActual = (AltoTapaDisco * (C / TapasMostradasH)) + (EspacioEntreDiscosV * ((C / TapasMostradasH) + 1))
                        'imageFONDO(c - TapasMostradasH).Top + _
                         imageFONDO(c - TapasMostradasH).Height _
                         EspacioEntreDiscosV
        Else
            LineaTopActual = EspacioEntreDiscosV
        End If
        tERR.Anotar "LTA(" + CStr(C) + ")", LineaTopActual
        If C / TapasMostradasH = C \ TapasMostradasH Then
            'es una tapa al principio de linea!!!!
            lblDISCO(C).Left = IniCentrarH + MargDer
            TapaCD(C).Left = TapaCD(0).Left
            TapaCD(C).Top = LineaTopActual + MargSup
            tERR.Anotar "TCD(" + CStr(C) + ").TOP", TapaCD(C).Top
            If RotulosArriba Then
                lblDISCO(C).Top = LineaTopActual
                tERR.Anotar "LBL(" + CStr(C) + ").TOP", lblDISCO(C).Top
                TapaCD(C).Visible = True
                imageFONDO(C).Visible = True
                If MostrarRotulos Then
'                   TapaCD(c).Top =lblDISCO(c).Top + lblDISCO(c).Height + 50
                    lblDISCO(C).Visible = True
                    lblDisco2(C).Visible = True
                Else
'                   TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + 50
                End If
            Else
'                If MostrarRotulos Then
'                    TapaCD(c).Top = lblDISCO(c - TapasMostradasH).Top + lblDISCO(c - TapasMostradasH).Height + EspacioEntreDiscosV
'                Else
'                    TapaCD(c).Top = TapaCD(c - TapasMostradasH).Top + TapaCD(c - TapasMostradasH).Height + EspacioEntreDiscosV
'                End If
                lblDISCO(C).Top = LineaTopActual + AltoTapaDisco - (2 * MargInf) 'TapaCD(c).Top + TapaCD(c).Height - MargInf '+ 150
                tERR.Anotar "LBL(" + CStr(C) + ").TOP", lblDISCO(C).Top
                TapaCD(C).Visible = True
                imageFONDO(C).Visible = True
                If MostrarRotulos Then
                    lblDISCO(C).Visible = True
                    lblDisco2(C).Visible = True
                End If
            End If
        Else 'el c-1 tiene el mismo top, es cualquiera de una linea que no sea el pri de la izq
            'una tapa comun que se acomoda a la derecha de la anterior
            If RotulosArriba Then
                lblDISCO(C).Left = lblDISCO(C - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH + MargDer
                lblDISCO(C).Top = lblDISCO(C - 1).Top
                tERR.Anotar "LBL(" + CStr(C) + ").TOP", lblDISCO(C).Top
                TapaCD(C).Left = lblDISCO(C).Left
                TapaCD(C).Top = TapaCD(C - 1).Top
                tERR.Anotar "TCD(" + CStr(C) + ").TOP", TapaCD(C).Top
                TapaCD(C).Visible = True
                imageFONDO(C).Visible = True
            Else
                TapaCD(C).Left = TapaCD(C - 1).Left + AnchoTapaDisco + EspacioEntreDiscosH
                TapaCD(C).Top = TapaCD(C - 1).Top
                
                tERR.Anotar "TCD(" + CStr(C) + ").TOP", TapaCD(C).Top
                lblDISCO(C).Left = TapaCD(C).Left
                lblDISCO(C).Top = lblDISCO(C - 1).Top
                tERR.Anotar "LBL(" + CStr(C) + ").TOP", lblDISCO(C).Top
                TapaCD(C).Visible = True
                imageFONDO(C).Visible = True
            End If
            If MostrarRotulos Then
                lblDISCO(C).Visible = True
                lblDisco2(C).Visible = True
            End If
        End If
        
        imageFONDO(C).Picture = imgUNSELBUP.Picture
        imageFONDO(C).Top = TapaCD(C).Top - 2 * MargSup 'TapaCD(c).Top - 150
        imageFONDO(C).Left = TapaCD(C).Left - MargIzq  ' TapaCD(c).Left - 200
        imageFONDO(C).Width = AnchoTapaDisco 'TapaCD(c).Width + MargDer + MargIzq 'TapaCD(c).Width + 400
        imageFONDO(C).Height = AltoTapaDisco + MargSup 'TapaCD(c).Height + MargSup + MargInf 'TapaCD(c).Height + lblDISCO(c).Height + 200
        
        TapaCD(C).ZOrder
        imageFONDO(C).ZOrder
        lblDisco2(C).ZOrder
        lblDISCO(C).ZOrder
    Loop
    'tERR.AppendLog "LISTO TAPAS"
    tERR.Anotar "acfa"
    
    lblV = TR.Trad("versión%99%") + " " + Trim(CStr(App.Major)) + "." + Trim(CStr(App.Minor)) + "." + Trim(CStr(App.Revision))
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
    
    TEMA_REPRODUCIENDO = TR.Trad("Sin reproducción actual%99%")
    TEMA_SIGUIENTE = TR.Trad("No hay próximo tema%98%No hay canciones en la lista de reproducción%99%")
    TEMAS_EN_LISTA = 0
    
    'buscar discos = todas las carpetas en AP\discos\*.*
    'y meterlos en la matriz
        
    'usar el que lee los discos con matrices temporales y _
    sumar todas esas matrics a Matriz_Discos _
    fijarse que el orden no sea alfabetico, solo alfabetico _
    dentro de cada origen de discos
    
    'ya se cargo en el ini!
    'PartOrigenes = Split(Origenes, "*")

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
    '********************************************
    'calcular el ancho total para poder centrar
    Dim WiRit As Long
    'asegurarse que entre en pantalla
    Do
        WiRit = 0
        For AAA = 0 To UBound(PartOrigenes)
            lRITMO(0).Caption = fso.GetBaseName(PartOrigenes(AAA))
            WiRit = WiRit + lRITMO(0).Width + 160
        Next AAA
        
        If WiRit > frDiscos.Width Then
            tERR.Anotar "chichica", WiRit, frDiscos.Width, lRITMO(0).Font.Size
            'no achicar demasiado la letra!
            If lRITMO(0).Font.Size < 9 Then
                Exit Do
            Else
                lRITMO(0).Font.Size = lRITMO(0).Font.Size - 1
                lRITMO2(0).Font.Size = lRITMO(0).Font.Size
            End If
        Else
            Exit Do
        End If
    Loop
    'MsgBox lRITMO(0).Font.Size'para ver a que tamaño cierra
    lRITMO(0).Left = (picFondo2.Width / 2) - (WiRit / 2)
    '********************************************
    
    For AAA = 0 To UBound(PartOrigenes)
        If AAA > 0 Then
            Load lRITMO(AAA)
            Load lRITMO2(AAA)
            
            lRITMO(AAA).Visible = True
            lRITMO2(AAA).Visible = True
        End If
        
        lRITMO(AAA).Caption = fso.GetBaseName(PartOrigenes(AAA))
        lRITMO2(AAA).Caption = lRITMO(AAA).Caption
        
        If AAA > 0 Then lRITMO(AAA).Left = lRITMO(AAA - 1).Left + lRITMO(AAA - 1).Width + 160
        
        lRITMO2(AAA).Left = lRITMO(AAA).Left + 15
        lRITMO2(AAA).Top = lRITMO(AAA).Top + 15
        
    Next AAA
    '
    
    '=============================================================================
    '=============================================================================
    Dim MD As Long
    Randomize
    MD = CLng(Rnd * 49) + 15
    
    tERR.Anotar "001-0063"
    If K.sabseee("3pm") <= CGratuita And UBound(MATRIZ_DISCOS) > MD Then
        'limite de discos
        tERR.Anotar "001-0064"
        MsgBox TR.Trad("Esta es una version demo y no se pueden " + _
            "cargar muchos discos." + vbCrLf + _
            "Para conseguir la versión sin límite de discos " + _
            "envie un e-mail a: %98%Cuando aun no tiene " + _
            "licencia no puede ver todos los discos de musica disponibles " + _
            "en la PC%99%") + "tbrsoft@hotmail.com / tbrsoft@cpcipc.org."
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
    Dim CantNoUsados As Long 'para saber cuanto se eliminan por cuestion de indices en matriz
    CantNoUsados = 0
    For AAA = 0 To UBound(MATRIZ_DISCOS)
        'obtengo la lista de arhivos
        dDI = txtInLista(MATRIZ_DISCOS(AAA), 0, ",")
        If dDI = "_RANK_" Then 'este ni siquiera existe en el disco
            CantMM = 10
        Else
            'OM- al iniciar el sistema se define que discos estan vacios para no _
                cargarlos directamente que quedan mal mostrar discos vacios
            'si desea vender cosas extrañas las bubsco si no nada
            
            If VentaExtras Then
                MMs = ObtenerArchMM(dDI, , 1)
            Else
                MMs = ObtenerArchMM(dDI)
            End If
            tERR.Anotar "caam", AAA, dDI
            'veo que tenga al menos 1!
            CantMM = UBound(MMs)
        End If
        
        If CantMM = 0 Then
            tERR.Anotar "caak", AAA, dDI
            'si se quita aqui en un for que depende del ubound voy a generar errores!!!
            'QuitaIndiceMatriz MATRIZ_DISCOS, AAA
            IsQuitar = IsQuitar + CStr(AAA - CantNoUsados) + " "
            'no espero que entiendas esto de restar la cantidad de usados
            CantNoUsados = CantNoUsados + 1
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
    
    Unload frmINI
    
    
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
    
    'ver si hay validacion por creditos
    VALIDAR = LeerConfig("Validar", "0")
    'se usa la validacion cada vez que descuenta un credito entonces no voy a hacer que lea cada vez
    ValidarCada = LeerConfig("ValidarCada", "3000")
    AvisarAntes = LeerConfig("AvisarAntes", "500")
    
    If VALIDAR Then
        'ver si existe el archivo Creditos Validar
        If fso.FileExists(GPF("radliv")) Then
            'leer el archivo de creditos vaildados
            CreditosValidar = CLng(LeerArch1Linea(GPF("radliv")))
            tERR.Anotar "acfi", CreditosValidar
        Else
            tERR.Anotar "acfj"
            EscribirArch1Linea GPF("radliv"), "0"
            CreditosValidar = 0
            CrearNuevoCodigoValidar 'graba el archivo con un numero al azar
            'lo mantiene hasta que se genera uno nuevo al terminar el periodo de control
        End If
        'ver cual es el máximo y si hay que avisar
    End If
    
    imgExtraObjeto.Stretch = False
    
    tbrPassImg1.IniciarPASS
    
'    'si no tiene el foco ponerlo!!!
'    If TF.GetState <> 1 Then TF.PonerFoco
    
    'lo prendo por mas que no haya protecto configurado por que lo uso para salir de los
    'discos tambien!
    Timer3.Interval = 3000
    
    'si quedaron temas pendientes cargarlos
    Select Case ReINI
        Case "LISTA" 'solo la lista despues del tema actual
            tLST.ListaAbrirDeDisco GPF("casc1001")
            EMPEZAR_SIGUIENTE 3
        Case "NADA"
            'no hacer nada
            'borrar la lista
            'borrra los temas 'y los creditos?
            If fso.FileExists(GPF("casc1001")) Then fso.DeleteFile GPF("casc1001"), True
            Timer1.Interval = 10000
    End Select
    'quito contadores de tiempo
    TimePressTeclaCart = -1
    TimePressTeclaOK = -1
    VerSiTocaVMute
    
    Exit Sub
NoLoadIndex:
MiErr:

    Select Case LCase(tERR.GetLastLog)
        Case "eaar22"
            'no se pudo inicializar el bluetooth correctamente
            MsgBox "No se puede inicializar el bluetooth, se eliminara esta configuración" + vbCrLf + _
                "Asegúrese de tener un dispositivo bluetooth con driver bluesoleil 1.6.4 o superior instalado"
            ChangeConfig "TengoBluetooth", "0"
            Resume Next
        Case "eaar22a"
            'no se pudo inicializar el bluetooth correctamente
            MsgBox "No se puede módulo de grabación de CD, se eliminara esta configuración" + vbCrLf + _
                "Asegúrese de tener NERO 7 o superior instalado y contar con una grabadora de CD"
            ChangeConfig "TengoCD", "0"
            Resume Next
    End Select

    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdu"
    Resume Next
End Sub

Private Sub BeginRoll()
    '--------
    RollCRED.SetInterval 60
    RollCRED.SetVarColor 7
    RollCRED.MaxlargoRenglon = 30
    RollCRED.TextoACola TR.Trad("disfrute su música%99%"), &HA9C8C9       '0 es
    RollCRED.TextoACola TR.Trad("desafios digitales%99%"), &HA9C8C9     '1 es lista de precios
    RollCRED.TextoACola "tbrSoft", &HA9C8C9              '2 es texto de SL al publico
    RollCRED.TextoACola "tbrSoft", &HA9C8C9             '3 es texto gratis al publico
    If TOTAL_DISCOS < 2 Then
        RollCRED.TextoACola TR.Trad("NO HAY DISCOS" + vbCrLf + _
                            "Presione F9 para un " + vbCrLf + _
                            "asistente rápido o 'C' para " + vbCrLf + _
                            "configuración avanzada%98%" + _
                            "Arranco la fonola y no encontro musica%99%"), &HA9C8C9
    End If
    
    If K.sabseee("3pm") <= CGratuita Then
        RollCRED.TextoACola TR.Trad("PRESIONE F9" + vbCrLf + _
            "PARA AYUDA BASICA%98%Con F3 se abre el menu de " + _
            "licencia para comprar el programa%99%"), &HA9C8C9
        RollCRED.TextoACola TR.Trad("PRESIONE F3" + vbCrLf + _
            "PARA USAR EL SOFTWARE" + vbCrLf + _
            "SIN RESTRICCIONES%98%Con F3 se abre el menu de " + _
            "licencia para comprar el programa%99%"), &HA9C8C9
    End If
    
    RollCRED.INI
    
    RollSONG.SetInterval 70
    RollSONG.SetVarColor 5
    RollSONG.MaxlargoRenglon = 30
    RollSONG.TextoACola TR.Trad("Sin reproducción%99%"), &HA9C8C9 'cancion que se esta reproduciendo + rank
    RollSONG.TextoACola TR.Trad("no hay proximas canciones%99%"), &HA9C8C9 'la proxima cancion
    RollSONG.TextoACola TR.Trad("no hay proximas canciones%99%"), &HA9C8C9 'algun elemento del ranking
    If TOTAL_DISCOS < 2 Then
        RollSONG.TextoACola TR.Trad("NO HAY DISCOS" + vbCrLf + _
                            "Presione F9 para un " + vbCrLf + _
                            "asistente rápido o 'C' para " + vbCrLf + _
                            "configuración avanzada%98%" + _
                            "Arranco la fonola y no encontro musica%99%"), &HA9C8C9
    End If
    If K.sabseee("3pm") <= CGratuita Then
        RollSONG.TextoACola TR.Trad("PRESIONE F9" + vbCrLf + _
            "PARA AYUDA BASICA%98%Con F3 se abre el menu de " + _
            "licencia para comprar el programa%99%"), &HA9C8C9
        RollSONG.TextoACola TR.Trad("PRESIONE F3" + vbCrLf + _
            "PARA USAR EL SOFTWARE" + vbCrLf + _
            "SIN RESTRICCIONES%98%Con F3 se abre el menu de " + _
            "licencia para comprar el programa%99%"), &HA9C8C9
    End If
    RollSONG.INI
    
    tERR.Anotar "acep", K.sabseee("3pm")
    If K.sabseee("3pm") <= aSinCargar Then
        RollCRED.ReplaceIndex 3, TR.Trad("Este espacio sera suyo" + vbCrLf + _
                                 "cuando adquiera la" + vbCrLf + _
                                 "version full de 3PM" + _
                                 "%98%Espacio publicitario en texto no " + _
                                 "disponible por que esta en versión sin " + _
                                 "licencia aún%99%")
    Else
        RollCRED.ReplaceIndex 3, textoUsuario
    End If
    
    'cargar la cantidad de tapas que corresponda
    'SE CARGAN EN ini YA ES configurable
    'TapasMostradasH = 4: TapasMostradasV = 3
    '-----------------
    If K.sabseee("3pm") = Supsabseee Then
        If fso.FileExists(GPF("tslpri112")) Then
            tERR.Anotar "aceq"
            Set TE = fso.OpenTextFile(GPF("tslpri112"), ForReading, False)
                Dim NewT As String
                NewT = TE.ReadAll
            TE.Close
            tERR.Anotar "aceq", Len(NewT)
            RollCRED.ReplaceIndex 2, NewT
        Else
            tERR.Anotar "acer"
            RollCRED.ReplaceIndex 2, TR.Trad("Software desarrollado" + vbCrLf + _
                                     "por %99%") + "tbrSoft" + vbCrLf + _
                                     "www.tbrsoft.com" + vbCrLf + _
                                     "info@tbrsoft.com" + vbCrLf + _
                                     "tbrsoft@cpcipc.org."
        End If
    Else
        tERR.Anotar "aces"
        RollCRED.ReplaceIndex 2, TR.Trad("Software desarrollado" + vbCrLf + _
                                     "por %99%") + "tbrSoft" + vbCrLf + _
                                     "www.tbrsoft.com" + vbCrLf + _
                                     "info@tbrsoft.com" + vbCrLf + _
                                     "tbrsoft@cpcipc.org."
    End If
    '-----------------
End Sub

Public Sub SelDisco(nDisco As Long)
    
    On Error GoTo MiErr
    
    'version 7 con fondo cheto
    imageFONDO(nDisco).Visible = False
    TapaCD(nDisco).Visible = False
    imageFONDO(nDisco).Picture = imgSELBUP.Picture
    'lblDisco(nDisco).ForeColor = vbWhite
    tERR.Anotar "acfp", nDisco, nDiscoSEL, nDiscoGral
    
    nDiscoSEL = nDisco
    TapaCD(nDisco).Visible = True
    imageFONDO(nDisco).Visible = True
    
    Dim AAA As Long
    
    Dim FolRit As String
    Dim FolSel As String
    LineRitmo.Visible = False
    Dim LeftRitmoSel As Long
    For AAA = 0 To UBound(PartOrigenes)
        'ver que ritmo esta
        FolSel = UCase(fso.GetBaseName(fso.GetParentFolderName(txtInLista(MATRIZ_DISCOS(nDiscoGral), 0, ","))))
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
    
    'calcular el largo total ya que no son fuentes de tamaño igual por letra
    Dim IniC333 As Long
    IniC333 = 0
    For AAA = 65 To 90
        'lLETRAS(AAA - 65).Caption = Chr(AAA)
        IniC333 = IniC333 + lLETRAS(AAA - 65).Width + 60
    Next AAA
            
    IniC333 = (picFondo2.Width / 2) - (IniC333 / 2)
    
    For AAA = 65 To 90
        
        If AAA > 65 Then
            lLETRAS(AAA - 65).Left = lLETRAS(AAA - 66).Left + lLETRAS(AAA - 66).Width + 60
        Else 'es el primero ponerlo debajo del ritmo
            lLETRAS(0).Left = IniC333
        End If
        
        lLETRAS2(AAA - 65).Left = lLETRAS(AAA - 65).Left + 15
        lLETRAS2(AAA - 65).Top = lLETRAS(AAA - 65).Top + 15
        
        If UCase(Left(lblDISCO(nDisco), 1)) = UCase(lLETRAS(AAA - 65).Caption) Then
            lLETRAS(AAA - 65).ForeColor = vbRed
            LineLETRA.X1 = lLETRAS(AAA - 65).Left
            LineLETRA.X2 = lLETRAS(AAA - 65).Left + lLETRAS(AAA - 65).Width
            LineLETRA.Y1 = lLETRAS(AAA - 65).Top + 90
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
    If EsVideo Then
        OrdenarListaModoVideo
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdv"
    Resume Next
    
End Sub

Public Sub UnSelDisco(nDisco As Long)
    On Error GoTo MiErr
    
    'al iniciar el sistema puede hacer esto que se muestre
    If nDisco = 99999 Then Exit Sub
    
    tERR.Anotar "acfs", nDisco, nDiscoSEL, nDiscoGral, LastDiscoSel
    
    'imageFONDO(nDisco).Visible = False
    'TapaCD(nDisco).Visible = False
    
    imageFONDO(nDisco).Picture = imgUNSELBUP.Picture
    
    'lblDisco(nDisco).ForeColor = vbBlack
    'TapaCD(nDisco).Visible = True
    'imageFONDO(nDisco).Visible = True
    tERR.Anotar "acft", LastDiscoSel, EsVideo
    L(LastDiscoSel).ForeColor = vbBlack
    L(LastDiscoSel).BackColor = vbWhite
    If EsVideo Then
        OrdenarListaModoVideo
    End If
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
    Dim C As Integer
    C = 1
    
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
            
            If ArchTapa = "_RANK_" Then
                TapaCD(NDR).Picture = LOP.GetPicture("1", "2") 'tapa rank predeterminada o comun predeterminada
            Else
                tERR.Anotar "acge", ArchTapa
                If Right(ArchTapa, 1) <> "\" Then ArchTapa = ArchTapa + "\"
                ArchTapa = ArchTapa + "tapa.jpg"
                
                TapaCD(NDR).Picture = LOP.GetPicture(ArchTapa, "2")
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
    Else 'no elige directo
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
                    'a veces puedo estar pasando a la ultima página (para adelante)
                    'If TapaCD(DiscosEnPagina - 1).Visible = False Then
                    UnSelDisco DiscosEnPagina - 1 'gggggg
                Else
                    'si es el primer inicio al desseleccionar
                    'se muestra la tapa despintada y se oculta la elegida
                    If nDiscoSEL <> 99999 Then
                        UnSelDisco ((TapasMostradasH * TapasMostradasV) - 1)
                    End If
                End If
            Else 'no es 46
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
            
        Else 'no elige el primero
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

    'TF.Detener

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
    nDiscoSEL = nDiscoGral - (PagNum * (TapasMostradasH * TapasMostradasV))
    tERR.Anotar "acgy", PagNum, nDiscoSEL
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

'modo de seleccion de discos ORIGINAL
Private Sub SuperSel2(ByVal Index As Integer)
    
    On Local Error GoTo ErrSSel
    
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
    imgDiscoSEL.Width = (picFondoDisco.Width / 5)
    imgDiscoSEL.Height = (picFondoDisco.Height / 4)
    imgDiscoSEL.Top = cmdTouchArriba.Top + cmdTouchArriba.Height + 120 + btBUYDisco.Height + 60 + btBuyCancion.Height + 120 'picFondoDisco.Height / 2 - imgDiscoSEL.Height / 2
    imgDiscoSEL.Left = 500 'picFondoDisco.Width / 4 - imgDiscoSEL.Width / 2
    imgDiscoSEL.Visible = True
    
    lblDiscoSEL.Visible = False
    lblDiscoSEL2.Visible = False
    
    lblDiscoSEL.Caption = lblDISCO(Index).Caption
    lblDiscoSEL.Font.Size = lblDISCO(Index).Font.Size
    lblDiscoSEL.Top = imgDiscoSEL.Top + imgDiscoSEL.Height
    lblDiscoSEL.Left = imgDiscoSEL.Left
    lblDiscoSEL.Width = imgDiscoSEL.Width - 200
    lblDiscoSEL.Height = 500
    
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
    imgFondoDiscoSel.Width = (picFondoDisco.Width / 5) + 200
    imgFondoDiscoSel.Height = imgDiscoSEL.Height + lblDiscoSEL.Height + 200
    imgFondoDiscoSel.Top = imgDiscoSEL.Top - 100
    imgFondoDiscoSel.Left = imgDiscoSEL.Left - 200
    imgFondoDiscoSel.Visible = True
    
    imgDiscoSEL.ZOrder
    imgFondoDiscoSel.ZOrder
    lblDiscoSEL2.ZOrder
    lblDiscoSEL.ZOrder
    
    imgListaSong.Visible = False
    imgListaSong.Stretch = True
    
    IMF = ExtraData.getDef.getImagePath("MarcoFondodelosdiscos")
    imgListaSong.Picture = LoadPicture(IMF)
    
    Dim IND As Long
    IND = ExtraData.getDef.GetIndexImage("MarcoChicoIndicadores")
    Dim MargDer As Long, MargIzq As Long, MargSup As Long, MargInf As Long
    MargSup = imgListaSong.Height * ExtraData.getDef.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = imgListaSong.Height * ExtraData.getDef.GetFinalMargenInferiorTra(IND) / 100
    MargDer = imgListaSong.Width * ExtraData.getDef.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = imgListaSong.Width * ExtraData.getDef.GetFinalMargenIzquierdoTra(IND) / 100
    
    
    imgListaSong.Top = 150
    If MostrarTouch Then
        imgListaSong.Height = (picFondoDisco.Height) - imgSELEC.Height - 150
    Else
        imgListaSong.Height = (picFondoDisco.Height) - 150
    End If
    
    imgListaSong.Left = imgFondoDiscoSel.Left + imgFondoDiscoSel.Width + 200 ' picFondoDisco.Width / 2
    imgListaSong.Width = picFondoDisco.Width - imgListaSong.Left - 300
    
    'ya esta agrandado
    lblDATA.Font.Size = lblDiscoSEL2.Font.Size
    lblDATA.Width = imgFondoDiscoSel.Width - 200
    lblDATA.Height = picFondoDisco.Height - (imgFondoDiscoSel.Top + imgFondoDiscoSel.Height)
    lblDATA.Left = imgFondoDiscoSel.Left + 100
    lblDATA.Top = imgFondoDiscoSel.Top + imgFondoDiscoSel.Height + 100
    
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
    Dim PerfilEncontrado As Long
    If UbicDiscoActual = "_RANK_" Then
        MATRIZ_TEMAS = ObtenerRankComoMM(30) 'joia tengo un ranking decente!!!
        PerfilActual = -1 'perfil de ranking!!
    Else
        'OM- entro a un disco y quiero mostrar la lista de lo que hay
        If VentaExtras Then
            PerfilEncontrado = 1 'valor para que entre buscando
            MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual, True, PerfilEncontrado)
        Else
            'parta que ni busque perfiles, es un disco comun de 3PM
            MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual, True)
            PerfilEncontrado = 1 'es el de 3PM base
        End If
                
        'de la forma!
        'D:\musica\Cuartetazo\Alma Fuerte-En vivo Obras 2001\02 - Almafuerte.mp3#02 - Almafuerte.mp3
        PerfilActual = PerfilEncontrado
    End If
    tERR.Anotar "caah2", UBound(MATRIZ_TEMAS)
    'usar esto y no una variable para saber de discos vacios
    If UBound(MATRIZ_TEMAS) = 0 Then
        lblCanciones(0).Caption = TR.Trad("NO HAY CANCIONES EN ESTE DISCO!%99%")
        tERR.AppendLog "No hay temas en el disco: " + UbicDiscoActual + ".acpu"
        EstoyEnDisco = 1 'PARA QUE PUEDA SALIR!!!!
        Exit Sub
    End If
    
    'ordenar la lista de canciones
    Dim DataTXT As String
    If UbicDiscoActual = "_RANK_" Then
        DataTXT = TR.Trad("Estos son los mas escuchados!%99%")
    Else
        Dim ArchDaTa As String
        ArchDaTa = UbicDiscoActual + "data.txt"
        If fso.FileExists(ArchDaTa) Then
            Dim A As TextStream
            Set A = fso.OpenTextFile(ArchDaTa, ForReading, False)
                If A.AtEndOfStream = False Then
                    DataTXT = A.ReadAll
                Else
                    DataTXT = TR.Trad("No hay datos adicionales de este disco%99%")
                End If
            A.Close
        Else
            DataTXT = TR.Trad("No hay datos adicionales de este disco%99%")
        End If
    End If
    
    Select Case PerfilActual
        Case 1: DataTXT = DataTXT + vbCrLf + "3PM BASE"
        Case 2: DataTXT = DataTXT + vbCrLf + "3PM RINGTONES"
        Case 3: DataTXT = DataTXT + vbCrLf + "3PM WALLPAPERS"
        Case 4: DataTXT = DataTXT + vbCrLf + "3PM JAVA"
        Case 5: DataTXT = DataTXT + vbCrLf + "3PM ISOs"
        Case 6: DataTXT = DataTXT + vbCrLf + "3PM Videos para mobil"
    End Select
    
    'estos perfiles incluye una imagen por cada objeto de la lista
    If PerfilActual >= 3 And PerfilActual <= 6 Then
    
        'MODO 2 abajo a la derecha dentro de la misma lista de discos ...
        'se mezcla la imagen con los elementos de abajo de la lista
'        picExtraObjeto.Left = imgListaSong.Left + ((2 * imgListaSong.Width) / 3)
'        picExtraObjeto.Width = imgListaSong.Width / 3
'
'        picExtraObjeto.Height = imgListaSong.Height / 3
'        picExtraObjeto.Top = imgListaSong.Top + ((2 * imgListaSong.Height) / 3)
        
        'MODO 1
        picExtraObjeto.Left = imgListaSong.Left
        picExtraObjeto.Width = imgListaSong.Width
        picExtraObjeto.Height = imgListaSong.Height / 2
        picExtraObjeto.Top = picFondoDisco.Height - picExtraObjeto.Height - 45
        picExtraObjeto.Picture = imgListaSong.Picture 'el marco de abajo si lo hubiera es el mismo que arriba
        'achico la lista de imagenes por que esta es la referencia por los renglones _
            que se muestran y dependen tambien el movimiento de los objetos de la lista
        imgListaSong.Height = imgListaSong.Height - picExtraObjeto.Height
    
        picExtraObjeto.Visible = True
    Else
        picExtraObjeto.Visible = False
    End If
    
    imgListaSong.Visible = True
    
    'nombres de los botones de comprar
    Select Case PerfilActual
        Case 1
            btBUYDisco.Caption = "Comprar Disco"
            btBuyCancion.Caption = "Comprar canción"
        Case 2
            btBUYDisco.Caption = "Comprar todos"
            btBuyCancion.Caption = "Comprar Ringtone"
        Case 3
            btBUYDisco.Caption = "Comprar todos"
            btBuyCancion.Caption = "Comprar Wallpaper"
        Case 4
            btBUYDisco.Caption = "Comprar todos"
            btBuyCancion.Caption = "Comprar Juego"
    End Select
    lblDATA.Caption = DataTXT:    lblDATA2.Caption = lblDATA.Caption
    lblDATA.Visible = True:       lblDATA2.Visible = True
    
    Dim C As Integer, nombreTemas As String
    Dim pathTema As String
    C = 1
    Dim AltoRenglon As Long
    AltoRenglon = lblCanciones(0).Height + 30
    tERR.Anotar "caai", AltoRenglon
    Dim EXT As String

    'establecer los limites donde van los elemntos para leer despues
    'limite superior = imgListaSong.Top + imgListaSong.Height - AltoRenglon - MargInf
    'limite inferior = MargSup +  AltoRenglon
    imgListaSong.Tag = "LS:" + _
                       CStr(imgListaSong.Top + imgListaSong.Height - AltoRenglon - MargInf) + _
                       "|LI:" + _
                       CStr(imgListaSong.Top + MargSup + 90) + _
                       "|MD:" + _
                       CStr(MargDer) + _
                       "|MI:" + _
                       CStr(MargIzq)
    
    Do While C <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(C), 0, "#")
        nombreTemas = txtInLista(MATRIZ_TEMAS(C), 1, "#")
        EXT = LCase(txtInLista(pathTema, 1, "."))
        
        'quitar el molesto .mp3 o lo que fuera
        Select Case LCase(EXT)
            Case "mp3"
                EXT = "" 'se sobreentiende que todo es mp3" (mp3-Musica)"
'            Case "mp4"
'                EXT = " (mp4-Musica)"
            Case "wma"
                EXT = TR.Trad(" (wma-Musica)%99%")
            Case "mpeg", "mpg", "avi", "wmv"
                TR.SetVars LCase(EXT)
                EXT = TR.Trad(" (%01%-Video)%98%La variable 1" + _
                    "es MPG o AVI o WMV, es el formato de video " + _
                    "de un archivo%99%")
            Case "vob"
                EXT = TR.Trad(" (DVD!)%99%")
            Case "dat"
                EXT = TR.Trad(" (VCD-Video)%99%")
            Case "mn0", "mn1"
                EXT = TR.Trad(" (KARAOKE)%99%")
            Case "jpg", "jpeg", "bmp", "gif"
                EXT = TR.Trad(" (Wallpaper)%99%")
            Case "jar"
                EXT = TR.Trad(" (Java)%99%")
            'mm91
            'formatos de imagenes de nero
            'NR3: cd de mp3s    /    'NRA: cd de audio    /  'NRB: cd-rom de arranque
            'NRC: nero usf/iso  /    'NRD: nero DVD       /  'NRE: cd extra
            'NRG: imagen        /    'NRH: cd-rom hibrido /  'NRI: cd-rom iso
            'NRM: cd mixto      /    'NRU: cd-rom udf     /  'NRV: cd supervideo
            'NRW: cd rom wma    /    'CDC: cd cover no tiene nada que ver con imagenes parece
            Case "iso"
                EXT = TR.Trad(" (Imagen ISO)%99%")
            Case "nrg"
                EXT = TR.Trad(" (Imagen NERO)%99%")
            Case "nr3"
                EXT = TR.Trad(" (Imagen NERO MP3)%99%")
            Case "nra"
                EXT = TR.Trad(" (Imagen NERO AUDIO)%99%")
            Case "nrb"
                EXT = TR.Trad(" (Imagen NERO INICIO)%99%")
            Case "nrc"
                EXT = TR.Trad(" (Imagen NERO UDF/ISO)%99%")
            Case "nrd"
                EXT = TR.Trad(" (Imagen NERO DVD)%99%")
            Case "nre"
                EXT = TR.Trad(" (Imagen NERO CD EXTRA)%99%")
            Case "nrh"
                EXT = TR.Trad(" (Imagen NERO CD HIBR)%99%")
            Case "nri"
                EXT = TR.Trad(" (Imagen NERO ISO)%99%")
            Case "nrm"
                EXT = TR.Trad(" (Imagen NERO CD MIXTO)%99%")
            Case "nru"
                EXT = TR.Trad(" (Imagen NERO CD UDF)%99%")
            Case "nrv"
                EXT = TR.Trad(" (Imagen NERO SUPERVIDEO)%99%")
            Case "nrw"
                EXT = TR.Trad(" (Imagen NERO WMA)%99%")
            Case "3gp"
                EXT = TR.Trad(" (Video para movil)%99%")
        End Select
        
        nombreTemas = fso.GetBaseName(nombreTemas) + EXT
        Load lblCanciones(C)
        Load lblCanciones2(C)
        
        lblCanciones(C).Caption = nombreTemas
        tERR.Anotar "caaj", C, nombreTemas
        lblCanciones(C).Tag = pathTema
        lblCanciones(C).Top = imgListaSong.Top + MargSup + 90 + ((C - 1) * AltoRenglon)
        lblCanciones2(C).Top = lblCanciones(C).Top + 15
        lblCanciones2(C).Left = lblCanciones(C).Left + 15
        'tiene autosize
        
        C = C + 1 'ver que el proximo entre
        
    Loop
    
    Dim TotalSong As Long
    TotalSong = C - 1
    'en adelante se usa como referencia el ubound asi que lo corto directamente asi!
    'ReDim Preserve MATRIZ_TEMAS(TotalSong)
    'no se corta mas porque se muestra todo
    tERR.Anotar "caaj20", CargarDuracionTemas
    
    If CargarDuracionTemas Then
        'ahora cargar las duaciones
        Dim NoCargoDuracion As Long
        NoCargoDuracion = 0
        C = 1
        Dim MP3tmp As New MP3Info
        Do While C <= UBound(MATRIZ_TEMAS)
            pathTema = lblCanciones(C).Tag
            'si es mp3 usar el rápido, si no usar el viejo
            'mm91 no se puede tener la duracion de otros tipos de archivos
            Dim est As String
            est = UCase(Right(pathTema, 3))
            Select Case est
                Case "MP3"
                    MP3tmp.FileName = pathTema
                    DuracionTema = MP3tmp.DurationSTR
                Case "mpg", "avi", "mpeg"
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
            End Select
            lblCanciones(C).Caption = lblCanciones(C).Caption + " (" + DuracionTema + ")"
            C = C + 1
        Loop
        Set MP3tmp = Nothing
    End If

    'revisar especificamente que no haya nada mas largo que lo que se puede
    C = 1
    Do While C <= UBound(MATRIZ_TEMAS)
    
        'si o si dejar un margen
        If lblCanciones(C).Width > (imgListaSong.Width * 0.9) Then
            Dim D As Long
            For D = 1 To 35 'con estas pasadas debe quedar ok
                'que nunca de error!!!!
                If Len(lblCanciones(C).Caption) > 10 Then
                    lblCanciones(C).Caption = _
                        Mid(lblCanciones(C).Caption, 1, Len(lblCanciones(C).Caption) - 10) + "..."
                Else
                    Exit For
                End If
                'ver si con eso alcanza
                If lblCanciones(C).Width < (imgListaSong.Width * 0.9) Then Exit For
            Next D
        End If
        
        C = C + 1
    Loop


    C = 1
    Do While C <= UBound(MATRIZ_TEMAS)
        lblCanciones(C).Left = imgListaSong.Left + (imgListaSong.Width / 2 - lblCanciones(C).Width / 2)
        lblCanciones2(C).Left = lblCanciones(C).Left + 15
        
        'ver que no se muestren mas canciones de las que entren
        'estos como se pasan deben ser invisibles
        If lblCanciones(C).Top > (imgListaSong.Top + imgListaSong.Height _
                - AltoRenglon - MargInf) Then
            
            lblCanciones(C).Visible = False
            lblCanciones2(C).Visible = False
            lblCanciones2(C).Tag = "OUT DOWN" 'fuera de la visulización (por debajo)!!
        
        Else
            lblCanciones(C).Visible = True
            lblCanciones2(C).Visible = True
            lblCanciones2(C).Tag = "IN" 'fuera de la visulización !!
        End If
        
        
        lblCanciones2(C).ZOrder 'lo necesito paar poder hacerle click
        lblCanciones(C).ZOrder
        C = C + 1
    Loop
    
    
    lblNOCREDIT.Left = imgListaSong.Left + imgListaSong.Width - lblNOCREDIT.Width
    lblNOCREDIT.Top = imgListaSong.Top + imgListaSong.Height '- lblNOCREDIT.Height - 120
    
    If MostrarTouch Then
        cmdTouchAbajo.Top = 120 'imgListaSong.Top + cmdTouchArriba.Height + 120
        cmdTouchAbajo.Left = (imgListaSong.Left / 2 - cmdTouchAbajo.Width) ' - 120
        
        cmdTouchArriba.Top = 120 'imgListaSong.Top + 120
        cmdTouchArriba.Left = (imgListaSong.Left / 2) '+ 120
        
        cmdTouchArriba.Visible = True
        cmdTouchAbajo.Visible = True
        
        btBUYDisco.Top = cmdTouchArriba.Top + cmdTouchArriba.Height
        btBuyCancion.Top = btBUYDisco.Top + btBUYDisco.Height + 60
        
        btBUYDisco.Left = 120
        btBuyCancion.Left = 120
                
        If VendoMusica Then
            btBuyCancion.Visible = True
            btBUYDisco.Visible = True
        End If
        
        imgSELEC.Left = imgListaSong.Left + 90 ' imgListaSong.Left + (imgListaSong.Width / 3 - imgSELEC.Width)
        'imgSELEC.Left = imgListaSong.Left + (imgListaSong.Width / 3 - imgSELEC.Width)
        imgSELEC.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        ImgSelecVIP.Left = imgSELEC.Left + imgSELEC.Width + 60
        ImgSelecVIP.Top = imgSELEC.Top
        
        imgSALIR.Left = ImgSelecVIP.Left + ImgSelecVIP.Width + 60
        imgSALIR.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        'lblNOCREDIT.Top = imgSELEC.Height + imgSELEC.Top + 60
        
        imgSELEC.Visible = True
        imgSALIR.Visible = True
        ImgSelecVIP.Visible = (CreditosXaVipMusica > 0) And (PerfilActual = 1)
    Else
        btBuyCancion.Visible = False
        btBUYDisco.Visible = False
    End If
    
    EstoyEnDisco = 1
    OkInState1 = 0
    selDiscoI 1
    
    Exit Sub
    
ErrSSel:
    tERR.AppendLog "SSEL444-2", tERR.ErrToTXT(Err)
    Resume Next
End Sub

'se elige un elemento de la lista de discos
'se usa tambien para moverse dentro de la lista, los cometarios internos describen el funcionamiento
Private Function selDiscoI(I As Integer) As Long

    On Local Error GoTo ErrSDI

    'elegir una cancion de la lista
    'como aqui pueden venir con el mouse (que aun no pone en cero el contador de los botones)
    SecSinTecla = 0
    
    If EstoyEnDisco = 1 And PerfilActual = 1 Then
        lblNOCREDIT.Caption = getStrMusicaVIP 'deja vacio si no esta activado el vip o pone lo que va
    Else
        lblNOCREDIT.Caption = ""
    End If
    
    tERR.Anotar "sdi", I
    Dim TMPi As Long 'para saber siempre que se eligio originalmente
    TMPi = I
    'elegir disco en el index
    'si solo quiero el que sigue pongo -1 o -2 para el anterior
    '-3 solo para saber cual esta elegido por ejemplo para reproducirlo
    'devuelve -99 si no hay nada mas para elegir
    Dim C As Long
    Dim sSel As Long
    sSel = -1 'bandera de que nada esta elegido
    If I < 0 Then
        'necesito saber cual esta elegido
        For C = 1 To UBound(MATRIZ_TEMAS)
            If lblCanciones(C).BackStyle = 1 Then
                sSel = C
                Exit For
            End If
        Next C
    Else 'ya sabe lo que quiere
        sSel = I
    End If
    tERR.Anotar "sdi2", I, sSel
    'el que sigue
    If I = -1 Then sSel = sSel + 1
    'el anterior
    If I = -2 Then sSel = sSel - 1
    'ver que no se pase
    tERR.Anotar "sdi3", I, sSel, UBound(MATRIZ_TEMAS)
    'el limite para ambos casos estaba en 1 y funiocnaba ok
    'pero en disco de una sola cancion anda ok on el cero que parece que es el que va
    If sSel < 1 Then sSel = UBound(MATRIZ_TEMAS)
    If sSel > UBound(MATRIZ_TEMAS) Then sSel = 1
    tERR.Anotar "sdi4", I, sSel, lblDiscoSEL.Caption, lblCanciones.UBound
    I = sSel
    
    'ver si el que voy a elegir se puede elegir
    Dim CO As Long
    CO = 0
    Do
        If lblCanciones(I).Tag = "" Then 'lo pongo asi cuando una cancion se elije
            tERR.Anotar "sdi4B", I, sSel, lblCanciones.UBound
            If TMPi = -1 Then sSel = sSel + 1 'si iva para adelante sigo para adelante
            If TMPi = -2 Then sSel = sSel - 1 'si iva para atras sigo para atras
            
            If sSel < 1 Then sSel = UBound(MATRIZ_TEMAS)
            If sSel > UBound(MATRIZ_TEMAS) Then sSel = 1
            
        Else
            tERR.Anotar "sdi4B", I, sSel, lblCanciones.UBound
            I = sSel
            Exit Do 'ya encontre!
        End If
        I = sSel
        CO = CO + 1
        'si dio toda la vuelta me voy!
        If CO >= lblCanciones.UBound Then
            selDiscoI = -99
            Exit Function
        End If
    Loop
    
    tERR.Anotar "sdi5", I, sSel, UBound(MATRIZ_TEMAS)
    For C = 1 To UBound(MATRIZ_TEMAS) 'poner los colores que corresponde marcando el elegido
        lblCanciones(C).BackColor = vbBlack
        lblCanciones2(C).BackColor = lblCanciones(C).BackColor
        If C = I Then
            lblCanciones(C).BackStyle = 1
            lblCanciones2(C).BackStyle = 1
            lblCanciones(C).BorderStyle = 1
            'si es un wallpaper hacer una previsualizacion y si es una aplicacion java ver si tiene sshot ..
            'la funcion sola define si es un objeto que tenga preview posible
            ShowPreview lblCanciones(C).Tag
            
            'revisar además que este visible el elegido si la lista fuera mas larga de lo que corresponde
            If lblCanciones2(C).Tag = "OUT DOWN" Then
                'corro todos para arriba....
                'hatsa que quede!
                Do
                    MoverListaTemas -1
                    If lblCanciones2(C).Tag = "IN" Then Exit Do
                Loop
            End If
            
            If lblCanciones2(C).Tag = "OUT UP" Then
                'corro todos para abajo
                ' ... hasta que quede !!!!
                Do
                    MoverListaTemas 1
                    If lblCanciones2(C).Tag = "IN" Then Exit Do
                Loop
            End If
        Else
            lblCanciones(C).BackStyle = 0
            lblCanciones2(C).BackStyle = 0
            lblCanciones(C).BorderStyle = 0
        End If
    Next C
    
    selDiscoI = I
    
    lblXY1.Caption = CStr(I) + "/" + CStr(UBound(MATRIZ_TEMAS))
    
    lblXY1.Left = imgListaSong.Left + CLng(GetTag(imgListaSong.Tag, "MI")) 'imgListaSong.Left + imgListaSong.Width - (lblXY1.Width * 2)
    lblXY1.Top = CLng(GetTag(imgListaSong.Tag, "LI")) 'limite inferior
    
    lblXY1.Visible = True
    lblXY1.ZOrder
    
    Exit Function
    
ErrSDI:
    tERR.AppendLog "SDIerr:" + _
        CStr(I) + ":" + _
        CStr(lblCanciones.UBound) + ":" + _
        CStr(TMPi), _
        tERR.ErrToTXT(Err)
        
End Function

Private Sub ShowPreview(sF As String)
    'el parametro es el path completo del objeto elegido
    Dim imgToShow As String 'image a mostrar
    Dim txtInfoPV As String 'info adicional del extra
    Dim sizeKB As Long 'tamaño en KB de lo elegido
    Dim VerTamanoIMG As Boolean 'si es wallpaper me importa mostrar los pixles, si el screenshot de un jar no!!! se podria confundir con los pixles del juego!
    Select Case LCase(fso.GetExtensionName(sF))
        Case "jar"
            'buscar la imagen a mostrar
            imgToShow = Mid(sF, 1, Len(sF) - Len(fso.GetExtensionName(sF))) + "jpg"
            If fso.FileExists(imgToShow) = False Then
                imgToShow = GetTpPred 'trato de mostrar la imagen predeterminada
                If fso.FileExists(imgToShow) = False Then imgToShow = ""
            End If
            sizeKB = FileLen(sF) / 1024
            VerTamanoIMG = False
            
            'ver si ya se abrio este jar y ta tengo el manifiesto a mano
            Dim FilManif As String
            FilManif = Mid(sF, 1, Len(sF) - Len(fso.GetExtensionName(sF))) + "mf"
            If fso.FileExists(FilManif) Then
                txtInfoPV = "INFO JAR:" + vbCrLf + GetFullStringArch(FilManif)
            Else 'deja creado el archivo de manifiesto para entrrar mas rapido la proxima vez
                Dim JJ As New tbrJAR
                txtInfoPV = "INFO JAR:" + vbCrLf + JJ.getStringManifest(sF)
            End If
        
        'mm91
        'las imagenes iso pueden tener un archivo txt de informacion y una imagen JPG tambien
        Case "iso", "nrg", "nr3", "nra", "nrb", "nrc", "nrd", "nre", "nrh", "nri", "nrm", "nru", "nrv", "nrw"
            'buscar la imagen a mostrar
            imgToShow = Mid(sF, 1, Len(sF) - Len(fso.GetExtensionName(sF))) + "jpg"
            If fso.FileExists(imgToShow) = False Then
                imgToShow = GetTpPred 'trato de mostrar la imagen predeterminada
                If fso.FileExists(imgToShow) = False Then imgToShow = ""
            End If
            sizeKB = FileLen(sF) / 1024
            VerTamanoIMG = False
        Case "3gp" 'mm91 PODRIA abrir el video ????
            'seria mejor que tener que hacer una imagen
            'buscar la imagen a mostrar
            imgToShow = Mid(sF, 1, Len(sF) - Len(fso.GetExtensionName(sF))) + "jpg"
            If fso.FileExists(imgToShow) = False Then
                imgToShow = GetTpPred 'trato de mostrar la imagen predeterminada
                If fso.FileExists(imgToShow) = False Then imgToShow = ""
            End If
            sizeKB = FileLen(sF) / 1024
            VerTamanoIMG = False
        
        Case "jpg", "jpeg", "bmp", "gif"
            imgToShow = sF
            sizeKB = FileLen(sF) / 1024
            VerTamanoIMG = True
        Case Else 'si es un mp3 o video o kar o sea nada con preview salgo
            Exit Sub
    End Select
    
    txtInfoPV = txtInfoPV + vbCrLf + "Necesita " + CStr(sizeKB) + " KB"
        
    imgExtraObjeto.Visible = False
    lblInfoPreview.Visible = False
    imgExtraObjeto.Stretch = False
    imgExtraObjeto.Picture = LoadPicture(imgToShow)
    
    If VerTamanoIMG Then
        Dim AnchoPix As Long: AnchoPix = imgExtraObjeto.Width / 15
        Dim AltoPix As Long: AltoPix = imgExtraObjeto.Height / 15
        txtInfoPV = txtInfoPV + vbCrLf + _
            "Ocupa " + CStr(AnchoPix) + " x " + CStr(AltoPix) + " pixeles (ancho x alto)"
    End If
    
    'buscar texto
    'en JPG sera el tamaño en pixeles y el peso en MB
    'en java será el peso en MB + requisitos del soft (cldc)+ tamaño en pixeles recomedado
    Dim FilInfo As String
    FilInfo = Mid(sF, 1, Len(sF) - Len(fso.GetExtensionName(sF))) + "txt"
    If fso.FileExists(FilInfo) Then
        Dim A As TextStream
        Set A = fso.OpenTextFile(ArchDaTa, ForReading, False)
            txtInfoPV = txtInfoPV + vbCrLf + A.ReadAll
        A.Close
    Else
        'txtInfoPV = txtInfoPV + vbCrLf + "Sin información adicional"
    End If
    
    lblInfoPreview.Caption = txtInfoPV
    
    'ver si hay que achicarlo
    Dim CoefH As Single, CoefW As Single
    CoefH = imgExtraObjeto.Height / (picExtraObjeto.Height * 0.9)
    CoefW = imgExtraObjeto.Width / ((picExtraObjeto.Width / 2) * 0.9)
    
    If CoefW > 1 Or CoefH > 1 Then
        imgExtraObjeto.Stretch = True
        Dim newW As Single, NewH As Single
        If CoefH > CoefW Then
            newW = imgExtraObjeto.Width / CoefH
            NewH = imgExtraObjeto.Height / CoefH
        Else
            newW = imgExtraObjeto.Width / CoefW
            NewH = imgExtraObjeto.Height / CoefW
        End If
        
        imgExtraObjeto.Width = newW
        imgExtraObjeto.Height = NewH
        
    End If
    
    lblInfoPreview.Width = (picExtraObjeto.Width / 2) * 0.9
    lblInfoPreview.Height = picExtraObjeto.Height * 0.9
    lblInfoPreview.Top = picExtraObjeto.Top + (picExtraObjeto.Height / 2 - lblInfoPreview.Height / 2)
    lblInfoPreview.Left = picExtraObjeto.Left + (picExtraObjeto.Width / 4 - lblInfoPreview.Width / 2)
    
    imgExtraObjeto.Top = picExtraObjeto.Top + (picExtraObjeto.Height / 2 - imgExtraObjeto.Height / 2)
    'de la mitrad hacia la derecha
    imgExtraObjeto.Left = picExtraObjeto.Left + (picExtraObjeto.Width / 2) + _
        (picExtraObjeto.Width / 4 - imgExtraObjeto.Width / 2)
    
    imgExtraObjeto.Visible = True
    lblInfoPreview.Visible = True
    'picExtraObjeto.ZOrder
    'imgExtraObjeto.ZOrder
End Sub

Private Sub MoverListaTemas(Direccion As Long)
    'direccion puede ser positivo para abajo o negativo para arriba
    
    'un numero que uso
    
    'ver si hay como medir!
    If lblCanciones.UBound < 2 Then Exit Sub
    
    Dim DifReng As Long
    DifReng = lblCanciones(2).Top - lblCanciones(1).Top
    
    Dim C As Long
    'para que no se vea feo escondo todo
    For C = 1 To lblCanciones.UBound
        lblCanciones(C).Visible = False
        lblCanciones2(C).Visible = False
    Next C
    
    For C = 1 To lblCanciones.UBound
        'si direccion es negativo resta y sube, no hace falta un If DIRECCION>0
        lblCanciones(C).Top = lblCanciones(C).Top + (DifReng * Direccion)
        lblCanciones2(C).Top = lblCanciones(C).Top
        'poner el tag de si se esta viendo o no
        If lblCanciones(C).Top < CLng(GetTag(imgListaSong.Tag, "LI")) Then
            lblCanciones2(C).Tag = "OUT UP"
            lblCanciones(C).Visible = False
            lblCanciones2(C).Visible = False
            GoTo SSIG
        End If
        
        If lblCanciones(C).Top > CLng(GetTag(imgListaSong.Tag, "LS")) Then 'LS es limite superior
            lblCanciones2(C).Tag = "OUT DOWN"
            lblCanciones(C).Visible = False
            lblCanciones2(C).Visible = False
            GoTo SSIG
        End If
        
        'SI LLEGO HASTA ACA ESTA ADENTRO!
        lblCanciones2(C).Tag = "IN"
        lblCanciones(C).Visible = True
        lblCanciones2(C).Visible = True

SSIG:
    Next C
End Sub

Private Sub UnSuperSel()
    tERR.Anotar "sdi6"
    
    EstoyEnDisco = 1
    
    btOKPacha.Caption = "Ingresar a disco"
    Dim IMF As String
    IMF = ExtraData.getDef.getImagePath("touchizqnormal")
    t1.Picture = LoadPicture(IMF)
    
    IMF = ExtraData.getDef.getImagePath("touchderechanormal")
    t3.Picture = LoadPicture(IMF)
    
    lblNOCREDIT.Caption = ""
    imgSELEC.Visible = False
    imgSALIR.Visible = False
    ImgSelecVIP.Visible = False
    Dim M As Long
    On Local Error Resume Next
    For M = 0 To (TapasMostradasH * TapasMostradasV) - 1
        If TapaCD(M).Tag = "1" Then
            tERR.Anotar "sdi7", M
            TapaCD(M).Visible = True
            imageFONDO(M).Visible = True
            If MostrarRotulos Then
                lblDISCO(M).Visible = True
                lblDisco2(M).Visible = True
            End If
        End If
    Next M
    
    'imgDiscoSEL.Picture = imageFONDO(nDiscoGral).Picture
    imgDiscoSEL.Visible = False
    lblDiscoSEL.Visible = False
    lblDiscoSEL2.Visible = False
    imgFondoDiscoSel.Visible = False
    imgListaSong.Visible = False
    picExtraObjeto.Visible = False
    imgExtraObjeto.Visible = False
    lblInfoPreview.Visible = False
    lblDATA.Visible = False
    lblDATA2.Visible = False
    
    lblXY1.Visible = False
    
    'descargar todos los objetos cargados
    For M = 1 To 200 'no debo permitir que se cargue mas de 90
        'refuerzo estupido porque rene dice que se siguen viendo ???
        lblCanciones(M).Visible = False
        lblCanciones2(M).Visible = False
        Unload lblCanciones(M)
        Unload lblCanciones2(M)
    Next M
    
    'If MostrarTouch Then
        cmdTouchArriba.Visible = False
        cmdTouchAbajo.Visible = False
        btBuyCancion.Visible = False
        btBUYDisco.Visible = False
    'End If
    btSalir.Visible = False
    EstoyEnDisco = 0

End Sub

Private Sub ImgSelecVIP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokelegidovip")
    ImgSelecVIP.Picture = LoadPicture(IMF)
End Sub

Private Sub ImgSelecVIP_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokcomunvip")
    ImgSelecVIP.Picture = LoadPicture(IMF)
    EjecutarDeTouch True 'si o si es vip
End Sub

Private Sub imgSELEC_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokelegido")
    imgSELEC.Picture = LoadPicture(IMF)
    'XXXX aqui tambien contar el tiempo para temas VIP !!!!
    TimePressTeclaOK = Timer
End Sub

Private Sub imgSELEC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokcomun")
    imgSELEC.Picture = LoadPicture(IMF)
    
    TimePressTeclaOK = Timer - TimePressTeclaOK
    Dim isVip As Boolean 'ver si lo dejo mucho apretado es un VIP
    isVip = (TimePressTeclaOK > 1.5) And PrecNowVIP > 0
    TimePressTeclaOK = -1 'listo ya lo use
    
    EjecutarDeTouch isVip
End Sub

Private Sub imgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonsalirapretado")
    imgSALIR.Picture = LoadPicture(IMF)
End Sub

Private Sub imgSalir_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonsalirnormal")
    imgSALIR.Picture = LoadPicture(IMF)
    UnSuperSel
End Sub

Private Sub imgSELEC2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokelegido")
    imgSelec2.Picture = LoadPicture(IMF)
End Sub

Private Sub imgSELEC2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    IMF = ExtraData.getDef.getImagePath("botonokcomun")
    imgSelec2.Picture = LoadPicture(IMF)
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
    
    lCredPacha.Caption = lblCreditos.Caption
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

Private Sub lblNOCREDIT_Change()
    'CENTRARSE BIEN YA QUE TIENE AUTOSIZE
    ' Y SE USA DOS TEXTO "CREDITO INSUFICIENTE"   Y   "Credito VIP x $ x.xx"
    lblNOCREDIT.Left = imgListaSong.Left + imgListaSong.Width - lblNOCREDIT.Width
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
    Tapa = fso.GetParentFolderName(MP3.FileName(iAlias)) + "\tapa.jpg"
    
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
    
    'XXXX y si ya empezo antes ?
    'no ha jodido hasta ahora ....
    If iAlias = 2 Then
        If Salida2 Then
            frmVIDEO.picKAR_V.Picture = LoadPicture
            frmVIDEO.picKAR_V.Cls
            frmVIDEO.picKAR_V.Visible = False
        Else
            picKAR.Picture = LoadPicture
            picKAR.Cls
            picKAR.Visible = False
        End If
        EsVideo = False
        EsKar = False
        EstoyEnModoVideoMiniSelDisco = False
    End If
    
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
    'X
    
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
    Dim ES As String
    ES = EMPEZAR_SIGUIENTE(5)
    If ES <> 4 Then
        'sigue algo que no es video!
        VerSiTocaVMute
    End If
    
    'si no hay tema a continuacion y termino un video no se acomodaba
    'en empezar_siguiente ya esvideo se puso en false!
    If ES = 6 Then 'solo si no hay nada a continuacion
        UpdateVista 'se reacomoda a lo normal al terminar una cancion
    End If
Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acdz2"
    Resume Next
End Sub

Private Sub ByeTema(retEmpSig As Long)
    UpdateVista 'ya cuando se empieza a ir se acomoda
End Sub

Private Sub MP3_FaltaNextEvKAR(dMiliSec As Double)
    
    If Salida2 Then
        If (dMiliSec > 0) Then
            frmVIDEO.LF1_V = Format(dMiliSec, "00")
            frmVIDEO.LF2_V.Caption = frmVIDEO.LF1_V.Caption
        End If
        frmVIDEO.LF1_V.Visible = (dMiliSec > 0)
        frmVIDEO.LF2_V.Visible = frmVIDEO.LF1_V.Visible
    Else
        If (dMiliSec > 0) Then
            LF1 = Format(dMiliSec, "00")
            LF2.Caption = LF1.Caption
        End If
        LF1.Visible = (dMiliSec > 0)
        LF2.Visible = LF1.Visible
    End If
End Sub

Private Sub MP3_mmError(txtMasHist As String)
    tERR.AppendLog txtMasHist
End Sub

Private Sub ShowPaso(Inis33 As String, iAlias As Long, SP As Long)
    List1.List(iAlias) = Inis33 + " PLAY" + CStr(iAlias) + ":" + CStr(SP) + _
            ":" + CStr(TotalTema(iAlias)) + "(" + _
            CStr(MP3.HastaTema(iAlias)) + "):" + CStr(MP3.Volumen(iAlias)) + _
            " CUT:" + CStr(CORTAR_TEMA(iAlias)) + _
            " ToySal:" + CStr(YaEsoySaliendoGrat_Cortar(iAlias)) + " video/k:" + CStr(EsVideo) + "-" + CStr(EsKar)
End Sub

Private Sub MP3_Played(SecondsPlayed As Long, iAlias As Long, MS As Long)

    'cualquier cosa se corrige despues!
    EnableFF = True:    EnableNextMusic = True
    
    If iAlias = 2 Then
        If Salida2 Then
            frmVIDEO.lblTimeK_V = MP3.Falta(2)
            frmVIDEO.lblTimeK2_V = frmVIDEO.lblTimeK_V
        Else
            lblTimeK = MP3.Falta(2)
            lblTimeK2 = lblTimeK
        End If
        
        Exit Sub
    End If
    
    tERR.Anotar "acgv0", MS, iAlias, ThisFade, SegFade
    
    ShowPaso "==", iAlias, MS
    
    List1.List(4) = "IAA:" + CStr(IAA)
    List1.List(5) = "IAANext:" + CStr(IAANext)
    
    If iAlias = 3 Then Exit Sub
    
    On Error GoTo MiErr
    
    Dim NV As Long 'para nuevos voluemnes si se tienen  que cambiar
    
    'los primeros X segundos van en FadeIn sea el momento que sea
    If (SecondsPlayed - varSecPlay) <= ThisFade Then
        '**********************************************
        IenPlenaCancion(iAlias) = 1 'indica que esta empezando
        '**********************************************
        YaEsoySaliendoGrat_Cortar(iAlias) = False
        
        List1.List(6) = "ININ:" + CStr(iAlias)
        EnableFF = False:        EnableNextMusic = False
        
        ShowPaso "++", iAlias, MS
    
        tERR.Anotar "acgv2", CORTAR_TEMA(iAlias), VolumenIni, VolumenIni2
        Dim NewSec As Long
        NewSec = (MS - (varSecPlay * 1000))
        
        If CORTAR_TEMA(iAlias) Then
            NV = CLng(VolumenIni2 * (NewSec / 1000) * CSng(1 / ThisFade))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
        Else
            NV = CLng(VolumenIni * (NewSec / 1000) * CSng(1 / ThisFade))
            If NV > 100 Then NV = 100: If NV < 0 Then NV = 0
        End If
        tERR.Anotar "acgv12", NewSec, varSecPlay, NV, ThisFade
        
        MP3.Volumen(iAlias) = NV
        
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
    If ((SecondsPlayed - varSecPlay) > ThisFade) And (F > ThisFade) Then
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
        If ((SecondsPlayed - varSecPlay) > SegFade) And (ThisFade <> SegFade) Then
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
            
            Dim lRet As Long
            lRet = EMPEZAR_SIGUIENTE(1)
            
            ByeTema lRet
            
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
            
            'si es un tema gratuito corre el riesgo de cortarse de nuevo un poco mas abajo
            If MP3.EsGratis(iAlias) Then
                YaEsoySaliendoGrat_Cortar(iAlias) = True
            Else
                GoTo SIGUE55
            End If
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
        If K.sabseee("3pm") <= CGratuita Then
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

Private Sub picFondo_DblClick()
    'YaCerrar3PM False, True
End Sub

Private Sub picVideo_Resize(Index As Integer)
    picKAR.Width = picVideo(Index).Width
    picKAR.Height = picVideo(Index).Height
    picKAR.Top = picVideo(Index).Top
    picKAR.Left = picVideo(Index).Left
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

'Private Sub TF_PerdioFoco(hwndFoco As Long)
'    'TF.PonerFoco
'End Sub

Private Sub Timer1_Timer()
    On Error GoTo MiErr
    
    'controla el tiempo sin uso (sin ejecucion de temas)
    If MP3.IsPlaying(0) Or MP3.IsPlaying(1) Or MP3.IsPlaying(2) Then Exit Sub
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
        
        Z = Int(Rnd * TemasDisponibles)
        Z = Z + 1
        CC = 0
        tERR.Anotar "achb", Z
        If fso.FileExists(GPF("rd3_444")) = False Then
            fso.CreateTextFile GPF("rd3_444"), True
            'me voy al azar ya que no hay para elegirdel rank
            tERR.Anotar "achc.NORANK"
            GoTo MataReloj
        End If
        Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
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
            tERR.Anotar "ache", TT, CC, Z
            If CC = Z Then
                Dim TemaAzar As String
                TemaAzar = txtInLista(TT, 1, ",")
                'si tuve los discos cargados en una unidad o una ubicación distinta a la que aparece
                'en el ranking, me da un error por que el archivo no existe
                If fso.FileExists(TemaAzar) Then
                    tERR.Anotar "achg", TemaAzar
                    MP3.EsGratis(IAANext) = True
                    CORTAR_TEMA(IAANext) = True 'este tema se eligio al azar no va entero
                    SecSinUso = 0
                    TE.Close
                    EjecutarTema TemaAzar, False, "AZAR"
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

'reloj para salir de un disco o el protctor de pantalla
Private Sub Timer3_Timer()
    On Error GoTo MiErr
    
    '*****************************************************
    'ACA APROVECHO ESTE RELOJ PARA OTRAS COSAS
    Dim MT As Long, MU As Long
    my_MEM.GetMem MT, MU
    List1.List(20) = "MEM: " + CStr(MT) + "/" + CStr(MU)
    
    'si no tiene el foco ponerlo!!!
    If ForceFocus(Me.HWND) = False Then tERR.AppendSinHist "NOFOCO"
    'If TF.GetState <> 1 Then TF.PonerFoco
    '*****************************************************
    
    If Protector = 0 Then Exit Sub 'SE QUEDA PARA SALIR DE LOS DISCOS'Timer3.Interval = 0
    'para el reloj del protector. Lo ha inhabilitado
    'controla el tiempo sin uso (sin tocar teclas)
    SecSinTecla = SecSinTecla + (Timer3.Interval / 1000)
    'dragones: destino de fuego
    'no protector en video
    If EsVideo Then
        SecSinTecla = 0
        Exit Sub
    End If
    tERR.Anotar "achn", SecSinTecla, EsperaTecla
    ' a los x segundos sale del disco!
    If SecSinTecla > 15 And EsVideo = False Then
        If EstoyEnDisco = 1 Then 'corregido con el manu
            UnSuperSel
        End If
    End If
    
    If SecSinTecla > EsperaTecla And EsVideo = False Then
        If Protector = 3 Then 'empezar a moverse solo!
            Form_KeyDown TeclaDER, 0
        Else 'ir a los demas protectores
            frmProtect.Show 1
        End If
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
    If fso.FileExists(GPF("rd3_444")) = False Then
        tERR.Anotar "achp"
        fso.CreateTextFile GPF("rd3_444"), True
        TemasEnRank = 0
        Exit Function
    End If
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    
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
    Dim Cl As Long 'contador de L
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
        Cl = 0
        Do While Cl < TOTAL_DISCOS
            L(Cl).Top = L(Cl).Top - HayQueCorrerse
            Cl = Cl + 1
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
        Cl = 0
        Do While Cl < TOTAL_DISCOS
            L(Cl).Top = L(Cl).Top + HayQueCorrerse
            Cl = Cl + 1
        Loop
    End If
    
Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aceg"
    Resume Next
End Sub

Public Sub SelTema(n As Integer)
    T(n).BackStyle = 1
    T(n).ForeColor = vbWhite
    'T(n).BackColor = &H0&
    'T(n).ForeColor = &H80FFFF
End Sub

Public Sub UnSelTema(n As Integer)
    T(n).BackStyle = 0
    T(n).ForeColor = vbBlack
    'T(n).BackColor = &H80FFFF
    'T(n).ForeColor = &H0&
End Sub

Public Sub OrdenarListaTemaVideo()
    On Error GoTo MiErr
    'asegurarme que el disco elegido se ve en la lista
    tERR.Anotar "achw"
    Dim Cl As Long 'contador de L
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
        Cl = 0
        Do While Cl <= UBound(MATRIZ_TEMAS)
            T(Cl).Top = T(Cl).Top - HayQueCorrerse
            Cl = Cl + 1
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
        Cl = 0
        Do While Cl <= UBound(MATRIZ_TEMAS)
            T(Cl).Top = T(Cl).Top + HayQueCorrerse
            Cl = Cl + 1
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
    
    
    'picFondo2.Height = 765
    'no usarlo porque el tamaño de la fuente de define automaticamente y no queda bien.
    'seria verse o no nada mas
    If QuitaBarraSup Then picFondo2.Height = 0
    If QuitaBarraInf Then
        tbrPassImg1.Visible = False
        picFondo.Height = picFondo.Height / 2
        lblCreditos.Top = picFondo.Height / 2 - lblCreditos.Height / 2
    Else
        lblCreditos.Top = 30
    End If
    
    'barra inferior
    'picFondo.Height = 1500 'ejemplo de achicar para ver mas de discos
    picFondo.Top = Screen.Height - picFondo.Height
    
    picFondoPacha.Top = Screen.Height - picFondoPacha.Height
    picFondoPacha.Width = Me.Width
    picFondoPacha.Left = 0
    
    'barra superior
    picFondo2.Top = 0
    lRITMO(0).Top = 0
        
    'contenedor de los discos
    frDiscos.Top = picFondo2.Height
    frDiscos.Height = picFondo.Top - picFondo2.Height
    If PachaMode = 11000 Then
        picFondo.Visible = False
        frDiscos.Height = picFondoPacha.Top - picFondo2.Height
        'mostrar los botones horizontales
        Dim IMF As String
        
        lCredPacha.Left = 60
        lCredPacha.Top = picFondoPacha.Height / 2 - lCredPacha.Height / 2
        t1.BorderStyle = 0
        btOKPacha.Caption = "Ingresar a disco"
        t3.BorderStyle = 0
        
'        IMF = ExtraData.getDef.GetImagePath("botonokcomun")
'        t2.Picture = LoadPicture(IMF)
        
        IMF = ExtraData.getDef.getImagePath("touchizqnormal")
        t1.Picture = LoadPicture(IMF)
        
        IMF = ExtraData.getDef.getImagePath("touchderechanormal")
        t3.Picture = LoadPicture(IMF)
        
        btOKPacha.Width = 1500
        btOKPacha.Height = picFondoPacha.Height + 60  'los 60 se unde por abajo
        
        btOKPacha.Left = picFondoPacha.Width / 2 - btOKPacha.Width / 2
        btOKPacha.Top = picFondoPacha.Height - btOKPacha.Height + 60
        
        t1.Left = btOKPacha.Left - t1.Width - 320
        t1.Top = picFondoPacha.Height / 2 - t1.Height / 2
        
        t3.Left = btOKPacha.Left + btOKPacha.Width + 320
        t3.Top = picFondoPacha.Height / 2 - t3.Height / 2
        
        t1.Visible = True
        btOKPacha.Visible = True
        t3.Visible = True
        If PachaMode = 11000 Then
            picFondoPacha.Visible = True
            picFondoPacha.ZOrder
        End If
    End If
End Sub

Public Sub UpdateVista()
    'frDiscos.Visible = False
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
    'tomo como referencia el frDiscos que me define casi todo
    
    'ver si le saco el pedazo de la lista
            
    If HabilitarVUMetro Then
        frDiscos.Width = Me.Width - (AnchoBarra * 2) - 50
        frDiscos.Left = AnchoBarra + 30
    Else
        frDiscos.Left = 0
        frDiscos.Width = Me.Width - 50
    End If
            
    picFondo2.Left = frDiscos.Left
    picFondo2.Width = frDiscos.Width 'screen.Width
    
    If EstoyEnModoVideoMiniSelDisco Then
        frDiscos.Width = frDiscos.Width - frModoVideo.Width
    End If
    
    IMF = ExtraData.getDef.getImagePath("MarcoFondodelosdiscos")
    frDiscos.PaintPicture LoadPicture(IMF), 0, 0, frDiscos.Width, frDiscos.Height
    
    'listo frDiscos, ahora si lo tomo como referencia
    '**********************************************************
        
    frModoVideo.Visible = EstoyEnModoVideoMiniSelDisco
    lblModoVideo.Visible = EstoyEnModoVideoMiniSelDisco
    
    'si entre a la lista de discos hago lugar para eso
    If EstoyEnModoVideoMiniSelDisco Then
        AcomodarModoTexto 1
    End If
    
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
    
    IMF = ExtraData.getDef.getImagePath("MarcoChicoIndicadores")
    tERR.Anotar "aceu2", IMF
    'picFondo.Picture = LoadPicture(imF)
    picFondo.PaintPicture LoadPicture(IMF), 0, 0, picFondo.Width, picFondo.Height

    'dentro del picfondo hay que reacomodar
    lblCreditos.Left = picFondo.Width / 2 - lblCreditos.Width / 2
    
    'que se reacomode la sombra de los creditos, es necesario cuando se cambia el tamaño de picFondo!!!
    lblCreditos_Change
    
    'usar los datos del skin para saber como colocarlos
    Dim MargDer As Long, MargIzq As Long, MargSup As Long, MargInf As Long
    Dim IND As Long
    IND = ExtraData.getDef.GetIndexImage("MarcoChicoIndicadores")
    tERR.Anotar "aceu3", IND
    MargSup = picFondo.Height * ExtraData.getDef.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = picFondo.Height * ExtraData.getDef.GetFinalMargenInferiorTra(IND) / 100
    MargDer = picFondo.Width * ExtraData.getDef.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = picFondo.Width * ExtraData.getDef.GetFinalMargenIzquierdoTra(IND) / 100
    
    If MargSup = 0 Then MargSup = 60
    If MargInf = 0 Then MargInf = 60
    
    tbrPassImg1.Left = picFondo.Width / 2 - tbrPassImg1.Width / 2
    RollCRED.Width = tbrPassImg1.Left - MargDer - 90 'tbrPassImg1.Left * 0.75
    RollSONG.Width = (picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) - MargIzq - 90 '(picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) * 0.75
    RollSONG.Height = picFondo.Height - MargSup - MargInf   'picFondo.Height * 0.7
    RollCRED.Height = picFondo.Height - MargSup - MargInf   ' picFondo.Height * 0.7
    RollCRED.Left = MargDer '(tbrPassImg1.Left / 2) - (RollCRED.Width / 2)
    RollSONG.Left = (tbrPassImg1.Left + tbrPassImg1.Width) '(tbrPassImg1.Left + tbrPassImg1.Width) + _
                    ((picFondo.Width - (tbrPassImg1.Left + tbrPassImg1.Width)) / 2) - _
                    (RollSONG.Width / 2)
    RollSONG.Top = MargSup  'picFondo.Height / 2 - RollSONG.Height / 2
    RollCRED.Top = MargSup  'picFondo.Height / 2 - RollCRED.Height / 2
    
    'al momento de definir el skin se define un porcentaje que ocuapa cada uno de los 4 margenes
        
    IND = ExtraData.getDef.GetIndexImage("MarcoFondoDeLosDiscos")
    tERR.Anotar "aceu4", IMF
    MargSup = frDiscos.Height * ExtraData.getDef.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = frDiscos.Height * ExtraData.getDef.GetFinalMargenInferiorTra(IND) / 100
    MargDer = frDiscos.Width * ExtraData.getDef.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = frDiscos.Width * ExtraData.getDef.GetFinalMargenIzquierdoTra(IND) / 100
    
    picFondoDisco.Top = MargSup
    picFondoDisco.Left = MargDer
    picFondoDisco.Height = frDiscos.Height - MargSup - MargInf
    picFondoDisco.Width = frDiscos.Width - MargDer - MargIzq
    
    tERR.Anotar "aceu5", Salida2
    
    picFondo.Visible = (PachaMode <> 11000) 'imagen de fondo de los indicadores en modo simple
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
        picKAR.Visible = False
        frmVIDEO.picKAR_V.Visible = False
    End If
    
    'frDiscos.Visible = True
    
    
End Sub

Private Sub tUltra_Change()
    'me esta hablando el control
    
    If tUltra.tExt = "" Then Exit Sub
    tERR.Anotar tUltra.tExt
    Dim SP() As String
    SP = Split(tUltra.tExt, "|")
    
    Select Case CLng(SP(0)) 'lista de solicitudes
        Case 10000 'setear un hwnd para recibnir los errores
            tERR.AppendSinHist "go-tius:" + CStr(SP(1))
            tERR.Set_Hwn CLng(SP(1))
        Case 11000
            YaCerrar3PM
        Case 12000
            frDiscos.Visible = Not frDiscos.Visible
    End Select
    
    'para que cualquier cosa sea un cambio real
    tUltra.tExt = ""
End Sub

Private Sub TUsb_Change()
    If TUsb.tExt = "" Then Exit Sub
    
    Dim SP() As String
    SP = Split(TUsb.tExt, "|")
    
    Select Case SP(0)
        Case "0" 'entro drive
            'ver si esta el archivo de validación de este equipo
            'ademas grabar el registro de todo en el PC para saber que se grabo
            Ucdate SP(1)
            
        Case "1" 'sale drive
            
    End Select
    
    TUsb.tExt = ""
End Sub

Private Sub txtS3_Change()
    
    If txtS3.tExt = "" Then Exit Sub
    
    Dim P As String
    P = txtS3.tExt
    tERR.Anotar "acei88:" + P
    
    Dim SP() As String
    SP = Split(P, ":")
    
    If InStr(SP(0), "sD") Then
        tERR.Anotar "acei89:" + SP(1)
        Select Case CLng(SP(1))
            Case TeclaIZQx2: SendKeys Chr(TeclaIZQ)
            Case TeclaDERx2: SendKeys Chr(TeclaDER)
            Case TeclaPagAdx2: SendKeys Chr(TeclaPagAd)
            Case TeclaPagAtx2: SendKeys Chr(TeclaPagAt)
            Case TeclaOKx2: SendKeys Chr(TeclaOK)
            Case TeclaCancionVIPx2: SendKeys Chr(TeclaCancionVIP)
            Case TeclaCarritox2: SendKeys Chr(TeclaCarrito)
            Case TeclaESCx2: SendKeys Chr(TeclaESC)
            Case TeclaConfigx2: SendKeys Chr(TeclaConfig)
            Case TeclaCerrarSistemax2: SendKeys Chr(TeclaCerrarSistema)
            Case TeclaShowContadorx2: SendKeys Chr(TeclaShowContador)
            Case TeclaPutCeroContadorx2: SendKeys Chr(TeclaPutCeroContador)
            Case TeclaFFx2: SendKeys Chr(TeclaFF)
            Case TeclaBajaVolumenx2: SendKeys Chr(TeclaBajaVolumen)
            Case TeclaSubeVolumenx2: SendKeys Chr(TeclaSubeVolumen)
            Case TeclaNextMusicx2: SendKeys Chr(TeclaNextMusic)
            Case TeclaNewFichax2
                siganlIn = siganlIn + 1
                tERR.Anotar "acei90:" + CStr(siganlIn)
                Form_KeyUp TeclaNewFicha, 0  'especial directo 'SendKeys Chr(TeclaNewFicha)
                
            Case TeclaNewFicha2x2:
                siganlIn = siganlIn + 1
                tERR.Anotar "acei91:" + CStr(siganlIn)
                Form_KeyUp TeclaNewFicha2, 0 'SendKeys Chr(TeclaNewFicha2)
        End Select
    End If
    
    'vaciarlos !!!
    txtS3 = ""
    
End Sub

Public Function WaitOk(sCancion As String) As Long
    
    'mostrar en el mismo pic un preaviso de que va a empezar el karaoke
    Dim EspCancion As String
    EspCancion = TR.Trad("Presione tecla Derecha para comenzar el " + _
        "karaoke o Izquierda para salir." + vbCrLf + _
        "Cancione elegida: %99%") + fso.GetBaseName(sCancion)
    
    LastTecla = 0
    
    MP3.DoClose IAA
    'al iniciar el sistema si estaba pendiente no se ve!!!
    Me.Visible = True
    Me.Refresh
    
    If Salida2 Then
        frmVIDEO.picKAR_V.AutoRedraw = True
        frmVIDEO.lblWAIT_V.Width = frmVIDEO.picKAR_V.Width / 2
        frmVIDEO.lblWAIT_V.Left = frmVIDEO.picKAR_V.Width / 2 - frmVIDEO.lblWAIT_V.Width / 2
        frmVIDEO.lblWAIT_V.Height = 8000
        frmVIDEO.lblWAIT_V.Top = 800
        frmVIDEO.picKAR_V.Visible = True
        frmVIDEO.picKAR_V.ZOrder
        frmVIDEO.lblWAIT_V.Caption = EspCancion
        frmVIDEO.lblWAIT_V.Visible = True
        frmVIDEO.LF1_V.Visible = True
        frmVIDEO.LF2_V.Visible = True
        frmVIDEO.LF1_V.ZOrder
        picVideo(0).Visible = False
        picVideo(1).Visible = False
    Else
        picKAR.AutoRedraw = True
        lblWait.Width = picKAR.Width / 2
        lblWait.Left = picKAR.Width / 2 - lblWait.Width / 2
        lblWait.Height = 8000
        lblWait.Top = 800
        picKAR.Visible = True
        picKAR.ZOrder
        lblWait.Caption = EspCancion
        lblWait.Visible = True
        LF1.Visible = True
        LF2.Visible = True
        LF1.ZOrder
    End If
    
    Dim RR As Single
    RR = Timer
    Dim RR2 As Long 'tiempo que falta para autocomenzar
    
    Do
        DoEvents
        If LastTecla = TeclaDER Then
            lblWait.Visible = False
            frmVIDEO.lblWAIT_V.Visible = False
            Exit Do
        End If
        
        If LastTecla = TeclaIZQ Then
            'ver donde se indica que esta fuera
            'si no todavia piensa que esta reproduciendo un video!!
            
            EsVideo = False
            EsKar = False
            EstoyEnModoVideoMiniSelDisco = False
        
            WaitOk = 1000
            frmVIDEO.picKAR_V.Visible = False
            picKAR.Visible = False
            
            UpdateVista
            Exit Function 'salir e ir al que sigue
        End If
        
        RR2 = 30 - (Timer - RR)
        If Salida2 Then
            frmVIDEO.LF1_V.Caption = RR2
            frmVIDEO.LF2_V = frmVIDEO.LF1_V
        Else
            LF1.Caption = RR2
            LF2.Caption = LF1.Caption
        End If
        If RR2 <= 0 Then Exit Do
        
    Loop
    
    lblWait.Visible = False
    frmVIDEO.lblWAIT_V.Visible = False
    
    Dim sCancion2 As String
    sCancion2 = SYSfolder + "nowpl.mas"
    tERR.Anotar "DD11", sCancion2, Salida2
    'VER SI ESTA ENCRIPTADO O NO!
    If LCase(Right(sCancion, 3)) = "mn1" Then
        'desencriptarlo!
        'cada karaoke pertenece a una coleccion o CD con un identificador
        'estos son los primeros X bytes
        'en base a este yo se que clave le corresponde
        
        'probar uno por uno los CDs existentes
        Dim KYY As String, PX As String 'clave,prefijo encontrados
        KYY = GetH(sCancion, PX)
        
        Dim NOP As String
        If KYY = "NOIDENTIFICOCD" Then
            'este no pertenece a ningun cd oficial de tbrSoft de karaoke
            WaitOk = 1005
            'avisar lo que paso!!!!!!!!!
            
            NOP = TR.Trad("Este Karaoke no pertenece " + _
                    "a ningún cd oficial de tbrSoft%99%")
            If Salida2 Then
                frmVIDEO.lblWAIT_V.Visible = True
                frmVIDEO.lblWAIT_V.Caption = NOP
            Else
                lblWait.Visible = True
                lblWait.Caption = NOP
            End If

            RR = Timer
            Do
                DoEvents
                RR2 = 5 - (Timer - RR)
                If RR2 <= 0 Then Exit Do
                If Salida2 Then
                    frmVIDEO.LF1_V.Caption = RR2
                    frmVIDEO.LF2_V = frmVIDEO.LF1_V
                Else
                    LF1.Caption = RR2
                    LF2.Caption = LF1.Caption
                End If
            Loop
            
            frmVIDEO.picKAR_V.Visible = False
            picKAR.Visible = False
            Exit Function 'salir e ir al que sigue
        End If
        
        If KYY = "NIBOSTA" Then
            'si existe en la coleccion pero no tiene licencia para el
            WaitOk = 1009
            'avisar lo que paso!!!!!!!!!
            NOP = TR.Trad("No tiene la licencia para ejecutar este karaoke%98%" + _
                "Además de la licencia de 3PM cada CD de karaoke oficial " + _
                "tiene su licencia propia%99%")
            
            If Salida2 Then
                frmVIDEO.lblWAIT_V.Visible = True
                frmVIDEO.lblWAIT_V.Caption = NOP
            Else
                lblWait.Visible = True
                lblWait.Caption = NOP
            End If
            
            RR = Timer
            Do
                DoEvents
                RR2 = 5 - (Timer - RR)
                If RR2 <= 0 Then Exit Do
                If Salida2 Then
                    frmVIDEO.LF1_V.Caption = RR2
                    frmVIDEO.LF2_V = frmVIDEO.LF1_V
                Else
                    LF1.Caption = RR2
                    LF2.Caption = LF1.Caption
                End If
            Loop
            
            EsVideo = False
            EsKar = False
            EstoyEnModoVideoMiniSelDisco = False
            frmVIDEO.picKAR_V.Visible = False
            picKAR.Visible = False
            
            UpdateVista
            Exit Function 'salir e ir al que sigue
        End If
        '**************************************
        'LISTO SI ESTA HABILITADO Y TENGO CLAVE
        MP3.doTem True, KYY, sCancion, sCancion2, PX
        '**************************************
    Else 'es un MN0 desencriptado
        fso.CopyFile sCancion, sCancion2, True
    End If
    
    Dim R As Long
    If Salida2 Then
        R = MP3.DoOpenKar(sCancion2, frmVIDEO.picKAR_V, frmVIDEO.shKAR_V)
    Else
        R = MP3.DoOpenKar(sCancion2, picKAR, shKAR)
    End If
    
    'no estaba y aparentemente quedaba en cero lo que lño cai limpiar el pic y se iva todo al karajo
    'pero la musica seguia ???
    TotalTema(2) = frmIndex.MP3.LengthInSec(2)
    
    If R > 0 Then
        WaitOk = R
        tERR.AppendLog "DDa" + CStr(R)
        Exit Function
    End If
    
    WaitOk = 0
    If Salida2 Then
        frmVIDEO.picKAR_V.Visible = True
        frmVIDEO.picKAR_V.ZOrder
    Else
        picKAR.Visible = True
        picKAR.ZOrder
    End If
    'xxxx ver que diga si lo va a querer grabar a bluetooth al final
    MP3.DoPlayKar
End Function

Private Function GetH(AR As String, PX As String) As String
    'devuelve la clave para abrirlo
    'solo se ingresa el archivo MN1
    'en el parametro PX devuelve el prefijo encontrado
    
    Dim KKY As String
    KKY = "NOIDENTIFICOCD"
    PX = "NO"
    
    Dim J As Long
    Dim resPX As String
    For J = 0 To 6 'pruebo todos los cds posibles
        resPX = MP3.GetPrefixKar(AR, Len(CDK_prefix(J)))
        
        If resPX = CDK_prefix(J) Then
            'encontre a que cd pertenece!!!
            'VER SI TIENE LICENCIA PARA EL CD1
            If K.sabseee("mLicenciaCD001Kar") >= EComun Then
                'CON QUE TENGA LA LICENCIA DEL 1 ALCAZA (nuevo set 08)
                'la posibilidad de grabacion sera en superlicencia de karaoke
                
                'si tiene la licencia
                KKY = CDK_qey(J)
                PX = CDK_prefix(J)
            Else 'se que cd es pero no tioene licencia para este
                KKY = "NIBOSTA"
                'ni bosta
                'no puede usar este CD
            End If
            'salgo, otro cd no va a haber ...
            Exit For
        End If
    Next J
    
    GetH = KKY
End Function
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    lblNOCREDIT.Caption = TR.Trad("CREDITO INSUFICIENTE%99%")
    lblCanciones(0).Caption = TR.Trad("Lista de canciones%99%")
    lblDISCO(0).Caption = TR.Trad("Complete al menos la primera hoja de discos cargados%98%Hoja se refiere a una pagina completa con discos%99%")
    lblDisco2(0).Caption = lblDISCO(0).Caption
    lblDiscoSEL.Caption = lblDISCO(0).Caption
    lblDiscoSEL2.Caption = lblDISCO(0).Caption
    lblCanciones2(0).Caption = TR.Trad("Lista de canciones%99%")
    Label1.Caption = TR.Trad("VERSION DEMOSTRATIVA%99%" + _
        " tbrSoft Argentina. www.tbrsoft.com")
    lblTEMAS.Caption = TR.Trad("Temas del disco elegido%99%")
    lblModoVideo.Caption = TR.Trad("Discos en Modo Video%98%Lista de textos " + _
        "de los discos (sin imagenes) por que se esta ejecutando un video " + _
        "o karaoke %99%")
End Sub


'***********************************************************************
Private Sub tBT_Change()
    If tBT.tExt = "" Then Exit Sub
    
    Dim SP() As String
    SP = Split(tBT.tExt, "|")
    
    Select Case SP(0)
        Case "0"
            
        Case "1" 'sale drive
            'termino de buscar dispositivos
            tERR.Anotar "BTM_IF"
        Case "2"
            'connection service status
            tERR.Anotar "BTM_CSR:", SP(1)
        Case "3"
            tERR.Anotar "BTM_SND_OK"
        Case "4"
            tERR.Anotar "BTM_SND_BAD"
        Case "5"
            'encontro un dispositivo
            tERR.Anotar "BTM_DEV", SP(1), SP(2)
    End Select
    
    tBT.tExt = ""
End Sub

Private Sub SuperSel(ByVal Index As Integer)
    tERR.Anotar "PachaMode", PachaMode
    If PachaMode = 10000 Then SuperSel2 Index 'el original
    If PachaMode = 11000 Then SuperSel3 Index 'el de pacha
End Sub

'modo de seleccion de discos segun pacha
Private Sub SuperSel3(ByVal Index As Integer)
        
    On Local Error GoTo ErrSSel
    
    'elegir el disco normalmente
    SelDisco CLng(Index)
    
    Dim IMF As String
    IMF = ExtraData.getDef.getImagePath("tocuharribacomun")
    t1.Picture = LoadPicture(IMF)

    IMF = ExtraData.getDef.getImagePath("touchabajocomun")
    t3.Picture = LoadPicture(IMF)
    
    btOKPacha.Caption = "Escuchar canción"
    
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
    
    imgListaSong.Visible = False
    imgListaSong.Stretch = True
    
    IMF = ExtraData.getDef.getImagePath("MarcoFondodelosdiscos")
    imgListaSong.Picture = LoadPicture(IMF)
    picExtraObjeto.Picture = imgListaSong.Picture
    Dim IND As Long
    IND = ExtraData.getDef.GetIndexImage("MarcoChicoIndicadores")
    Dim MargDer As Long, MargIzq As Long, MargSup As Long, MargInf As Long
    MargSup = imgListaSong.Height * ExtraData.getDef.GetFinalMargenSuperiorTra(IND) / 100
    MargInf = imgListaSong.Height * ExtraData.getDef.GetFinalMargenInferiorTra(IND) / 100
    MargDer = imgListaSong.Width * ExtraData.getDef.GetFinalMargenDerechoTra(IND) / 100
    MargIzq = imgListaSong.Width * ExtraData.getDef.GetFinalMargenIzquierdoTra(IND) / 100
    
    imgListaSong.Top = 150
    If MostrarTouch Then
        imgListaSong.Height = (picFondoDisco.Height) - imgSELEC.Height - 300 - lblNOCREDIT.Height
    Else
        imgListaSong.Height = (picFondoDisco.Height) - 300
    End If
    
    imgListaSong.Left = 500
    imgListaSong.Width = picFondoDisco.Width - imgListaSong.Left - (picFondoDisco.Width / 5) - 500
    imgListaSong.Visible = True
    
    imgDiscoSEL.Visible = False
    imgDiscoSEL.Stretch = True
    imgDiscoSEL.Picture = TapaCD(Index).Picture
    imgDiscoSEL.Width = (picFondoDisco.Width / 5)
    imgDiscoSEL.Height = (picFondoDisco.Height / 4)
    imgDiscoSEL.Top = 300
    imgDiscoSEL.Left = imgListaSong.Left + imgListaSong.Width + 200
    imgDiscoSEL.Visible = True
    
    lblDiscoSEL.Visible = False
    lblDiscoSEL2.Visible = False
    
    lblDiscoSEL.Caption = lblDISCO(Index).Caption
    lblDiscoSEL.Font.Size = lblDISCO(Index).Font.Size
    lblDiscoSEL.Top = imgDiscoSEL.Top + imgDiscoSEL.Height
    lblDiscoSEL.Left = imgDiscoSEL.Left
    lblDiscoSEL.Width = imgDiscoSEL.Width - 200
    lblDiscoSEL.Height = 500
    
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
    imgFondoDiscoSel.Width = (picFondoDisco.Width / 5) + 200
    imgFondoDiscoSel.Height = imgDiscoSEL.Height + lblDiscoSEL.Height + 200
    imgFondoDiscoSel.Top = imgDiscoSEL.Top - 100
    imgFondoDiscoSel.Left = imgDiscoSEL.Left - 200
    imgFondoDiscoSel.Visible = True
    
    imgDiscoSEL.ZOrder
    imgFondoDiscoSel.ZOrder
    lblDiscoSEL2.ZOrder
    lblDiscoSEL.ZOrder
    
    'ya esta agrandado
    lblDATA.Font.Size = lblDiscoSEL2.Font.Size
    lblDATA.Width = imgFondoDiscoSel.Width - 200
    lblDATA.Height = picFondoDisco.Height - (imgFondoDiscoSel.Top + imgFondoDiscoSel.Height)
    lblDATA.Left = imgFondoDiscoSel.Left + 100
    lblDATA.Top = imgFondoDiscoSel.Top + imgFondoDiscoSel.Height + 100
    
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
    Dim PerfilEncontrado As Long
    If UbicDiscoActual = "_RANK_" Then
        MATRIZ_TEMAS = ObtenerRankComoMM(30) 'joia tengo un ranking decente!!!
        PerfilActual = -1 'perfil de ranking
    Else
        'OM- entro a un disco y quiero ver lo que hay
        If VentaExtras Then
            PerfilEncontrado = 1 'valor para que entre buscando
            MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual, True, PerfilEncontrado)
        Else
            'parta que ni busque perfiles, es un disco comun de 3PM
            MATRIZ_TEMAS = ObtenerArchMM(UbicDiscoActual, True)
            PerfilEncontrado = 1
        End If
        
        'de la forma!
        'D:\musica\Cuartetazo\Alma Fuerte-En vivo Obras 2001\02 - Almafuerte.mp3#02 - Almafuerte.mp3
        PerfilActual = PerfilEncontrado
    End If
    tERR.Anotar "caah3", UBound(MATRIZ_TEMAS)
    'usar esto y no una variable para saber de discos vacios
    If UBound(MATRIZ_TEMAS) = 0 Then
        lblCanciones(0).Caption = TR.Trad("NO HAY CANCIONES EN ESTE DISCO!%99%")
        tERR.AppendLog "No hay temas en el disco: " + UbicDiscoActual + ".acpu"
        EstoyEnDisco = 1 'PARA QUE PUEDA SALIR!!!!
        Exit Sub
    End If
    
    'ordenar la lista de canciones
    Dim DataTXT As String
    If UbicDiscoActual = "_RANK_" Then
        DataTXT = TR.Trad("Estos son los mas escuchados!%99%")
    Else
        Dim ArchDaTa As String
        ArchDaTa = UbicDiscoActual + "data.txt"
        If fso.FileExists(ArchDaTa) Then
            Dim A As TextStream
            Set A = fso.OpenTextFile(ArchDaTa, ForReading, False)
                If A.AtEndOfStream = False Then
                    DataTXT = A.ReadAll
                Else
                    DataTXT = TR.Trad("No hay datos adicionales de este disco%99%")
                End If
            A.Close
        Else
            DataTXT = TR.Trad("No hay datos adicionales de este disco%99%")
        End If
    End If
    
    Select Case PerfilActual
        Case 1: DataTXT = DataTXT + vbCrLf + "PERFIL 3PM BASE"
        Case 2: DataTXT = DataTXT + vbCrLf + "PERFIL RINGTONES"
        Case 3: DataTXT = DataTXT + vbCrLf + "PERFIL WALLPAPERS"
        Case 4: DataTXT = DataTXT + vbCrLf + "PERFIL JAVA"
    End Select
    
    lblDATA.Caption = DataTXT:    lblDATA2.Caption = lblDATA.Caption
    lblDATA.Visible = True:       lblDATA2.Visible = True
    
    Dim C As Integer, nombreTemas As String
    Dim pathTema As String
    C = 1
    Dim AltoRenglon As Long
    AltoRenglon = lblCanciones(0).Height + 30
    tERR.Anotar "caai", AltoRenglon
    Dim EXT As String

    'establecer los limites donde van los elemntos para leer despues
    'limite superior = imgListaSong.Top + imgListaSong.Height - AltoRenglon - MargInf
    'limite inferior = MargSup +  AltoRenglon
    imgListaSong.Tag = "LS:" + _
                       CStr(imgListaSong.Top + imgListaSong.Height - AltoRenglon - MargInf) + _
                       "|LI:" + _
                       CStr(imgListaSong.Top + MargSup + 90) + _
                       "|MD:" + _
                       CStr(MargDer) + _
                       "|MI:" + _
                       CStr(MargIzq)
    
    Do While C <= UBound(MATRIZ_TEMAS)
        pathTema = txtInLista(MATRIZ_TEMAS(C), 0, "#")
        nombreTemas = txtInLista(MATRIZ_TEMAS(C), 1, "#")
        EXT = LCase(txtInLista(pathTema, 1, "."))
        
        'quitar el molesto .mp3 o lo que fuera
        Select Case LCase(EXT)
            Case "mp3"
                EXT = "" 'se sobreentiende que todo es mp3" (mp3-Musica)"
'            Case "mp4"
'                EXT = " (mp4-Musica)"
            Case "wma"
                EXT = TR.Trad(" (wma-Musica)%99%")
            Case "mpeg", "mpg", "avi", "wmv"
                TR.SetVars LCase(EXT)
                EXT = TR.Trad(" (%01%-Video)%98%La variable 1" + _
                    "es MPG o AVI o WMV, es el formato de video " + _
                    "de un archivo%99%")
            Case "vob"
                EXT = TR.Trad(" (DVD!)%99%")
            Case "dat"
                EXT = TR.Trad(" (VCD-Video)%99%")
            Case "mn0", "mn1"
                EXT = TR.Trad(" (KARAOKE)%99%")
            Case "jpg", "jpeg", "bmp", "gif"
                EXT = TR.Trad(" (Wallpaper)%99%")
            Case "jar"
                EXT = TR.Trad(" (Java)%99%")
            'mm91
            'formatos de imagenes de nero
            'NR3: cd de mp3s    /    'NRA: cd de audio    /  'NRB: cd-rom de arranque
            'NRC: nero usf/iso  /    'NRD: nero DVD       /  'NRE: cd extra
            'NRG: imagen        /    'NRH: cd-rom hibrido /  'NRI: cd-rom iso
            'NRM: cd mixto      /    'NRU: cd-rom udf     /  'NRV: cd supervideo
            'NRW: cd rom wma    /    'CDC: cd cover no tiene nada que ver con imagenes parece
            Case "iso"
                EXT = TR.Trad(" (Imagen ISO)%99%")
            Case "nrg"
                EXT = TR.Trad(" (Imagen NERO)%99%")
            Case "nr3"
                EXT = TR.Trad(" (Imagen NERO MP3)%99%")
            Case "nra"
                EXT = TR.Trad(" (Imagen NERO AUDIO)%99%")
            Case "nrb"
                EXT = TR.Trad(" (Imagen NERO INICIO)%99%")
            Case "nrc"
                EXT = TR.Trad(" (Imagen NERO UDF/ISO)%99%")
            Case "nrd"
                EXT = TR.Trad(" (Imagen NERO DVD)%99%")
            Case "nre"
                EXT = TR.Trad(" (Imagen NERO CD EXTRA)%99%")
            Case "nrh"
                EXT = TR.Trad(" (Imagen NERO CD HIBR)%99%")
            Case "nri"
                EXT = TR.Trad(" (Imagen NERO ISO)%99%")
            Case "nrm"
                EXT = TR.Trad(" (Imagen NERO CD MIXTO)%99%")
            Case "nru"
                EXT = TR.Trad(" (Imagen NERO CD UDF)%99%")
            Case "nrv"
                EXT = TR.Trad(" (Imagen NERO SUPERVIDEO)%99%")
            Case "nrw"
                EXT = TR.Trad(" (Imagen NERO WMA)%99%")
            Case "3gp"
                EXT = TR.Trad(" (Video para movil)%99%")
        End Select
        nombreTemas = fso.GetBaseName(nombreTemas) + EXT
        Load lblCanciones(C)
        Load lblCanciones2(C)
        
        lblCanciones(C).Caption = nombreTemas
        tERR.Anotar "caaj", C, nombreTemas
        lblCanciones(C).Tag = pathTema
        lblCanciones(C).Top = imgListaSong.Top + MargSup + 90 + ((C - 1) * AltoRenglon)
        lblCanciones2(C).Top = lblCanciones(C).Top + 15
        lblCanciones2(C).Left = lblCanciones(C).Left + 15
        'tiene autosize
        
        C = C + 1 'ver que el proximo entre
    Loop
    
    Dim TotalSong As Long
    TotalSong = C - 1
    'en adelante se usa como referencia el ubound asi que lo corto directamente asi!
    'ReDim Preserve MATRIZ_TEMAS(TotalSong)
    'no se corta mas porque se muestra todo
    
    If CargarDuracionTemas Then
        'ahora cargar las duaciones
        Dim NoCargoDuracion As Long
        NoCargoDuracion = 0
        C = 1
        Dim MP3tmp As New MP3Info
        Do While C <= UBound(MATRIZ_TEMAS)
            pathTema = lblCanciones(C).Tag
            'si es mp3 usar el rápido, si no usar el viejo
            'mm91 no se puede tener la duracion de otros tipos de archivos
            Dim est As String
            est = UCase(Right(pathTema, 3))
            Select Case est
                Case "MP3"
                    MP3tmp.FileName = pathTema
                    DuracionTema = MP3tmp.DurationSTR
                Case "mpg", "avi", "mpeg"
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
            End Select
            lblCanciones(C).Caption = lblCanciones(C).Caption + " (" + DuracionTema + ")"
            C = C + 1
        Loop
        Set MP3tmp = Nothing
    End If

    'revisar especificamente que no haya nada mas largo que lo que se puede
    C = 1
    Do While C <= UBound(MATRIZ_TEMAS)
        'si o si dejar un margen
        If lblCanciones(C).Width > (imgListaSong.Width * 0.9) Then
            Dim D As Long
            For D = 1 To 35 'con estas pasadas debe quedar ok
                'que nunca de error!!!!
                If Len(lblCanciones(C).Caption) > 10 Then
                    lblCanciones(C).Caption = _
                        Mid(lblCanciones(C).Caption, 1, Len(lblCanciones(C).Caption) - 10) + "..."
                Else
                    Exit For
                End If
                'ver si con eso alcanza
                If lblCanciones(C).Width < (imgListaSong.Width * 0.9) Then Exit For
            Next D
        End If
        
        C = C + 1
    Loop


    C = 1
    Do While C <= UBound(MATRIZ_TEMAS)
        lblCanciones(C).Left = imgListaSong.Left + (imgListaSong.Width / 2 - lblCanciones(C).Width / 2)
        lblCanciones2(C).Left = lblCanciones(C).Left + 15
        
        'ver que no se muestren mas canciones de las que entren
        'estos como se pasan deben ser invisibles
        If lblCanciones(C).Top > (imgListaSong.Top + imgListaSong.Height _
                - AltoRenglon - MargInf) Then
            
            lblCanciones(C).Visible = False
            lblCanciones2(C).Visible = False
            lblCanciones2(C).Tag = "OUT DOWN" 'fuera de la visulización (por debajo)!!
        
        Else
            lblCanciones(C).Visible = True
            lblCanciones2(C).Visible = True
            lblCanciones2(C).Tag = "IN" 'fuera de la visulización !!
        End If
        
        
        lblCanciones2(C).ZOrder 'lo necesito paar poder hacerle click
        lblCanciones(C).ZOrder
        C = C + 1
    Loop
    
    lblNOCREDIT.Left = imgListaSong.Left + (imgListaSong.Width / 2 - lblNOCREDIT.Width / 2)
    
    'SOLO PACHA
    '**********************************************************
    btSalir.Width = imgDiscoSEL.Width
    btBuyCancion.Width = btSalir.Width
    btBUYDisco.Width = btSalir.Width
    'ES DEFORMADO POR LA FUNCION DE ACOMODAR LOS FORMULARIOS !!!!!!!!!!!!!!!!!!!!!!!!!!
    btSalir.Height = 705
    btBuyCancion.Height = 705
    btBUYDisco.Height = 705
    
    btSalir.Left = picFondoDisco.Width - btSalir.Width + 90
    btSalir.Top = picFondoDisco.Height - btSalir.Height
        
    btBuyCancion.Top = btSalir.Top - btBuyCancion.Height - SeparacionTocuhDerecho
    btBUYDisco.Top = btBuyCancion.Top - btBUYDisco.Height - SeparacionTocuhDerecho
    
    btBUYDisco.Left = picFondoDisco.Width - btBUYDisco.Width + 90
    btBuyCancion.Left = picFondoDisco.Width - btBuyCancion.Width + 90
    
    btBuyCancion.Visible = True
    btBUYDisco.Visible = True
    btSalir.Visible = True
    
    btBuyCancion.ZOrder
    btBUYDisco.ZOrder
    btSalir.ZOrder
    '**********************************************************
    
    If MostrarTouch Then
        cmdTouchAbajo.Top = 120 'imgListaSong.Top + cmdTouchArriba.Height + 120
        cmdTouchAbajo.Left = (imgListaSong.Left / 2 - cmdTouchAbajo.Width) ' - 120
        
        cmdTouchArriba.Top = 120 'imgListaSong.Top + 120
        cmdTouchArriba.Left = (imgListaSong.Left / 2) '+ 120
        
        cmdTouchArriba.Visible = True
        cmdTouchAbajo.Visible = True
        
        btBUYDisco.Top = cmdTouchArriba.Top + cmdTouchArriba.Height
        btBuyCancion.Top = btBUYDisco.Top + btBUYDisco.Height + 60
        
        btBUYDisco.Left = 120
        btBuyCancion.Left = 120
                
        If VendoMusica Then
            btBuyCancion.Visible = True
            btBUYDisco.Visible = True
        End If
        
        imgSELEC.Left = imgListaSong.Left + 90 ' imgListaSong.Left + (imgListaSong.Width / 3 - imgSELEC.Width)
        imgSELEC.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        ImgSelecVIP.Left = imgSELEC.Left + imgSELEC.Width + 60
        ImgSelecVIP.Top = imgSELEC.Top
        
        imgSALIR.Left = ImgSelecVIP.Left + ImgSelecVIP.Width + 60
        imgSALIR.Top = imgListaSong.Top + imgListaSong.Height + 60
        
        lblNOCREDIT.Top = imgSELEC.Height + imgSELEC.Top + 60
        
        imgSELEC.Visible = True
        imgSALIR.Visible = True
        'si esta activado lo uso
        ImgSelecVIP.Visible = (CreditosXaVipMusica > 0) And (PerfilActual = 1)
    Else
'        lblNOCREDIT.Top = imgListaSong.Height + imgListaSong.Top - lblNOCREDIT.Height - 120
'        btBuyCancion.Visible = False
'        btBUYDisco.Visible = False
    End If
    
    EstoyEnDisco = 1
    OkInState1 = 0
    selDiscoI 1
    
    Exit Sub
    
ErrSSel:
    tERR.AppendLog "SSEL444-3", tERR.ErrToTXT(Err)
    Resume Next

End Sub
