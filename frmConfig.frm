VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
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
   Begin VB.Frame frKKAR 
      BackColor       =   &H00000000&
      Caption         =   "Karaokes"
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
      Height          =   3045
      Left            =   7680
      TabIndex        =   264
      Top             =   5730
      Visible         =   0   'False
      Width           =   8865
      Begin VB.CheckBox chkGrabaKarQuick 
         BackColor       =   &H00000000&
         Caption         =   "Ofrecer la grabación de karaoke ni bien termina la canción"
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
         Height          =   300
         Left            =   150
         TabIndex        =   270
         Top             =   2010
         Width           =   6795
      End
      Begin VB.ComboBox cmbKbpsKar 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmConfig.frx":0442
         Left            =   4590
         List            =   "frmConfig.frx":046D
         Style           =   2  'Dropdown List
         TabIndex        =   267
         Top             =   1200
         Width           =   1185
      End
      Begin VB.ComboBox cmbGrabaKar 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmConfig.frx":04AC
         Left            =   4620
         List            =   "frmConfig.frx":04B9
         Style           =   2  'Dropdown List
         TabIndex        =   265
         Top             =   330
         Width           =   3975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmConfig.frx":04F8
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
         Index           =   32
         Left            =   600
         TabIndex        =   271
         Top             =   2340
         Width           =   7935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "kbps"
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
         Height          =   285
         Index           =   24
         Left            =   5820
         TabIndex        =   269
         Top             =   1230
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Compresion del MP3 resultante (mas es mayor calidad pero mayor tamaño de archivo). Se recomienda 128 Kbps"
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
         Height          =   615
         Index           =   23
         Left            =   120
         TabIndex        =   268
         Top             =   1020
         Width           =   4485
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Grabación de karaokes, asegúrese de disponer de una SuperLicencia de karaoke"
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
         Height          =   435
         Index           =   22
         Left            =   150
         TabIndex        =   266
         Top             =   270
         Width           =   4305
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
      Height          =   5220
      Left            =   3090
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   8835
      Begin tbrFaroButton.fBoton fBoton4 
         Height          =   525
         Left            =   6150
         TabIndex        =   257
         Top             =   4170
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "LEDs del teclado"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command28 
         Height          =   555
         Left            =   3090
         TabIndex        =   240
         Top             =   4440
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "especiales monedero"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.PictureBox PicContLetras 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2685
         Left            =   120
         ScaleHeight     =   2685
         ScaleWidth      =   8655
         TabIndex        =   81
         Top             =   180
         Width           =   8655
         Begin VB.PictureBox PicLetras 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   6465
            Left            =   30
            ScaleHeight     =   6465
            ScaleWidth      =   7995
            TabIndex        =   82
            Top             =   0
            Width           =   7995
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
               Index           =   18
               Left            =   2430
               TabIndex        =   262
               Top             =   7980
               Width           =   570
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               Index           =   18
               ItemData        =   "frmConfig.frx":0588
               Left            =   7020
               List            =   "frmConfig.frx":05D4
               Style           =   2  'Dropdown List
               TabIndex        =   260
               Top             =   6090
               Width           =   945
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
               Index           =   18
               ItemData        =   "frmConfig.frx":0632
               Left            =   2010
               List            =   "frmConfig.frx":075F
               Style           =   2  'Dropdown List
               TabIndex        =   259
               Top             =   6090
               Width           =   5000
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
               Index           =   17
               Left            =   7980
               TabIndex        =   254
               Top             =   5670
               Width           =   570
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
               Index           =   17
               ItemData        =   "frmConfig.frx":0EE1
               Left            =   2010
               List            =   "frmConfig.frx":100E
               Style           =   2  'Dropdown List
               TabIndex        =   252
               Top             =   5760
               Width           =   5000
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               Index           =   17
               ItemData        =   "frmConfig.frx":1790
               Left            =   7020
               List            =   "frmConfig.frx":17DC
               Style           =   2  'Dropdown List
               TabIndex        =   251
               Top             =   5760
               Width           =   945
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
               Index           =   16
               Left            =   5310
               TabIndex        =   230
               Top             =   7920
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   16
               ItemData        =   "frmConfig.frx":183A
               Left            =   2010
               List            =   "frmConfig.frx":1967
               Style           =   2  'Dropdown List
               TabIndex        =   228
               Top             =   1440
               Width           =   5000
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   16
               ItemData        =   "frmConfig.frx":20E9
               Left            =   7020
               List            =   "frmConfig.frx":2135
               Style           =   2  'Dropdown List
               TabIndex        =   227
               Top             =   1440
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               ItemData        =   "frmConfig.frx":2193
               Left            =   7015
               List            =   "frmConfig.frx":21DF
               Style           =   2  'Dropdown List
               TabIndex        =   208
               Top             =   90
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               ItemData        =   "frmConfig.frx":223D
               Left            =   7015
               List            =   "frmConfig.frx":2289
               Style           =   2  'Dropdown List
               TabIndex        =   207
               Top             =   420
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               ItemData        =   "frmConfig.frx":22E7
               Left            =   7015
               List            =   "frmConfig.frx":2333
               Style           =   2  'Dropdown List
               TabIndex        =   206
               Top             =   765
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               ItemData        =   "frmConfig.frx":2391
               Left            =   7015
               List            =   "frmConfig.frx":23DD
               Style           =   2  'Dropdown List
               TabIndex        =   205
               Top             =   1095
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   4
               ItemData        =   "frmConfig.frx":243B
               Left            =   7020
               List            =   "frmConfig.frx":2487
               Style           =   2  'Dropdown List
               TabIndex        =   204
               Top             =   1770
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   5
               ItemData        =   "frmConfig.frx":24E5
               Left            =   7015
               List            =   "frmConfig.frx":2531
               Style           =   2  'Dropdown List
               TabIndex        =   203
               Top             =   2085
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   6
               ItemData        =   "frmConfig.frx":258F
               Left            =   7015
               List            =   "frmConfig.frx":25DB
               Style           =   2  'Dropdown List
               TabIndex        =   202
               Top             =   2400
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":2639
               Left            =   7015
               List            =   "frmConfig.frx":2685
               Style           =   2  'Dropdown List
               TabIndex        =   201
               Top             =   2730
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":26E3
               Left            =   7015
               List            =   "frmConfig.frx":272F
               Style           =   2  'Dropdown List
               TabIndex        =   200
               Top             =   3075
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":278D
               Left            =   7015
               List            =   "frmConfig.frx":27D9
               Style           =   2  'Dropdown List
               TabIndex        =   199
               Top             =   3405
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":2837
               Left            =   7015
               List            =   "frmConfig.frx":2883
               Style           =   2  'Dropdown List
               TabIndex        =   198
               Top             =   3720
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":28E1
               Left            =   7015
               List            =   "frmConfig.frx":292D
               Style           =   2  'Dropdown List
               TabIndex        =   197
               Top             =   4050
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":298B
               Left            =   7015
               List            =   "frmConfig.frx":29D7
               Style           =   2  'Dropdown List
               TabIndex        =   196
               Top             =   4365
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":2A35
               Left            =   7015
               List            =   "frmConfig.frx":2A81
               Style           =   2  'Dropdown List
               TabIndex        =   195
               Top             =   4725
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":2ADF
               Left            =   7015
               List            =   "frmConfig.frx":2B2B
               Style           =   2  'Dropdown List
               TabIndex        =   194
               Top             =   5070
               Width           =   945
            End
            Begin VB.ComboBox cmbTECLAS2 
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
               ItemData        =   "frmConfig.frx":2B89
               Left            =   7015
               List            =   "frmConfig.frx":2BD5
               Style           =   2  'Dropdown List
               TabIndex        =   193
               Top             =   5400
               Width           =   945
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
               Index           =   15
               Left            =   4350
               TabIndex        =   187
               Top             =   10770
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
               ItemData        =   "frmConfig.frx":2C33
               Left            =   2010
               List            =   "frmConfig.frx":2D60
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   5400
               Width           =   5000
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
               ItemData        =   "frmConfig.frx":34E2
               Left            =   2010
               List            =   "frmConfig.frx":360F
               Style           =   2  'Dropdown List
               TabIndex        =   112
               Top             =   3075
               Width           =   5000
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
               ItemData        =   "frmConfig.frx":3D91
               Left            =   2010
               List            =   "frmConfig.frx":3EBE
               Style           =   2  'Dropdown List
               TabIndex        =   111
               Top             =   2745
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   6
               ItemData        =   "frmConfig.frx":4640
               Left            =   2010
               List            =   "frmConfig.frx":476D
               Style           =   2  'Dropdown List
               TabIndex        =   110
               Top             =   2415
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   5
               ItemData        =   "frmConfig.frx":4EEF
               Left            =   2010
               List            =   "frmConfig.frx":501C
               Style           =   2  'Dropdown List
               TabIndex        =   109
               Top             =   2085
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   4
               ItemData        =   "frmConfig.frx":579E
               Left            =   2010
               List            =   "frmConfig.frx":58CB
               Style           =   2  'Dropdown List
               TabIndex        =   108
               Top             =   1755
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   3
               ItemData        =   "frmConfig.frx":604D
               Left            =   2010
               List            =   "frmConfig.frx":617A
               Style           =   2  'Dropdown List
               TabIndex        =   107
               Top             =   1095
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   2
               ItemData        =   "frmConfig.frx":68FC
               Left            =   2010
               List            =   "frmConfig.frx":6A29
               Style           =   2  'Dropdown List
               TabIndex        =   106
               Top             =   765
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   1
               ItemData        =   "frmConfig.frx":71AB
               Left            =   2010
               List            =   "frmConfig.frx":72D8
               Style           =   2  'Dropdown List
               TabIndex        =   105
               Top             =   435
               Width           =   5000
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
               ForeColor       =   &H00404040&
               Height          =   315
               Index           =   0
               ItemData        =   "frmConfig.frx":7A5A
               Left            =   2010
               List            =   "frmConfig.frx":7B87
               Style           =   2  'Dropdown List
               TabIndex        =   104
               Top             =   90
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   103
               Top             =   7725
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
               Left            =   4365
               TabIndex        =   102
               Top             =   8055
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
               Left            =   4365
               TabIndex        =   101
               Top             =   8415
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
               Left            =   4365
               TabIndex        =   100
               Top             =   7410
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
               Left            =   4365
               TabIndex        =   99
               Top             =   7080
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
               Left            =   3690
               TabIndex        =   98
               Top             =   7950
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
               Left            =   4500
               TabIndex        =   97
               Top             =   7920
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
               Left            =   2940
               TabIndex        =   96
               Top             =   7950
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
               Left            =   7200
               TabIndex        =   95
               Top             =   7530
               Width           =   570
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
               ItemData        =   "frmConfig.frx":8309
               Left            =   2010
               List            =   "frmConfig.frx":8436
               Style           =   2  'Dropdown List
               TabIndex        =   94
               Top             =   3405
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   93
               Top             =   8745
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
               ItemData        =   "frmConfig.frx":8BB8
               Left            =   2010
               List            =   "frmConfig.frx":8CE5
               Style           =   2  'Dropdown List
               TabIndex        =   92
               Top             =   3735
               Width           =   5000
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
               Left            =   4410
               TabIndex        =   91
               Top             =   9075
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
               ItemData        =   "frmConfig.frx":9467
               Left            =   2010
               List            =   "frmConfig.frx":9594
               Style           =   2  'Dropdown List
               TabIndex        =   90
               Top             =   4065
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   89
               Top             =   9405
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
               ItemData        =   "frmConfig.frx":9D16
               Left            =   2010
               List            =   "frmConfig.frx":9E43
               Style           =   2  'Dropdown List
               TabIndex        =   88
               Top             =   4395
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   87
               Top             =   9735
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
               ItemData        =   "frmConfig.frx":A5C5
               Left            =   2010
               List            =   "frmConfig.frx":A6F2
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   4740
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   85
               Top             =   10080
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
               ItemData        =   "frmConfig.frx":AE74
               Left            =   2010
               List            =   "frmConfig.frx":AFA1
               Style           =   2  'Dropdown List
               TabIndex        =   84
               Top             =   5070
               Width           =   5000
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
               Left            =   4365
               TabIndex        =   83
               Top             =   10410
               Width           =   700
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tecla suma validacion"
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
               Index           =   21
               Left            =   -480
               TabIndex        =   261
               Top             =   6120
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tecla canción VIP"
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
               Index           =   20
               Left            =   -480
               TabIndex        =   253
               Top             =   5790
               Width           =   2445
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tecla Carrito"
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
               Height          =   240
               Index           =   18
               Left            =   -510
               TabIndex        =   229
               Top             =   1500
               Width           =   2445
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
               TabIndex        =   186
               Top             =   5445
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   0
               Left            =   -510
               TabIndex        =   127
               Top             =   120
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   14
               Left            =   -510
               TabIndex        =   126
               Top             =   2445
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
               Left            =   -540
               TabIndex        =   125
               Top             =   2790
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
               TabIndex        =   124
               Top             =   3075
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   5
               Left            =   -510
               TabIndex        =   123
               Top             =   2145
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   4
               Left            =   -510
               TabIndex        =   122
               Top             =   1815
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   3
               Left            =   -510
               TabIndex        =   121
               Top             =   1155
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   2
               Left            =   -510
               TabIndex        =   120
               Top             =   825
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
               ForeColor       =   &H00E0E0E0&
               Height          =   240
               Index           =   1
               Left            =   -510
               TabIndex        =   119
               Top             =   510
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
               TabIndex        =   118
               Top             =   3405
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
               TabIndex        =   117
               Top             =   3735
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
               TabIndex        =   116
               Top             =   4065
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
               TabIndex        =   115
               Top             =   4395
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
               TabIndex        =   114
               Top             =   4755
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
               TabIndex        =   113
               Top             =   5085
               Width           =   2445
            End
         End
         Begin VB.CommandButton Command24 
            BackColor       =   &H00E0E0E0&
            Height          =   1270
            Left            =   8160
            Picture         =   "frmConfig.frx":B723
            Style           =   1  'Graphical
            TabIndex        =   129
            Top             =   1350
            Width           =   465
         End
         Begin VB.CommandButton Command23 
            BackColor       =   &H00E0E0E0&
            Height          =   1270
            Left            =   8160
            Picture         =   "frmConfig.frx":BB65
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   60
            Width           =   465
         End
      End
      Begin VB.CheckBox chkS3 
         BackColor       =   &H00000000&
         Caption         =   "activar teclado tbrSoft"
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
         Height          =   225
         Left            =   90
         TabIndex        =   219
         Top             =   3090
         Width           =   2520
      End
      Begin VB.TextBox txtFrecTecladoTBR 
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
         Left            =   2640
         TabIndex        =   210
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   3060
         Width           =   600
      End
      Begin VB.VScrollBar vsFrecTecladoTBR 
         Height          =   330
         LargeChange     =   5
         Left            =   3240
         Max             =   5
         Min             =   500
         SmallChange     =   5
         TabIndex        =   209
         Top             =   3060
         Value           =   5
         Width           =   330
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
         Height          =   510
         Left            =   4920
         TabIndex        =   78
         Top             =   3000
         Width           =   3810
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
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   90
            TabIndex        =   79
            Top             =   210
            Width           =   1965
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
            ForeColor       =   &H00E0E0E0&
            Height          =   210
            Left            =   2100
            TabIndex        =   80
            Top             =   210
            Width           =   1650
         End
      End
      Begin VB.CheckBox chkPasarhoja 
         BackColor       =   &H00000000&
         Caption         =   "Pasa páginas con botones Adel-Atras."
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
         Height          =   285
         Left            =   90
         TabIndex        =   33
         Top             =   3630
         Width           =   4860
      End
      Begin VB.CheckBox chkApagarPC 
         BackColor       =   &H00000000&
         Caption         =   "Apagar PC al cerrar sistema."
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
         Height          =   330
         Left            =   90
         TabIndex        =   17
         Top             =   3330
         Width           =   4650
      End
      Begin VB.CheckBox chkUseAPITecla 
         BackColor       =   &H00000000&
         Caption         =   "Recibir las señales del monedero accediendo directamente al teclado (las pulsaciones largas provocan repeticiones)."
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
         Height          =   420
         Left            =   90
         TabIndex        =   212
         Top             =   3990
         Width           =   6150
      End
      Begin VB.CheckBox chkCS 
         BackColor       =   &H00000000&
         Caption         =   "Activar corrección de señales."
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
         Left            =   90
         TabIndex        =   222
         Top             =   4560
         Width           =   3990
      End
      Begin tbrFaroButton.fBoton fBoton6 
         Height          =   525
         Left            =   7470
         TabIndex        =   272
         Top             =   4170
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   926
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Port Address"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "miliSeg"
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
         Height          =   240
         Index           =   55
         Left            =   3120
         TabIndex        =   211
         Top             =   3120
         Width           =   1125
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
      Left            =   5460
      TabIndex        =   63
      Top             =   3600
      Visible         =   0   'False
      Width           =   5385
      Begin VB.CheckBox chkVidMudos 
         BackColor       =   &H00000000&
         Caption         =   "Usar la salida de TV para reproducir videos MUDOS, esto anula las imágenes grandes en el TV."
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
         TabIndex        =   145
         Top             =   2100
         Width           =   4995
      End
      Begin VB.VScrollBar vsPubliIMGCada 
         Height          =   330
         Left            =   4800
         Max             =   10
         Min             =   100
         TabIndex        =   67
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4170
         TabIndex        =   70
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   600
         Width           =   600
      End
      Begin VB.CheckBox ckPubIMG 
         BackColor       =   &H00000000&
         Caption         =   "Reproducir Publicidades (imágenes rotativas)."
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
         Height          =   240
         Left            =   210
         TabIndex        =   66
         Top             =   300
         Width           =   4515
      End
      Begin VB.VScrollBar vsPubliCada 
         Height          =   330
         Left            =   4920
         Max             =   1
         Min             =   100
         TabIndex        =   65
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   4260
         TabIndex        =   68
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
         ForeColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   270
         TabIndex        =   64
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   30
         Left            =   210
         TabIndex        =   71
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
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   29
         Left            =   375
         TabIndex        =   69
         Top             =   1650
         Width           =   3840
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
      Left            =   8220
      TabIndex        =   21
      Top             =   7140
      Visible         =   0   'False
      Width           =   8655
      Begin VB.CheckBox chkQuitaBarraInf 
         BackColor       =   &H00000000&
         Caption         =   "Reducir barra inferior"
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
         Height          =   225
         Left            =   5640
         TabIndex        =   256
         Top             =   2250
         Width           =   2625
      End
      Begin VB.CheckBox chkQuitaBarraSup 
         BackColor       =   &H00000000&
         Caption         =   "Quitar barra superior de ritmos y letras"
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
         Height          =   405
         Left            =   5640
         TabIndex        =   255
         Top             =   1800
         Width           =   2625
      End
      Begin VB.CheckBox chkOutTemasWhenSel 
         BackColor       =   &H00000000&
         Caption         =   "Salir de listado de música al hacer una selección."
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
         Height          =   330
         Left            =   60
         TabIndex        =   135
         Top             =   1710
         Width           =   4875
      End
      Begin VB.CheckBox chkTouch 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar botones touch-screen."
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
         Left            =   60
         TabIndex        =   57
         Top             =   2040
         Width           =   3345
      End
      Begin VB.CheckBox chkMostrarRotulos 
         BackColor       =   &H00000000&
         Caption         =   "Mostrar rótulos de discos."
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
         Left            =   60
         TabIndex        =   24
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
         ForeColor       =   &H00E0E0E0&
         Height          =   420
         Left            =   5640
         TabIndex        =   72
         Top             =   150
         Width           =   2805
      End
      Begin VB.CheckBox chkBloquearMusicaElegida 
         BackColor       =   &H00000000&
         Caption         =   "Evitar selección múltiple de una misma canción en un disco."
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
         Height          =   300
         Left            =   60
         TabIndex        =   74
         Top             =   1440
         Width           =   5475
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
         ForeColor       =   &H00E0E0E0&
         Height          =   405
         Left            =   5640
         TabIndex        =   75
         Top             =   570
         Width           =   2625
      End
      Begin VB.CheckBox chkNoVumVID 
         BackColor       =   &H00000000&
         Caption         =   "Quitar vumetro (medidor de sonido) en videos."
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
         Height          =   330
         Left            =   60
         TabIndex        =   73
         Top             =   1140
         Width           =   4875
      End
      Begin VB.TextBox TxtUSUARIO 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1155
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Text            =   "frmConfig.frx":BFA7
         Top             =   2550
         Width           =   2970
      End
      Begin VB.TextBox txtDiscosV 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   5640
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1305
         Width           =   600
      End
      Begin VB.VScrollBar vsDiscosV 
         Height          =   330
         LargeChange     =   10
         Left            =   6240
         Max             =   1
         Min             =   6
         TabIndex        =   28
         Top             =   1320
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsDiscosH 
         Height          =   330
         LargeChange     =   10
         Left            =   6240
         Max             =   1
         Min             =   6
         TabIndex        =   27
         Top             =   990
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtDiscosH 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   5640
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   990
         Width           =   600
      End
      Begin VB.CheckBox chkDistorcionarTapas 
         BackColor       =   &H00000000&
         Caption         =   "Distorsionar tapas de discos ocupando 100% pantalla."
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
         Left            =   60
         TabIndex        =   25
         Top             =   450
         Width           =   5115
      End
      Begin VB.CheckBox chkRotulosArriba 
         BackColor       =   &H00000000&
         Caption         =   "Colocar rótulos arriba de las tapas de discos."
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
         Left            =   60
         TabIndex        =   23
         Top             =   690
         Width           =   5355
      End
      Begin VB.CheckBox chkRankToPeople 
         BackColor       =   &H00000000&
         Caption         =   "Exponer el Ranking al público."
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
         Left            =   60
         TabIndex        =   22
         Top             =   210
         Width           =   5295
      End
      Begin tbrFaroButton.fBoton Command10 
         Height          =   375
         Left            =   5520
         TabIndex        =   220
         Top             =   2790
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   12632256
         fCapt           =   "Protector de pantalla"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command20 
         Height          =   375
         Left            =   5520
         TabIndex        =   221
         Top             =   3180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   12632256
         fCapt           =   "Publicidades"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton1 
         Height          =   375
         Left            =   5520
         TabIndex        =   225
         Top             =   3570
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   12632256
         fCapt           =   "Imagenes inicio Windows"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton2 
         Height          =   465
         Left            =   90
         TabIndex        =   226
         Top             =   4560
         Width           =   3105
         _ExtentX        =   5477
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   12632256
         fCapt           =   "Elegir / modificar SKIN"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton5 
         Height          =   465
         Left            =   3240
         TabIndex        =   263
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   820
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Otros textos"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00533422&
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
         ForeColor       =   &H00FFFFC0&
         Height          =   255
         Left            =   60
         TabIndex        =   134
         Top             =   4230
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   540
         TabIndex        =   49
         Top             =   2310
         Width           =   2205
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discos-Vertical"
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
         Height          =   285
         Index           =   15
         Left            =   6630
         TabIndex        =   31
         Top             =   1350
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Discos-Horizontal"
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
         Height          =   285
         Index           =   16
         Left            =   6630
         TabIndex        =   30
         Top             =   1050
         Width           =   3345
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
      Height          =   2415
      Left            =   3150
      TabIndex        =   32
      Top             =   6930
      Visible         =   0   'False
      Width           =   4185
      Begin VB.OptionButton chkProtectAvance 
         BackColor       =   &H00000000&
         Caption         =   "Avance de discos automático"
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
         Left            =   30
         TabIndex        =   258
         Top             =   1230
         Width           =   4100
      End
      Begin VB.OptionButton chkProtectOriginal 
         BackColor       =   &H00000000&
         Caption         =   "Usar Protector de pantalla original (tapas de los discos)."
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
         Height          =   405
         Left            =   30
         TabIndex        =   60
         Top             =   750
         Width           =   4100
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
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   30
         TabIndex        =   59
         Top             =   510
         Width           =   4100
      End
      Begin VB.OptionButton chkNoProtector 
         BackColor       =   &H00000000&
         Caption         =   "No usar protector de pantalla."
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
         Height          =   225
         Left            =   30
         TabIndex        =   58
         Top             =   240
         Width           =   4100
      End
      Begin VB.VScrollBar vsEsperaTecla 
         Height          =   330
         LargeChange     =   5
         Left            =   3750
         Max             =   5
         Min             =   1200
         SmallChange     =   5
         TabIndex        =   61
         Top             =   1620
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
         Left            =   3120
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1620
         Width           =   600
      End
      Begin VB.VScrollBar vsDuracionProtect 
         Height          =   330
         LargeChange     =   10
         Left            =   3750
         Max             =   0
         Min             =   900
         SmallChange     =   10
         TabIndex        =   62
         Top             =   1980
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
         Left            =   3120
         TabIndex        =   34
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1980
         Width           =   600
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   270
         X2              =   3840
         Y1              =   1530
         Y2              =   1530
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
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   7
         Left            =   180
         TabIndex        =   37
         Top             =   1680
         Width           =   2925
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Duración del protector"
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
         Height          =   285
         Index           =   17
         Left            =   150
         TabIndex        =   36
         Top             =   2040
         Width           =   2925
      End
   End
   Begin VB.Frame frCreditos 
      BackColor       =   &H00000000&
      Caption         =   "Créditos"
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
      Height          =   5085
      Left            =   3090
      TabIndex        =   146
      Top             =   210
      Visible         =   0   'False
      Width           =   8685
      Begin VB.VScrollBar vsCreditosXaVipMusica 
         Height          =   330
         LargeChange     =   10
         Left            =   3510
         Max             =   0
         Min             =   100
         TabIndex        =   249
         Top             =   4230
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   3990
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   248
         Text            =   "frmConfig.frx":BFE9
         Top             =   3540
         Width           =   4575
      End
      Begin VB.TextBox txtCreditosXaVipMusica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2790
         TabIndex        =   246
         TabStop         =   0   'False
         Text            =   "5"
         Top             =   4230
         Width           =   750
      End
      Begin VB.TextBox txtPesosVIPMusica 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2790
         TabIndex        =   245
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   4590
         Width           =   1050
      End
      Begin tbrFaroButton.fBoton command3 
         Height          =   435
         Left            =   6990
         TabIndex        =   231
         Top             =   210
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "en cero"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.ComboBox cmbSCM 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmConfig.frx":C16A
         Left            =   6330
         List            =   "frmConfig.frx":C174
         Style           =   2  'Dropdown List
         TabIndex        =   188
         Top             =   1020
         Width           =   2205
      End
      Begin VB.TextBox txtPrecioBase2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5400
         TabIndex        =   171
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   1140
         Width           =   810
      End
      Begin VB.TextBox txtExplicPrecios 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   6360
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   170
         Top             =   1770
         Width           =   2235
      End
      Begin VB.TextBox txtPrecioBASE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5400
         TabIndex        =   169
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
         TabIndex        =   168
         Top             =   1140
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosBilletes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   167
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1140
         Width           =   600
      End
      Begin VB.TextBox txtCreditosCuestaTema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2010
         TabIndex        =   166
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2985
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   2
         LargeChange     =   10
         Left            =   2610
         Max             =   0
         Min             =   100
         TabIndex        =   165
         Top             =   2985
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTemaVIDEO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4260
         TabIndex        =   164
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2985
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   2
         LargeChange     =   10
         Left            =   4860
         Max             =   0
         Min             =   100
         TabIndex        =   163
         Top             =   2985
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3000
         TabIndex        =   162
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2985
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5220
         TabIndex        =   161
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2985
         Width           =   1100
      End
      Begin VB.TextBox txtCreditosCuestaTema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2010
         TabIndex        =   160
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2625
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   1
         LargeChange     =   10
         Left            =   2610
         Max             =   0
         Min             =   100
         TabIndex        =   159
         Top             =   2625
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTemaVIDEO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4260
         TabIndex        =   158
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2625
         Width           =   600
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   1
         LargeChange     =   10
         Left            =   4860
         Max             =   0
         Min             =   100
         TabIndex        =   157
         Top             =   2625
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtPrecioM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3000
         TabIndex        =   156
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2625
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5220
         TabIndex        =   155
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2625
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioV 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   5220
         TabIndex        =   154
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2265
         Width           =   1100
      End
      Begin VB.TextBox txtPrecioM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   3000
         TabIndex        =   153
         TabStop         =   0   'False
         Text            =   "8,88"
         Top             =   2265
         Width           =   1100
      End
      Begin VB.VScrollBar vsCreditosCuestaTemaVIDEO 
         Height          =   330
         Index           =   0
         LargeChange     =   10
         Left            =   4860
         Max             =   0
         Min             =   100
         TabIndex        =   152
         Top             =   2265
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTemaVIDEO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   4260
         TabIndex        =   151
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2265
         Width           =   600
      End
      Begin VB.TextBox txtTemasXCredito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         TabIndex        =   150
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
         TabIndex        =   149
         Top             =   780
         Value           =   1
         Width           =   330
      End
      Begin VB.VScrollBar vsCreditosCuestaTema 
         Height          =   330
         Index           =   0
         LargeChange     =   10
         Left            =   2610
         Max             =   0
         Min             =   100
         TabIndex        =   148
         Top             =   2265
         Value           =   1
         Width           =   330
      End
      Begin VB.TextBox txtCreditosCuestaTema 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Left            =   2010
         TabIndex        =   147
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2265
         Width           =   600
      End
      Begin tbrFaroButton.fBoton fBoton3 
         Height          =   975
         Left            =   420
         TabIndex        =   250
         Top             =   3750
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   1720
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Poner en cero creditos actuales."
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         X1              =   2250
         X2              =   2250
         Y1              =   3450
         Y2              =   5070
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos para selecciones VIP"
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
         Height          =   765
         Index           =   19
         Left            =   2280
         TabIndex        =   247
         Top             =   3510
         Width           =   1725
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   150
         X2              =   8460
         Y1              =   3420
         Y2              =   3420
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   150
         X2              =   8460
         Y1              =   1590
         Y2              =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mostar créditos como"
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
         Height          =   165
         Index           =   45
         Left            =   6060
         TabIndex        =   189
         Top             =   810
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   42
         Left            =   4920
         TabIndex        =   184
         Top             =   1140
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X1"
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
         Height          =   315
         Index           =   54
         Left            =   1770
         TabIndex        =   183
         Top             =   2310
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X3"
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
         Height          =   315
         Index           =   46
         Left            =   1770
         TabIndex        =   182
         Top             =   3060
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "En cero X1= modo gratuito. En cero X2 o X3= no usa promociones."
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
         Height          =   1725
         Index           =   53
         Left            =   60
         TabIndex        =   181
         Top             =   1680
         Width           =   1545
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   52
         Left            =   4920
         TabIndex        =   180
         Top             =   780
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Créditos por cada señal de billetero (S)"
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
         Height          =   285
         Index           =   49
         Left            =   90
         TabIndex        =   179
         Top             =   1170
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "X2"
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
         Height          =   315
         Index           =   43
         Left            =   1770
         TabIndex        =   178
         Top             =   2700
         Width           =   255
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00E0E0E0&
         X1              =   120
         X2              =   6150
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Label lblContador2 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20264536538"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   1680
         TabIndex        =   177
         Top             =   270
         Width           =   2600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Contador histórico/Interno"
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
         Height          =   405
         Index           =   39
         Left            =   -210
         TabIndex        =   176
         Top             =   240
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Créditos para VIDEO/KARAOKE"
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
         Height          =   465
         Index           =   28
         Left            =   4260
         TabIndex        =   175
         Top             =   1770
         Width           =   2055
      End
      Begin VB.Label lblContador 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "20264536538"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   4320
         TabIndex        =   174
         Top             =   270
         Width           =   2595
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Créditos por cada señal de monedero (Q)"
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
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   173
         Top             =   810
         Width           =   3885
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Créditos para música"
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
         Height          =   465
         Index           =   26
         Left            =   2010
         TabIndex        =   172
         Top             =   1770
         Width           =   2055
      End
   End
   Begin VB.TextBox txtPO 
      Height          =   345
      Left            =   60
      TabIndex        =   244
      Top             =   1830
      Visible         =   0   'False
      Width           =   2805
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
      Height          =   3735
      Left            =   150
      TabIndex        =   41
      Top             =   9570
      Visible         =   0   'False
      Width           =   7875
      Begin tbrFaroButton.fBoton fBoton2 
         Height          =   645
         Left            =   5130
         TabIndex        =   243
         Top             =   2880
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1138
         fFColor         =   6553600
         fBColor         =   16761024
         fCapt           =   "Opciones internas del reproductor de video"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin VB.TextBox txtSegFadeB 
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   6840
         TabIndex        =   214
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2070
         Width           =   600
      End
      Begin VB.VScrollBar vsSegFadeB 
         Height          =   330
         Left            =   7470
         Max             =   1
         Min             =   10
         TabIndex        =   213
         Top             =   2070
         Value           =   10
         Width           =   330
      End
      Begin VB.VScrollBar vsSegFade 
         Height          =   330
         Left            =   7470
         Max             =   1
         Min             =   10
         TabIndex        =   143
         Top             =   1320
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   6840
         TabIndex        =   142
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1320
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
         ForeColor       =   &H00E0E0E0&
         Height          =   615
         Left            =   60
         TabIndex        =   139
         Top             =   3000
         Width           =   3960
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3240
         TabIndex        =   137
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   2190
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.VScrollBar vsCortaMusicaPaga 
         Height          =   330
         LargeChange     =   10
         Left            =   3870
         Max             =   10
         Min             =   100
         SmallChange     =   10
         TabIndex        =   136
         Top             =   2220
         Value           =   10
         Visible         =   0   'False
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
         ForeColor       =   &H00404040&
         Height          =   360
         ItemData        =   "frmConfig.frx":C190
         Left            =   1530
         List            =   "frmConfig.frx":C192
         Style           =   2  'Dropdown List
         TabIndex        =   76
         Top             =   2640
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3240
         TabIndex        =   53
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1500
         Width           =   600
      End
      Begin VB.VScrollBar VSSegEspera 
         Height          =   330
         LargeChange     =   10
         Left            =   3870
         Max             =   0
         Min             =   7200
         SmallChange     =   10
         TabIndex        =   52
         Top             =   1500
         Value           =   30
         Width           =   330
      End
      Begin VB.VScrollBar VsPorcTema 
         Height          =   330
         LargeChange     =   10
         Left            =   3870
         Max             =   10
         Min             =   100
         SmallChange     =   10
         TabIndex        =   51
         Top             =   1875
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3240
         TabIndex        =   50
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1860
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
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   3240
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1140
         Width           =   600
      End
      Begin VB.VScrollBar VSmaxFichas 
         Height          =   330
         Left            =   3870
         Max             =   0
         Min             =   6800
         TabIndex        =   46
         Top             =   1140
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
         Height          =   855
         Left            =   180
         TabIndex        =   42
         Top             =   240
         Width           =   7575
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
            ForeColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   60
            TabIndex        =   44
            Top             =   510
            Value           =   -1  'True
            Width           =   7440
         End
         Begin VB.OptionButton OpReiniFull 
            BackColor       =   &H00000000&
            Caption         =   "Se ejecutan todas las canciones pendientes en lista de ejecución."
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
            Height          =   285
            Left            =   60
            TabIndex        =   43
            Top             =   240
            Width           =   7485
         End
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tiempo de fade in / fade out al cancelar canciones"
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
         Height          =   645
         Index           =   56
         Left            =   4980
         TabIndex        =   215
         Top             =   1920
         Width           =   1755
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
         ForeColor       =   &H00E0E0E0&
         Height          =   645
         Index           =   25
         Left            =   4800
         TabIndex        =   144
         Top             =   1170
         Width           =   2025
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
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   40
         Left            =   90
         TabIndex        =   138
         Top             =   2280
         Visible         =   0   'False
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
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   27
         Left            =   360
         TabIndex        =   77
         Top             =   2700
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje ejecutar canción"
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
         Height          =   285
         Index           =   12
         Left            =   90
         TabIndex        =   54
         Top             =   1920
         Width           =   3075
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Espera autoejecutar canción"
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
         Height          =   285
         Index           =   9
         Left            =   210
         TabIndex        =   55
         Top             =   1530
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Máximo credito (0=no limite)"
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
         Height          =   240
         Index           =   8
         Left            =   180
         TabIndex        =   45
         Top             =   1200
         Width           =   2925
      End
   End
   Begin VB.Frame frAceleracion 
      BackColor       =   &H00000000&
      Caption         =   "Aceleración de 3PM"
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
      Height          =   1905
      Left            =   12180
      TabIndex        =   38
      Top             =   3600
      Visible         =   0   'False
      Width           =   7095
      Begin VB.CheckBox chkLoadTapaIni 
         BackColor       =   &H00000000&
         Caption         =   "Cargar todas las imagenes de los discos al iniciar."
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
         Height          =   420
         Left            =   150
         TabIndex        =   241
         Top             =   1440
         Width           =   6585
      End
      Begin VB.TextBox txtTamanoTapaPermitido 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
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
         Left            =   480
         TabIndex        =   191
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   1050
         Width           =   480
      End
      Begin VB.VScrollBar vsTamanoTapaPermitido 
         Height          =   330
         LargeChange     =   10
         Left            =   150
         Max             =   20
         Min             =   200
         SmallChange     =   10
         TabIndex        =   190
         Top             =   1050
         Value           =   200
         Width           =   330
      End
      Begin VB.CheckBox chkVUMeter 
         BackColor       =   &H00000000&
         Caption         =   "Habilitar Vumetro (consume procesador, no usar en equipos limitados)."
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
         Height          =   420
         Left            =   150
         TabIndex        =   40
         Top             =   540
         Width           =   6585
      End
      Begin VB.CheckBox chkCargarDuracionTemas 
         BackColor       =   &H00000000&
         Caption         =   "Cargar duración de canciones(demora extra)."
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
         Height          =   225
         Left            =   150
         TabIndex        =   39
         Top             =   270
         Width           =   5890
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tamaño máximo en KB permitido para portadas."
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
         Height          =   225
         Index           =   47
         Left            =   1020
         TabIndex        =   192
         Top             =   1110
         Width           =   4485
      End
   End
   Begin tbrFaroButton.fBoton Command19 
      Height          =   435
      Left            =   9150
      TabIndex        =   239
      Top             =   8490
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Comprar ahora"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command21 
      Height          =   435
      Left            =   7980
      TabIndex        =   238
      Top             =   8490
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Cluff"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Frame frIMGWIN 
      BackColor       =   &H00000000&
      Caption         =   "Imágenes inicio Windows (solo 98-Me)"
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
      Height          =   3435
      Left            =   12120
      TabIndex        =   223
      Top             =   60
      Visible         =   0   'False
      Width           =   7515
      Begin tbrFaroButton.fBoton cmdImg 
         Height          =   375
         Index           =   1
         Left            =   150
         TabIndex        =   232
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "cambiar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton cmdImg 
         Height          =   375
         Index           =   2
         Left            =   2580
         TabIndex        =   233
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "cambiar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton cmdImg 
         Height          =   375
         Index           =   3
         Left            =   5010
         TabIndex        =   234
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "cambiar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton cmdImgQ 
         Height          =   375
         Index           =   1
         Left            =   1260
         TabIndex        =   235
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton cmdImgQ 
         Height          =   375
         Index           =   2
         Left            =   3690
         TabIndex        =   236
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton cmdImgQ 
         Height          =   375
         Index           =   3
         Left            =   6120
         TabIndex        =   237
         Top             =   2940
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "quitar"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmConfig.frx":C194
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
         Height          =   675
         Index           =   31
         Left            =   90
         TabIndex        =   224
         Top             =   240
         Width           =   7215
      End
      Begin VB.Image img1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1995
         Left            =   90
         Stretch         =   -1  'True
         Top             =   930
         Width           =   2400
      End
      Begin VB.Image img2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1995
         Left            =   2490
         Stretch         =   -1  'True
         Top             =   930
         Width           =   2400
      End
      Begin VB.Image img3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1995
         Left            =   4920
         Stretch         =   -1  'True
         Top             =   930
         Width           =   2400
      End
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   450
      Left            =   120
      TabIndex        =   218
      Top             =   8430
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   794
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   660
      Left            =   120
      TabIndex        =   217
      Top             =   7740
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1164
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Frame Frame7 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   3765
      Left            =   60
      TabIndex        =   216
      Top             =   3750
      Width           =   2835
      Begin tbrFaroButton.fBoton Command12 
         Height          =   375
         Left            =   90
         TabIndex        =   8
         Top             =   210
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Créditos"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command13 
         Height          =   375
         Left            =   90
         TabIndex        =   9
         Top             =   600
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Teclado"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command6 
         Height          =   375
         Left            =   90
         TabIndex        =   14
         Top             =   2550
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Inicio 3PM"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command5 
         Height          =   375
         Left            =   90
         TabIndex        =   12
         Top             =   1770
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Karaokes"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command4 
         Height          =   375
         Left            =   90
         TabIndex        =   10
         Top             =   990
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Administrar discos"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command26 
         Height          =   375
         Left            =   90
         TabIndex        =   13
         Top             =   2160
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Importar/Exportar config"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command17 
         Height          =   375
         Left            =   90
         TabIndex        =   11
         Top             =   1380
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Validación de uso"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton fBoton1 
         Height          =   375
         Left            =   90
         TabIndex        =   15
         Top             =   2940
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Vender música"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command9 
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   3330
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Claves 3PM"
         fEnabled        =   0   'False
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Básicas"
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
      Height          =   1785
      Left            =   30
      TabIndex        =   56
      Top             =   0
      Width           =   2865
      Begin tbrFaroButton.fBoton Command11 
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Visualización"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command15 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   570
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   16776960
         fCapt           =   "Aceleración de 3PM"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   0
      End
      Begin tbrFaroButton.fBoton Command14 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Otras opciones"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command7 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1350
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Abrir MANUAL"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1515
      Left            =   30
      TabIndex        =   140
      Top             =   2190
      Width           =   2865
      Begin VB.TextBox txtClaveAdmin 
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   1020
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   660
         Width           =   1635
      End
      Begin tbrFaroButton.fBoton Command27 
         Height          =   435
         Left            =   150
         TabIndex        =   5
         Top             =   210
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Cambiar/Crear Clave"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command16 
         Height          =   375
         Left            =   2310
         TabIndex        =   6
         Top             =   210
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "?"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton Command31 
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1020
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Ingreso Administrador"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese clave"
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
         Index           =   41
         Left            =   120
         TabIndex        =   141
         Top             =   630
         Width           =   915
      End
   End
   Begin VB.HScrollBar HSvolumen 
      Height          =   240
      LargeChange     =   10
      Left            =   9000
      Max             =   100
      TabIndex        =   131
      Top             =   5610
      Width           =   2895
   End
   Begin VB.HScrollBar HSVolumen2 
      Height          =   240
      LargeChange     =   10
      Left            =   9000
      Max             =   100
      TabIndex        =   130
      Top             =   5940
      Width           =   2895
   End
   Begin VB.Frame frConfigVis 
      BackColor       =   &H00000000&
      Caption         =   "Opcion elegida"
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
      Height          =   5445
      Left            =   2940
      TabIndex        =   242
      Top             =   0
      Width           =   9015
   End
   Begin VB.Label lblHLP 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Detalle/Ayuda de la opción elegida."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3345
      Left            =   2970
      TabIndex        =   19
      Top             =   5520
      Width           =   4575
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   7800
      TabIndex        =   133
      Top             =   5640
      Width           =   1140
   End
   Begin VB.Line LineScroll 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      X1              =   9000
      X2              =   11850
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line LineScroll2 
      BorderColor     =   &H00808080&
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
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   7560
      TabIndex        =   132
      Top             =   5910
      Width           =   1380
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   12000
      X2              =   12000
      Y1              =   0
      Y2              =   9000
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Visible         =   0   'False
      X1              =   0
      X2              =   12000
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label lblTBRcfg 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmConfig.frx":C232
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   2085
      Left            =   8010
      TabIndex        =   20
      Top             =   6360
      Width           =   2805
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TeclaConfOK As String
Dim TeclaConfESC As String

Private Color1 As Long
Private Color2 As Long
Private Color3 As Long
Private Color4 As Long
Private Color5 As Long
Private Color6 As Long

Public Sub SendW()
    Form_KeyDown TeclaCerrarSistema, 0
End Sub

Private Sub chkActivarERROR_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    HLP TR.Trad("Active solo en caso de que 3PM se " + _
        "ciere bruscamente con errores. NO ACTIVAR SI 3PM " + _
        "FUNCIONA CORRECTAMENTE. Luego de activar reinicie 3PM y luego " + _
        "de que se cierre con fallo busque en la carpeta de 3PM todos " + _
        "los archivos 'REG*****.W15' y envíelos a tbrsoft (info@tbrsoft.com) " + _
        "detallando el mensaje de error que informa 3PM antes de cerrarse. " + vbCrLf + _
        "Luego de esto recibira un email con el detalle de su error y la " + _
        "solución correspondiente%98%Referencia " + _
        "sobre como enviar a tbrSoft datos de error%99%")
End Sub

Private Sub chkApagarPC_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkApagarPC.ForeColor = Color5
    HLP TR.Trad("Tecla de cierre de 3PM. Si esta habilitado el " + _
        "apagado al cerrar 3PM Windows se cerrará tambien. Este cambio es " + _
        "automatico, no necesita reiniciar 3PM%99%")
End Sub

Private Sub chkApagarPC_LostFocus()
    chkApagarPC.ForeColor = Color6
End Sub

Private Sub chkBloquearMusicaElegida_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkBloquearMusicaElegida.ForeColor = Color5
    HLP TR.Trad("Si activa esta opción cuando ingrese a algún disco y seleccione " + _
        "algún tema este quedará bloqueado hasta que vuelva a abrir el " + _
        "disco. Esto evita la seleccion multiple de una misma selección varias " + _
        "veces continuadas%99%")
End Sub

Private Sub chkBloquearMusicaElegida_LostFocus()
    chkBloquearMusicaElegida.ForeColor = Color6
End Sub

Private Sub chkCargarDuracionTemas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkCargarDuracionTemas.ForeColor = Color5
    HLP TR.Trad("Cada vez que se habra un disco se pueden mostrar las " + _
        "duraciones de los temas. No se recomienda habilitar esta funcion " + _
        "salvo que cuente con un equipo potente%99%")
End Sub

Private Sub chkCargarDuracionTemas_LostFocus()
    chkCargarDuracionTemas.ForeColor = Color6
End Sub

Private Sub chkCS_Click()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkApagarPC.ForeColor = Color5
    HLP TR.Trad("Le permite corregir errores en la recepcion de las " + _
        "señales de su monedero / billetero electrónico. No lo active " + _
        "si no es muy necesario%99%")
End Sub

Private Sub chkDistorcionarTapas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkDistorcionarTapas.ForeColor = Color5
    HLP TR.Trad("Si habilita esta opcion las fotos " + _
        "se distorsionaran para ocupar todo el espacio disponible. Caso " + _
        "contrario se dejara el espacio sobrante como libre. Este cambio " + _
        "solo se vera una vez reiniciado 3PM%99%")
End Sub

Private Sub chkDistorcionarTapas_LostFocus()
    chkDistorcionarTapas.ForeColor = Color6
End Sub

Private Sub chkLoadTapaIni_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkLoadTapaIni.ForeColor = Color5
    HLP TR.Trad("Si las imágenes se cargan al inicio tendrá un consumo en memoria " + _
        " importante pero el paso de páginas será notoriamente rápido. Si no cuenta " + _
        "con muchos recursos será mejor que desactive esta opción haciendo mucho " + _
        "menor el consumo de memoria y demorando un poco el paso de páginas" + _
        "%99%")
End Sub

Private Sub chkLoadTapaIni_LostFocus()
    chkLoadTapaIni.ForeColor = Color6
End Sub

Private Sub chkMostrarRotulos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkMostrarRotulos.ForeColor = Color5
    HLP TR.Trad("Se recomienda dejar esta opción habilitada, ya que " + _
        "el usuario final debera identificar un disco solo por " + _
        "su tapa (no estara disponible el nombre del interprete y el " + _
        "nombre del disco). Este cambio solo se vera una vez reiniciado 3PM%99%")
End Sub

Private Sub chkMostrarRotulos_LostFocus()
    chkMostrarRotulos.ForeColor = Color6
End Sub

Private Sub chknoprotector_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkNoProtector.ForeColor = Color5
    HLP TR.Trad("Deshabilitar la función de protección de pantalla. " + _
        "No recomendado%99%")
End Sub

Private Sub chknoprotector_LostFocus()
    chkNoProtector.ForeColor = Color6
End Sub

Private Sub chkNoVumVID_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkNoVumVID.ForeColor = Color5
    HLP TR.Trad("Quitar el VUMetro (medidor de sonido) cuando los " + _
        "videos sean full-screen%99%")
End Sub

Private Sub chkNoVumVID_LostFocus()
    chkNoVumVID.ForeColor = Color6
End Sub

Private Sub chkOutTemasWhenSel_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkOutTemasWhenSel.ForeColor = Color5
    HLP TR.Trad("Salir inmediatamente del listado de musica al " + _
        "hacer una selección%99%")
End Sub

Private Sub chkOutTemasWhenSel_LostFocus()
    chkOutTemasWhenSel.ForeColor = Color6
End Sub

Private Sub chkPasarhoja_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkPasarhoja.ForeColor = Color5
    HLP TR.Trad("Habilitar a las teclas de desplazamiento simple " + _
        "para pasar páginas. Si esta inhabilitado al llegar al ultimo " + _
        "disco de una página volvera al primero disco de la misma " + _
        "(y viceversa). Este cambio es automatico, no necesita reiniciar 3PM%99%")
End Sub

Private Sub chkPasarhoja_LostFocus()
    chkPasarhoja.ForeColor = Color6
End Sub

Private Sub chkProtectAvance_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectAvance.ForeColor = Color5
    HLP TR.Trad("La selección de discos comienza a avanzar automáticamente.%99%")
End Sub

Private Sub chkProtectAvance_LostFocus()
    chkProtectAvance.ForeColor = Color6
End Sub

Private Sub chkProtectorCustom_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectorCustom.ForeColor = Color5
    HLP TR.Trad("Si desea mostrar imagenes personalizadas debera cargarlas " + _
        "en la carpeta FOTOS dentro de la carpeta en que se instalo 3PM. No use " + _
        "imagenes muy pesadas ya que puede afectar el rendimiento de 3PM. " + _
        "Se recomienda no sobrepasar los 100 KB%99%")
End Sub

Private Sub chkProtectorCustom_LostFocus()
    chkProtectorCustom.ForeColor = Color6
End Sub

Private Sub chkProtectOriginal_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkProtectOriginal.ForeColor = Color5
    HLP TR.Trad("Puede usar para proteger la pantalla el protector " + _
        "por defecto. Este muestra las tapas de los discos.%99%")
End Sub

Private Sub chkProtectOriginal_LostFocus()
    chkProtectOriginal.ForeColor = Color6
End Sub

Private Sub chkQuitaBarraInf_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkQuitaBarraInf.ForeColor = Color5
    HLP TR.Trad("Ampliar el espacio asigado a discos achicando la barra inferior de información%99%")
End Sub

Private Sub chkQuitaBarraInf_LostFocus()
    chkQuitaBarraInf.ForeColor = Color6
End Sub

Private Sub chkQuitaBarraSup_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkQuitaBarraSup.ForeColor = Color5
    HLP TR.Trad("Ampliar el espacio asigado a discos quitado la indicación de ritmo y letra elegidos%99%")
End Sub

Private Sub chkQuitaBarraSup_LostFocus()
    chkQuitaBarraSup.ForeColor = Color6
End Sub

Private Sub chkRankToPeople_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkRankToPeople.ForeColor = Color5
    HLP TR.Trad("3PM acumula los totales de ejecuciones de cada " + _
        "selección. Esto está " + _
        "ordenado, es consultable y puede mostrarse o no a los " + _
        "usuarios finales. Si se muestra permite tambien cargar temas " + _
        "desde aquí como un disco más evitando la busqueda de discos. " + _
        "Se recomienda dejar " + _
        "activado. Este cambio solo se verá una vez reiniciado 3PM%99%")
End Sub

Private Sub chkRankToPeople_LostFocus()
    chkRankToPeople.ForeColor = Color6
End Sub

Private Sub chkRotulosArriba_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkRotulosArriba.ForeColor = Color5
    HLP TR.Trad("Se dice rótulo al indicador del nombre de cada disco. " + _
        "Esta opción sirve para colocarlo encima de la foto. Si deshabilita " + _
        "esta opcion el rotulo aparecerá por debajo de la foto (valor " + _
        "recomendado). Este cambio solo se verá una vez reiniciado 3PM%99%")
End Sub

Private Sub chkRotulosArriba_LostFocus()
    chkRotulosArriba.ForeColor = Color6
End Sub

Private Sub chkS3_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkS3.ForeColor = Color5
    HLP TR.Trad("He adquirirdo la interfase de comunicaciones de " + _
        "3PM y deseo comenzar a escuchas sus señales%98%tbrSoft " + _
        "desarrollo un dispositivo que se comunica con la PC. Esta opción " + _
        "permite activarlo para su uso%99%")
End Sub

Private Sub chkS3_LostFocus()
    chkS3.ForeColor = Color6
End Sub

Private Sub chkSalida2_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkSalida2.ForeColor = Color5
    HLP TR.Trad("Habilitar la segunda salida para reproduccion de " + _
        "videos. Debe habilitarse la salida de TV como expansión del " + _
        "escritorio y configurarla con la misma definición de pixeles " + _
        "para ambas salidas%99%")
End Sub

Private Sub chkSalida2_LostFocus()
    chkSalida2.ForeColor = Color6
End Sub

Private Sub chkTouch_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkTouch.ForeColor = Color5
    HLP TR.Trad("Mostrar los botones para pantallas sensibles " + _
        "al tacto. Este cambio solo se verá una vez reiniciado 3PM%99%")
End Sub

Private Sub chkTouch_LostFocus()
    chkTouch.ForeColor = Color6
End Sub

Private Sub chkUseAPITecla_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkApagarPC.ForeColor = Color5
    HLP TR.Trad("Cambia la recepción de las teclas de monedero " + _
        "directamente desde al hardware de teclado%99%")
End Sub

Private Sub chkUseAPITecla_LostFocus()
    chkUseAPITecla.ForeColor = Color6
End Sub

Private Sub chkVidFullScreen_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVidFullScreen.ForeColor = Color5
    HLP TR.Trad("Mostrar los videos en pantalla completa cuando se ejecuten%99%")
End Sub

Private Sub chkVidFullScreen_LostFocus()
    chkVidFullScreen.ForeColor = Color6
End Sub

Private Sub chkVidMudos_Click()
    
    If PUBs.TotalPUBsMUTE = 0 Then
        MsgBox TR.Trad("No puede activar esta opción ya que no hay " + _
            "publicidades cargadas." + vbCrLf + _
            "Para cargar publicidades debera incluir en la " + _
            "carpeta 'PUBMUTE' (dentro de la carpeta en que instalo 3PM) uno " + _
            "o más ficheros AVI, MPG, DAT (VCD) o VOB (DVD)%99%")
        chkVidMudos = 0
    End If

End Sub

Private Sub chkVidMudos_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVidMudos.ForeColor = Color5
    HLP TR.Trad("Indica si se reproducirán publicidades por la salida " + _
        "de TV sin sonido. Esto no interrumpe ninguna otra reproducción " + _
        "de la rockola. Si se habilita esta opción deben colocarse ficheros " + _
        "de video AVI, MPG, VOB (DVD) o DAT (VCD) en la carpeta PUBMUTE " + _
        "(de la carpeta en la que instalo 3PM). Estos ficheros " + _
        "se reproducen continuamente salvo que algún usuario cargue algun video pago. " + _
        "Se reproducen en orden alfabético por lo que podrá modificar " + _
        "el nombre para definir el orden deseado. Habilitar esta opcion " + _
        "anulas las imagenes publictarias destinadas al tv%99%")
End Sub

Private Sub chkVidMudos_LostFocus()
    chkVidMudos.ForeColor = Color6
End Sub

Private Sub chkVUMeter_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    chkVUMeter.ForeColor = Color5
    HLP TR.Trad("Se llama VuMetro al medidor de nivel de sonido. Este " + _
        "es muy atractivo a la vista pero consume recursos de la PC. " + _
        "Por esto solo deberá usarse cuando el rendimiento del equipo " + _
        "no se vea afectado con el uso de este. Para PCs de bajos " + _
        "recursos (procesador y RAM) se recomienda dejar desactivado. " + _
        "Este cambio solo se vera una vez reiniciado 3PM%99%")
End Sub

Private Sub chkVUMeter_LostFocus()
    chkVUMeter.ForeColor = Color6
End Sub

Private Sub ckPUB_Click()
    If PUBs.TotalPUBs = 0 Then
        MsgBox TR.Trad("No puede activar esta opción ya que no hay " + _
            "publicidades cargadas." + vbCrLf + _
            "Para cargar publicidades deberá incluir en la carpeta 'PUB' " + _
            "(en la carpeta en que instalo 3PM) uno o más ficheros " + _
            "MP3, WMA, AVI, MPG, VOB (DVD) o DAT (VCD)%99%")
        ckPUB = 0
    End If
End Sub

Private Sub ckPUB_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    ckPUB.ForeColor = Color5
    HLP TR.Trad("Indica si se reproducirán publicidades. Si se habilita " + _
        "esta opción deben colocarse ficheros MP3, WMA, AVI, MPG, VOB " + _
        "(DVD) o DAT (VCD) en la carpeta PUB (de la carpeta en la que " + _
        "instalo 3PM). Estos ficheros se reproducen cada X (a configurar) " + _
        "temas y de a uno por vez. Se reproducen en orden alfabético por " + _
        "lo que podrá modificar el nombre para definir el orden deseado. " + _
        "Puede tambien duplicar ficheros para darle mayor repetición a " + _
        "alguna publicidad en particular%99%")
End Sub

Private Sub ckPUB_LostFocus()
    ckPUB.ForeColor = Color6
End Sub

Private Sub ckPubIMG_Click()
    If ckPubIMG Then
        If PUBs.TotalPUBsIMG = 0 Then
            MsgBox TR.Trad("No puede activar esta opción ya que no hay " + _
                "publicidades (de menos de 50KB) cargadas." + vbCrLf + _
                "Para cargar publicidades debera incluir en la carpeta " + _
                "'PUB' (en la carpeta en que instalo 3PM) uno o más " + _
                "ficheros JPG, BMP o GIF. Debera reiniciar 3PM para " + _
                "que este cambio surta efecto%99%")
            ckPubIMG = 0
        End If
    End If
End Sub

Private Sub ckPubIMG_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    ckPubIMG.ForeColor = Color5
    HLP TR.Trad("Indica si se reproducirán publicidades. Si se habilita " + _
        "esta opción deben colocarse ficheros JPG, BMP o GIF en la " + _
        "carpeta PUB (de la carpeta en la que instalo 3PM). Estos ficheros " + _
        "se muestran cada X (a configurar) segundos. Se muestran en " + _
        "orden alfabético por lo que podrá modificar el nombre para " + _
        "definir el orden deseado. Puede tambien duplicar ficheros " + _
        "para darle mayor repeticion a alguna publicidad en particular%99%")
End Sub

Private Sub ckPubIMG_LostFocus()
    ckPubIMG.ForeColor = Color6
End Sub

Private Sub cmbIDIOMA_Click()
    TR.Language = (GetBasePath) + cmbIDIOMA + ".idm"
    Traducir
End Sub

Private Sub Command1_Click() 'GRABAR BUTTON
    On Error GoTo MiErr
    tERR.Anotar "aclp2"
    'GRABAR BUTTON
    'cargar los datos del archivo GPF("config")
    'paso todo a una cadena, la encripto y luego la escribo
    Dim FullConfig As String
    ChangeConfig "ClaveAdmin", ClaveAdmin
    ChangeConfig "TeclaDerecha", txtTeclas(0)
    ChangeConfig "TeclaIzquierda", txtTeclas(1)
    ChangeConfig "TeclaOK", txtTeclas(2)
    ChangeConfig "TeclaCancionVIP", txtTeclas(17)
    ChangeConfig "teclaSumValidar", txtTeclas(18)
    ChangeConfig "TeclaCarrito", txtTeclas(16)
    ChangeConfig "TeclaESC", txtTeclas(3)
    ChangeConfig "TeclaNuevaFicha", txtTeclas(4)
    ChangeConfig "TeclaNuevaFicha2", txtTeclas(15)
    ChangeConfig "TeclaConfig", txtTeclas(5)
    ChangeConfig "TeclaPagAd", txtTeclas(6)
    ChangeConfig "TeclaPagAt", txtTeclas(7)
    ChangeConfig "TeclaCerrarSistema", txtTeclas(8)
    tERR.Anotar "aclq"
    ChangeConfig "ShowCreditsMode", CStr(cmbSCM.ListIndex)
    ShowCreditsMode = cmbSCM.ListIndex
    ChangeConfig "TeclaShowContador", txtTeclas(9)
    ChangeConfig "TeclaPutCeroContador", txtTeclas(10)
    ChangeConfig "TeclaFF", txtTeclas(11)
    ChangeConfig "TeclaBajaVolumen", txtTeclas(12)
    ChangeConfig "TeclaSubeVolumen", txtTeclas(13)
    ChangeConfig "TeclaNextMusic", txtTeclas(14)
    tERR.Anotar "aclq2"
    ChangeConfig "plusparam", txtPO.tExt
    ChangeConfig "TeclaDerechax2", CStr(cmbTECLAS2(0).ListIndex)
    ChangeConfig "TeclaIzquierdax2", CStr(cmbTECLAS2(1).ListIndex)
    ChangeConfig "TeclaOKx2", CStr(cmbTECLAS2(2).ListIndex)
    ChangeConfig "TeclaCancionVIPx2", CStr(cmbTECLAS2(17).ListIndex)
    ChangeConfig "teclaSumValidarX2", CStr(cmbTECLAS2(18).ListIndex)
    ChangeConfig "TeclaCarritox2", CStr(cmbTECLAS2(16).ListIndex)
    ChangeConfig "TeclaESCx2", CStr(cmbTECLAS2(3).ListIndex)
    ChangeConfig "TeclaNuevaFichax2", CStr(cmbTECLAS2(4).ListIndex)
    ChangeConfig "TeclaNuevaFicha2x2", CStr(cmbTECLAS2(15).ListIndex)
    ChangeConfig "TeclaConfigx2", CStr(cmbTECLAS2(5).ListIndex)
    ChangeConfig "TeclaPagAdx2", CStr(cmbTECLAS2(6).ListIndex)
    ChangeConfig "TeclaPagAtx2", CStr(cmbTECLAS2(7).ListIndex)
    ChangeConfig "TeclaCerrarSistemax2", CStr(cmbTECLAS2(8).ListIndex)
    ChangeConfig "TeclaShowContadorx2", CStr(cmbTECLAS2(9).ListIndex)
    ChangeConfig "TeclaPutCeroContadorx2", CStr(cmbTECLAS2(10).ListIndex)
    ChangeConfig "TeclaFFx2", CStr(cmbTECLAS2(11).ListIndex)
    ChangeConfig "TeclaBajaVolumenx2", CStr(cmbTECLAS2(12).ListIndex)
    ChangeConfig "TeclaSubeVolumenx2", CStr(cmbTECLAS2(13).ListIndex)
    ChangeConfig "TeclaNextMusicx2", CStr(cmbTECLAS2(14).ListIndex)
    
    ChangeConfig "FrecTecladoTBR", CStr(vsFrecTecladoTBR)
    tERR.Anotar "aclq3"
    
    ChangeConfig "ActivarCorreccionSignal", CStr(chkCS)
    ChangeConfig "ApagarAlCierre", CStr(chkApagarPC)
    ChangeConfig "UseAPITecla", CStr(chkUseAPITecla)
    ChangeConfig "ActivarERR", CStr(chkActivarERROR)
    ChangeConfig "TamanoTapaPermitido", CStr(vsTamanoTapaPermitido)
    
    tERR.Anotar "aclr"
    If opModo4Teclas Then
        ChangeConfig "IsMod46Teclas", "46"
        IsMod46Teclas = 46
    End If
    If opModo5Teclas Then
        ChangeConfig "IsMod46Teclas", "5"
        IsMod46Teclas = 5
    End If
    ChangeConfig "RankToPeople", CStr(chkRankToPeople)
    ChangeConfig "MaximoFichas", txtMaxFichas
    ChangeConfig "EsperaMinutos", txtSECwait
    ChangeConfig "FastIni", CStr(chkFastINI)
    ChangeConfig "HabilitarVUMetro", CStr(chkVUMeter)
    ChangeConfig "LoadTapaIni", CStr(chkLoadTapaIni)
    
    
    ChangeConfig "QuitaBarraSup", CStr(chkQuitaBarraSup)
    ChangeConfig "QuitaBarraInf", CStr(chkQuitaBarraInf)
    ChangeConfig "VidfullScreen", CStr(chkVidFullScreen)
    tERR.Anotar "acls"
    ChangeConfig "Salida2", CStr(chkSalida2)
    ChangeConfig "NoVumVid", CStr(chkNoVumVID)
    ChangeConfig "OutTemasWhenSel", CStr(chkOutTemasWhenSel)
    ChangeConfig "BloquearMusicaElegida", CStr(chkBloquearMusicaElegida)
    'Valores de ReIni LISTA=solo lista NADA=arranca de cero
    If OpReiniFull Then
        ChangeConfig "ReINI", "LISTA"
    Else
        ChangeConfig "ReINI", "NADA"
    End If
    tERR.Anotar "aclt"
    ChangeConfig "Volumen", Trim(CStr(HSvolumen))
    ChangeConfig "Volumen2", Trim(CStr(HSVolumen2))
    ChangeConfig "EsperaTecla", txtEsperaTecla
    ChangeConfig "PorcentajeTema", txtPorcTema
    
    ChangeConfig "SegFade", txtSegFade
    SegFade = vsSegFade
    
    ChangeConfig "SegFadeB", txtSegFadeB
    SegFadeB = vsSegFadeB
    
    ChangeConfig "DiscosH", txtDiscosH
    ChangeConfig "DiscosV", txtDiscosV
    ChangeConfig "DuracionProtect", txtDuracionProtect
    tERR.Anotar "aclu"
    ChangeConfig "PasarHoja", CStr(chkPasarhoja)
    ChangeConfig "DistorcionarTapas", CStr(chkDistorcionarTapas)
    'valores para el protectore de pantalla
    '0=inhabilitado 1=Original 2=Carpeta Fotos 3= Video FullScreen
    tERR.Anotar "aclv"
    
    LCs3 = CStr(chkS3)
    ChangeConfig "UsarS3", CStr(chkS3)
    
    If chkNoProtector Then ChangeConfig "Protector", "0"
    If chkProtectOriginal Then ChangeConfig "Protector", "1"
    If chkProtectorCustom Then ChangeConfig "Protector", "2"
    If chkProtectAvance Then ChangeConfig "Protector", "3"
    
    ChangeConfig "CargarDuracionTemas", CStr(chkCargarDuracionTemas)
    ChangeConfig "MostrarRotulos", CStr(chkMostrarRotulos)
    ChangeConfig "RotulosArriba", CStr(chkRotulosArriba)
    ChangeConfig "TemasPorCredito", txtTemasXCredito
    ChangeConfig "CreditosBilletes", txtCreditosBilletes
    
    ChangeConfig "CreditosCuestaTema", txtCreditosCuestaTema(0)
    ChangeConfig "CreditosCuestaTema2", txtCreditosCuestaTema(1)
    ChangeConfig "CreditosCuestaTema3", txtCreditosCuestaTema(2)
    ChangeConfig "CreditosCuestaTemaVIDEO", txtCreditosCuestaTemaVIDEO(0)
    ChangeConfig "CreditosCuestaTemaVIDEO2", txtCreditosCuestaTemaVIDEO(1)
    ChangeConfig "CreditosCuestaTemaVIDEO3", txtCreditosCuestaTemaVIDEO(2)
    'upManu
    ChangeConfig "CreditosXaVipMusica", txtCreditosXaVipMusica
    
    ChangeConfig "PrecioBase", txtPrecioBASE
    ChangeConfig "PrecioBase2", txtPrecioBase2
    'si el idiota usa mas de un renglon estoy en problemas
    ChangeConfig "TextoUsuario", Replace(TxtUSUARIO, vbCrLf, Chr(5))
    
    ChangeConfig "MostrarTouch", CStr(chkTouch)
    tERR.Anotar "aclx"
    'publicidades
    ChangeConfig "MostrarPUB", CStr(ckPUB)
    ChangeConfig "MostrarPUBMute", CStr(chkVidMudos)
    ChangeConfig "MostrarPUBIMG", CStr(ckPubIMG)
    ChangeConfig "PubliCada", txtPubliCada
    ChangeConfig "PubliIMGCada", txtPubliImgCada
    ChangeConfig "Idioma", cmbIDIOMA
    tERR.Anotar "acly"
    
    'SI NO HAY que validar me aseguro que se borre el archivo de validacion sf + "radilav.cfg"
    If VALIDAR = False Then
        If fso.FileExists(GPF("radliv")) Then fso.DeleteFile GPF("radliv"), True
    End If
    
    tERR.Anotar "acma"
    'publicidades
    PUBs.SonarPublicidadesCada = Val(txtPubliCada)
    PUBs.HabilitarPublicidadesMp3Vid = ckPUB
    PUBs.HabilitarPublicidadesVMute = chkVidMudos
    
    PUBs.SonarPublicidadesIMGCada = Val(txtPubliImgCada)
    PUBs.HabilitarPublicidadesIMG = ckPubIMG
    
    IDIOMA = cmbIDIOMA
    TR.Language = (GetBasePath) + IDIOMA + ".idm"
    
    tERR.Anotar "acmb"
    
    'todas las propiedades se quedan sin reiniciar
    'algunas no se necesitan
   
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
    'upManu
    CreditosXaVipMusica = LeerConfig("CreditosXaVipMusica", "0") 'predeterminado desactivado
    
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "0")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "0")
    
    textoUsuario = LeerConfig("TextoUsuario", "Cargue los datos de su empresa aqui")
    textoUsuario = Replace(textoUsuario, Chr(5), vbCrLf)
    
    QuitaBarraInf = LeerConfig("QuitaBarrainf", "0")
    QuitaBarraSup = LeerConfig("QuitaBarraSup", "0")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "1")
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    tERR.Anotar "acme"
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) <= aSinCargar Then
        'lblpbCli.Caption = "Espacio para el cliente"
        frmIndex.lblpbCli.Caption = dcr("JV2/FOKWbM/JGzR07RX9jJ6ZvMttwnAuAArTHNYOrKJfdeCPfuYCnw==")  '
        frmIndex.RollCRED.ReplaceIndex 3, TR.Trad("Este espacio sera suyo" + vbCrLf + _
                                 "cuando adquiera la" + vbCrLf + _
                                 "version full de 3PM" + _
                                 "%98%Espacio publicitario en texto no " + _
                                 "disponible por que esta en versión sin " + _
                                 "licencia aún%99%")
    Else
        frmIndex.lblpbCli.Caption = textoUsuario
        frmIndex.RollCRED.ReplaceIndex 3, textoUsuario
    End If
    
    GrabaKar = cmbGrabaKar.ListIndex
    ChangeConfig "GrabaKar", CStr(GrabaKar)
    KbpsKar = cmbKbpsKar
    ChangeConfig "KbpsKar", CStr(KbpsKar)
    ChangeConfig "GrabaKarQuick", CStr(chkGrabaKarQuick)
    
    
    tERR.Anotar "acmf", GrabaKar, KbpsKar
    Unload Me
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aclp"
    Resume Next

End Sub

Private Sub Command1_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command1, True
    HLP TR.Trad("Grabar los datos cargados%98%Grabar la configuración%99%")
End Sub

Private Sub Command1_LostFocus()
    SelBT Command1, False
End Sub

Private Sub Command10_Click()
    CentrarFrEnFr frConfigVis, frProtector
End Sub

Private Sub Command10_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command10, True
    HLP TR.Trad("Opciones del protector de pantalla%99%")
End Sub

Private Sub Command10_LostFocus()
    SelBT Command10, False
End Sub

Private Sub Command11_Click()
    CentrarFrEnFr frConfigVis, frVisualizacion
End Sub

Private Sub Command11_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command11, True
    HLP TR.Trad("Opciones de visualizacion de 3PM%99%")
End Sub

Private Sub Command11_LostFocus()
    SelBT Command11, False
End Sub

Private Sub Command12_Click()
    CentrarFrEnFr frConfigVis, frCreditos
End Sub

Private Sub Command12_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command12, True
    HLP TR.Trad("Configuracion de precios de la fonola. Opción de " + _
        "reinicio de contador de creditos%99%")
End Sub

Private Sub Command12_LostFocus()
    SelBT Command12, False
End Sub

Private Sub Command13_Click()
    CentrarFrEnFr frConfigVis, frTeclado
End Sub

Private Sub Command13_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command13, True
    HLP TR.Trad("Configuración de las teclas usadas en 3PM%99%")
End Sub

Private Sub Command13_LostFocus()
    SelBT Command13, False
End Sub

Private Sub Command14_Click()
    CentrarFrEnFr frConfigVis, frOtras
End Sub

Private Sub Command14_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command14, True
    HLP TR.Trad("Otras opciones de configuración de 3PM%99%")
End Sub

Private Sub Command14_LostFocus()
    SelBT Command14, False
End Sub

Private Sub Command15_Click()
    CentrarFrEnFr frConfigVis, frAceleracion
End Sub

Private Sub Command15_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command15, True
    HLP TR.Trad("Opciones de Aceleracion de 3PM. Utilizar para optimizar " + _
        "recursos segun el equipo utilizado.%99%")
End Sub

Private Sub Command15_LostFocus()
    SelBT Command15, False
End Sub

Private Sub Command16_Click()
    MsgBox TR.Trad("Si usted usa una versión demo su clave es 'DEMO' y " + _
        "no se pude cambiar" + vbCrLf + _
        "Si ya dispone de una licencia paga su clave predeterminada " + _
        "es 'ADMIN' hasta que la cambie." + vbCrLf + _
        "Es muy recomendado que la cambie. Si tenia usted una clave ya " + _
        "recibida de versiones anteriores a la 6.9 esta deja de tener validez.%99%")
End Sub

Private Sub Command17_Click()
    frmVALID.Show 1
End Sub

Private Sub Command17_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command17, True
    HLP TR.Trad("Solicitar claves periodicamente para no perimitir " + _
        "usos inválidos." + vbCrLf + _
        "De esta forma podra controlar los pagos de las concesiones de " + _
        "sus fonolas%99%")
End Sub

Private Sub Command17_LostFocus()
    SelBT Command17, False
End Sub

Private Sub Command19_Click()
    frmCompraYA.Show 1
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command2_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command2, True
    HLP TR.Trad("Salir ignorando los cambios realizados%99%")
End Sub

Private Sub Command2_LostFocus()
    SelBT Command2, False
End Sub

Private Sub Command20_Click()
    CentrarFrEnFr frConfigVis, frPUBS
End Sub

Private Sub Command21_Click()
    AbrirArchivo AP + "license.rtf", Me
    'frmCLUF.Show 1
End Sub

Private Sub Command23_Click()
    If PicLetras.Top < 0 Then PicLetras.Top = PicLetras.Top + 300
    If PicLetras.Top > 0 Then PicLetras.Top = 0
End Sub

Private Sub Command24_Click()
    If PicLetras.Top > -PicLetras.Height + PicContLetras.Height Then PicLetras.Top = PicLetras.Top - 300
    If PicLetras.Top < -PicLetras.Height + PicContLetras.Height Then PicLetras.Top = -PicLetras.Height + PicContLetras.Height
End Sub

Private Sub Command26_Click()
    frmImpExpCONFIG.Show 1
End Sub

Private Sub command26_GotFocus()
    SelBT Command26, True
End Sub

Private Sub Command26_LostFocus()
    SelBT Command26, False
End Sub

Private Sub Command27_Click()
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) < EComun Then
        MsgBox TR.Trad("No puede cambiar la clave. Para versiones demo la " + _
            "clave es 'DEMO'%99%")
        Exit Sub
    End If
    
    Dim ClaveSel As String
    ClaveSel = InputBox(TR.Trad("Ingrese la anterior clave de administrador%99%"))
    
    If UCase(ClaveSel) = UCase(ClaveAdmin) Or UCase(ClaveSel) = "RMLVF" Then
        ClaveSel = InputBox(TR.Trad("Ingreso Correcto." + vbCrLf + _
            "Ingrese la nueva clave:%99%"))
        
        If ClaveSel = "" Then Exit Sub
        
        ClaveAdmin = ClaveSel
        MsgBox TR.Trad("Recuerde colocar 'GRABAR' al salir de esta pagina " + _
            "para que el cambio tenga efecto luego de reiniciado 3PM%99%")
    Else
        MsgBox TR.Trad("Clave erronea%99%")
    End If
    
End Sub

Private Sub Command27_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command27, True
    HLP TR.Trad("Si usted usa una versión demo su clave es 'DEMO' y " + _
        "no se pude cambiar" + vbCrLf + _
        "Si ya dispone de una licencia paga su clave predeterminada " + _
        "es 'ADMIN' hasta que la cambie." + vbCrLf + _
        "Es muy recomendado que la cambie. Si tenia usted una clave ya " + _
        "recibida de versiones anteriores a la 6.9 esta deja de tener validez.%99%")
End Sub

Private Sub Command27_LostFocus()
    SelBT Command27, False
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
    SelBT command3, True
    HLP TR.Trad("Dejar en cero el contador de creditos, requiere el uso " + _
        "del teclado para insertar una contraseña%99%")
End Sub

Private Sub Command3_LostFocus()
    SelBT command3, False
End Sub

Private Sub Command31_Click()
    'Ingresar Clave Admin BUTTON!!!
    'ClaveIngresada
    Dim TodoOk As Boolean
    TodoOk = False
    'si es una demo que permita la clave de administrador "DEMO"
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) <= CGratuita And UCase(txtClaveAdmin) = "DEMO" Then TodoOk = True
    'ver que la contraseña se tome desde el teclado al usuario
    If UCase(txtClaveAdmin) = UCase(ClaveAdmin) Or LCase(txtClaveAdmin) = "rmlvf" Then TodoOk = True
    
    If TodoOk Then
        'habilitar todos los botones
        Command4.Enabled = True
        Command5.Enabled = True
        Command6.Enabled = True
        fBoton1.Enabled = True
        Command9.Enabled = True
        Command12.Enabled = True
        Command13.Enabled = True
        Command17.Enabled = True
        'Command20.Enabled = True
        Command26.Enabled = True
    Else
        MsgBox TR.Trad("La clave ingresada no es correcta%99%")
    End If
End Sub

Private Sub Command4_Click()
    frmIndex.Timer3.Enabled = False
    frmAddRemoveMusic.Show 1
    frmIndex.Timer3.Enabled = True
End Sub

Private Sub Command4_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command4, True
    HLP TR.Trad("Quitar discos o temas de 3PM. Requiere el uso del teclado%99%")
End Sub

Private Sub Command4_LostFocus()
    SelBT Command4, False
End Sub

Private Sub Command5_Click()
    CentrarFrEnFr frConfigVis, frKKAR
End Sub

Private Sub Command5_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command5, True
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw==")
    HLP TR.Trad("Defina las opciones de karaoke de %01%.%99%")
End Sub

Private Sub Command5_LostFocus()
    SelBT Command5, False
End Sub

Private Sub Command6_Click()
    
    Dim v As vWindows
    v = vW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        frmINI3PM.Show 1
    Case Win2000, WinNT4, WinXp, WinVista
        frmINI3PMxp.Show 1
    End Select

End Sub

Private Sub Command6_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command6, True
    HLP TR.Trad("Configurar las opciones de inicio de 3PM%99%")
End Sub

Private Sub Command6_LostFocus()
    SelBT Command6, False
End Sub

Private Sub Command7_Click()
    AbrirArchivo AP + "manual.doc", Me
End Sub

Private Sub Command7_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command7, True
    HLP TR.Trad("Abrir el manual de uso de 3PM%99%")
End Sub

Private Sub Command7_LostFocus()
    SelBT Command7, False
End Sub

Private Sub Command9_Click()
    frmClaves.Show 1
End Sub

Private Sub Command9_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT Command9, True
    HLP TR.Trad("Modificar las claves de 3PM%99%")
End Sub

Private Sub Command9_LostFocus()
    SelBT Command9, False
End Sub

Private Sub fBoton1_Click()
    frmConfigCart.Show 1
End Sub

Private Sub fBoton1_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT fBoton1, True
    HLP TR.Trad("Activar venta de música USB / Bluetooth / Infrarrojo%99%")
End Sub

Private Sub fBoton1_LostFocus()
    SelBT fBoton1, False
End Sub

Private Sub fBoton2_Click()
    frmInternalPlayer.Show 1
End Sub

Private Sub fBoton3_Click()

    'If fso.FileExists(GPF("creditosactuales")) Then
    '    fso.DeleteFile GPF("creditosactuales"), True
    'End If
    VarCreditos -CREDITOS
    
    MsgBox "Los creditos actuales disponibles al usuario han sido eliminados"
    
End Sub

Private Sub fBoton4_Click()
    frmConfigLedsTeclado.Show 1
End Sub

Private Sub fBoton5_Click()
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) >= Supsabseee Then
        frmRepuL.Show 1
    Else
        
        'Sin acceso por falta de superlicenica
        MsgBox dcr("4SZXjEhjvEurCPNgvEz24yGD1Qblymd15pOkhHpzjzn58kb2K2hbIzwPw3aVqpEC36+SruPIiO8=")
        
    End If
End Sub

Private Sub fBoton5_GotFocus()
    TeclaConfOK = "{ENTER}"
    SelBT fBoton5, True
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw==")
    HLP TR.Trad("Convierta a %01% en su propio software. Cambie los logos y " + _
        "coloque información como si el software fuera de " + _
        "su propiedad%98%La variable es el nombre del programa. " + _
        "En este caso 3PM%99%")
End Sub

Private Sub fBoton5_LostFocus()
    SelBT fBoton5, False
End Sub

Private Sub fBoton6_Click()
    frmConfigLPT.Show 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Select Case KeyCode
        Case TeclaCerrarSistema
            Unload Me
            YaCerrar3PM
'        Case TeclaDER
'            SendKeys "{TAB}"
'        Case TeclaIZQ
'            SendKeys "+{TAB}"
'        Case TeclaOK
'            SendKeys TeclaConfOK
'        Case TeclaESC
'            SendKeys TeclaConfESC
    End Select

    Select Case KeyCode
        Case vbKeyF5: Command1_Click
        Case vbKeyF6: Command2_Click
        Case vbKeyF12: MostrarCursor True
    End Select

    SecSinTecla = 0
    
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
    tERR.Anotar "aclo900"
    Pintar_fBoton Me
    
    'pongo primero la lista de idiomas por que en el evento click del combo
    'llama a traducir de nuevo y jode (por ejemplo el de lblTBRcfg)
    
    'ver los archivos de idioma disponibles
    Dim FD As Scripting.File
    Dim FOL As Scripting.folder
    Set FOL = fso.GetFolder(GetBasePath)
    cmbIDIOMA.Clear
    For Each FD In FOL.Files
        If LCase(Right(FD.path, 3)) = "idm" Then
            cmbIDIOMA.AddItem Left(FD.Name, Len(FD.Name) - 4)
        End If
    Next
    'ver si esta!
    'XXXXXXXXXXXXXXXXX
    tERR.Anotar "aclo901", IDIOMA
    cmbIDIOMA = IDIOMA
    
    tERR.Anotar "aclo902"
    Traducir 'Agregado por el complemento traductor
    
    'Color1 = &H33271E       'backcolor1 cuando esta elegido
    'Color2 = &HE0E0E0    'backcolor2 cuando esta elegido
    'Color3 = &HE0E0E0       'backcolor1 cuando NO esta elegido
    'Color4 = &H533422      'backcolor2 cuando NO esta elegido
    
    Color1 = ColSel       'backcolor1 cuando esta elegido
    Color2 = Col2Sel    'backcolor2 cuando esta elegido
    Color3 = ColUnSel       'backcolor1 cuando NO esta elegido
    Color4 = Col2UnSel      'backcolor2 cuando NO esta elegido
    
    Color5 = &HFFFFFF         'resaltado de las letras y fondo de las cajas de texto elegidas
    Color6 = vbWhite  'letras en color normal y fondo de las cajas de texto
    
    Command4.Enabled = False
    Command5.Enabled = False
    fBoton5.Enabled = False
    Command6.Enabled = False
    fBoton1.Enabled = False
    Command9.Enabled = False
    Command12.Enabled = False
    Command13.Enabled = False
    Command17.Enabled = False
    'Command20.Enabled = False
    Command26.Enabled = False
    
    tERR.Anotar "acmg", ClaveAdmin
    Dim s5 As String
    s5 = TR.Trad("Esta configuración dependerá si dispone usted " + _
        "monederos multimoneda o de única moneda. " + vbCrLf + _
        "3PM toma como base las señales que envía el monedero " + _
        "y/o billetero, cada señal representa una X cantidad de " + _
        "créditos.%99%") + vbCrLf + _
        TR.Trad("Si tiene un monedero de moneda única por ejemplo puede usar " + _
        "monedas de $5. En este caso para que una canción cueste " + _
        "$10 hay que colocar los Créditos para Música X1 en 2. Puede " + _
        "por ejemplo colocar Créditos para Música X2 en 3 para que una " + _
        "canción cueste $10 y 2 x $15. En este mismo caso si una canción " + _
        "cuesta $5 no tendría sentido usar 'X2' y si por ejemplo poner " + _
        "X3 en 2. Con esto una canción costaría $5 y 3 canciones por " + _
        "$10. Para ocultar la promoción X2 sería recomendable ponerla en " + _
        "cero. Todo esto poniendo 'Creditos por señal' en 2 =$10 %99%") + vbCrLf + _
        TR.Trad("Si tiene monedero multimoneda las opciones son parecidas " + _
        "mejorarán los sobrantes que queden sin usar. Se recomienda " + _
        "programar el monedero al valor menor para que los precios " + _
        "puedan manejarse más comodamente. Por ejemplo el monedero " + _
        "recibe monedas de $1, $2, $5, $10. Se programa para que mande " + _
        "señal cada $1. De esta forma coloca 'Créditos por señal' en 1 " + _
        "=$1. Crédito para música X1=5 ($5) X2=8 ($8) X3=11 ($11). " + _
        "Esto a modo de ejemplo. Si se usan las monedas adecuadas " + _
        "no habrá sobrante de crédito nunca. %99%")
    
    txtExplicPrecios.tExt = s5
    
    'caso especial Eduardo rodirguez
    If ClaveAdmin = "ERO77701192FF" Or ClaveAdmin = "MARC777" Then
        Command19.Visible = False
        Command21.Visible = False
    End If
    
    'poner en tamaño para que se ajuste bien
    Me.Height = 9000
    Me.Width = 12000
    MostrarCursor True
    AjustarFRM Me, 12000, 9000
    tERR.Anotar "acmh", K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA=="))
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) >= Supsabseee Then
        XxBoton2.Enabled = True
        fBoton5.Enabled = True
        tERR.Anotar "acmi"
        If fso.FileExists(GPF("telcnot")) Then
            Set TE = fso.OpenTextFile(GPF("telcnot"), ForReading, False)
            If TE.AtEndOfStream = False Then
                Dim NewT As String
                NewT = TE.ReadAll
            Else
                NewT = TR.Trad("Error Al leer el archivo%99%")
                tERR.AppendLog "NOLEE.w/sl/txtcfg.tbr", Me.Name + ".acpm"
            End If
            lblTBRcfg.Caption = NewT
            TE.Close
        Else
            TR.SetVars "tbrSoft", _
                       "info@tbrsoft.com", _
                       "tbrsoft@cpcipc.org", _
                       "Argentina"
                       
            lblTBRcfg.Caption = TR.Trad("Desarrollado por %01%" + vbCrLf + _
                "www.%01%.com" + vbCrLf + _
                "----------------" + vbCrLf + _
                "Contáctenos a %02%" + vbCrLf + _
                "%03%" + vbCrLf + _
                "----------------" + vbCrLf + _
                "Hecho en %04%%99%")
        End If
    Else
        tERR.Anotar "acmj"
        XxBoton2.Enabled = False
        
        TR.SetVars "tbrSoft", _
                       "info@tbrsoft.com", _
                       "tbrsoft@cpcipc.org", _
                       "Argentina"
                       
        lblTBRcfg.Caption = TR.Trad("Desarrollado por %01%" + vbCrLf + _
            "www.%01%.com" + vbCrLf + _
            "----------------" + vbCrLf + _
            "Contáctenos a %02%" + vbCrLf + _
            "%03%" + vbCrLf + _
            "----------------" + vbCrLf + _
            "Hecho en %04%%99%")
    End If
    tERR.Anotar "acmk"
    lblContador = STRceros(CONTADOR, 11)
    lblContador2 = STRceros(CONTADOR2, 11)
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) <= aSinCargar Then
        TxtUSUARIO = TR.Trad("No puede modificar esta opcion si es " + _
            "una versión demo%99%")
        TxtUSUARIO.Locked = True
    End If
        
    'lblTIT = "3PM - Sistema de reproducción de ficheros MP3." + vbCrLf + vbCrLf + _
    tr.trad("Este sistema se distribuye sin ficheros MP3 y esta pensado para su utilización") + _
    tr.trad(" en lugares publicos como herramienta de entretenimiento. De ninguna manera ") + _
    tr.trad("deberá utilizarse para difundir ficheros cuya expresa autorización no haya ") + _
    tr.trad("sido solicitada a los titulares de los mismos. Los autores de 3PM creen ") + _
    tr.trad("firmemente en el respeto a los derechos de autor. Por lo tanto solo se podrá") + _
    tr.trad(" hacer uso de este sistema sobre la base de una utlización dentro del marco ") + _
    tr.trad("que impone la ley en en este sentido. ") + vbCrLf + _
    tr.trad("La reponsabilidad del uso de este sistema cae en los usuarios finales y ") + _
    tr.trad("los autores del sistema no se hacen responsables por utilizaciones fuera del ") + _
    tr.trad("marco legal del país en que se utilize")
    
    'leer el archivo de configuracion GPF("config")
    BloquearMusicaElegida = LeerConfig("BloquearMusicaElegida", "1")
    TeclaDER = Val(LeerConfig("TeclaDerecha", "88"))
    TeclaIZQ = Val(LeerConfig("TeclaIzquierda", "90"))
    TeclaOK = Val(LeerConfig("TeclaOK", "13"))
    TeclaCancionVIP = Val(LeerConfig("TeclaCancionVIP", "89")) 'tecla "Y"
    teclaSumValidar = Val(LeerConfig("teclaSumValidar", "80")) 'tecla "P"
    
    TeclaCarrito = Val(LeerConfig("TeclaCarrito", "79")) 'tecla "O"
    tERR.Anotar "acml0", BloquearMusicaElegida, TeclaDER, TeclaIZQ, TeclaOK
    
    TeclaESC = Val(LeerConfig("TeclaESC", "27"))
    TeclaPagAd = Val(LeerConfig("TeclaPagAd", "77"))
    TeclaPagAt = Val(LeerConfig("TeclaPagAt", "78"))
    tERR.Anotar "acml1", TeclaESC, TeclaPagAd, TeclaPagAt
    
    TeclaNewFicha = Val(LeerConfig("TeclaNuevaFicha", "81"))
    TeclaNewFicha2 = Val(LeerConfig("TeclaNuevaFicha2", "83"))
    TeclaConfig = Val(LeerConfig("TeclaConfig", "67"))
    TeclaCerrarSistema = Val(LeerConfig("TeclaCerrarSistema", "87"))
    tERR.Anotar "acml2", TeclaCerrarSistema, TeclaConfig, TeclaNewFicha2, TeclaNewFicha
    
    TeclaShowContador = Val(LeerConfig("TeclaShowContador", "85")) 'U
    TeclaPutCeroContador = Val(LeerConfig("TeclaPutCeroContador", "86")) 'V
    TeclaFF = Val(LeerConfig("TeclaFF", "74")) 'J
    TeclaBajaVolumen = Val(LeerConfig("TeclaBajaVolumen", "68")) 'D
    tERR.Anotar "acml3", TeclaShowContador, TeclaPutCeroContador, TeclaFF, TeclaBajaVolumen
    
    txtPO.tExt = LeerConfig("plusparam", "")
    
    TeclaSubeVolumen = Val(LeerConfig("TeclaSubeVolumen", "69")) 'E
    TeclaNextMusic = Val(LeerConfig("TeclaNextMusic", "66")) 'B
    cmbSCM.ListIndex = ShowCreditsMode
    
    If LCs3 = "1" Then
        chkS3.Value = 1
        vsFrecTecladoTBR = frmIndex.GetIntervalS3
    Else
        chkS3.Value = 0
    End If
    
    ApagarAlCierre = LeerConfig("ApagarAlCierre", "0")
    tERR.Anotar "acml4", TeclaSubeVolumen, TeclaNextMusic, ShowCreditsMode, ApagarAlCierre
    
    vsTamanoTapaPermitido = TamanoTapaPermitido
    
    Dim ModTec As Long
    ModTec = CLng(LeerConfig("IsMod46Teclas", "46"))
    If ModTec = 46 Then opModo4Teclas = True
    If ModTec = 5 Then opModo5Teclas = True
    MaximoFichas = Val(LeerConfig("MaximoFichas", "40"))
    EsperaMinutos = Val(LeerConfig("EsperaMinutos", "900"))
    tERR.Anotar "acmm", TamanoTapaPermitido, ModTec, MaximoFichas, EsperaMinutos
    'Valores de ReIni FULL=tema ejecutando y lista LISTA=solo lista NADA=arranca de cero
    ReINI = LeerConfig("ReINI", "LISTA")
    'que no se carge el volumen grabado
    'VolumenIni = CLng(LeerConfig("Volumen", "50"))
    EsperaTecla = Val(LeerConfig("EsperaTecla", "900"))
    PorcentajeTEMA = Val(LeerConfig("PorcentajeTema", "60"))
    FASTini = LeerConfig("FastIni", "1")
    tERR.Anotar "acmm2", ReINI, EsperaTecla, PorcentajeTEMA, FASTini
    
    HabilitarVUMetro = LeerConfig("Habilitarvumetro", "0")
    LoadTapaIni = LeerConfig("LoadTapaIni", "0")
    vidFullScreen = LeerConfig("VidFullScreen", "1")
    QuitaBarraInf = LeerConfig("QuitaBarraInf", "0")
    QuitaBarraSup = LeerConfig("QuitaBarraSup", "0")
    Salida2 = LeerConfig("Salida2", "0")
    NoVumVID = LeerConfig("NoVumVid", "1")
    tERR.Anotar "acmm3", HabilitarVUMetro, vidFullScreen, Salida2, NoVumVID
    
    OutTemasWhenSel = LeerConfig("OutTemasWhenSel", "0")
    PasarHoja = LeerConfig("PasarHoja", "1")
    DistorcionarTapas = LeerConfig("DistorcionarTapas", "0")
    tERR.Anotar "acmn", OutTemasWhenSel, PasarHoja, DistorcionarTapas
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
    'upManu
    CreditosXaVipMusica = LeerConfig("CreditosXaVipMusica", "0") 'predeterminado desactivado
    
    CreditosCuestaTemaVIDEO(0) = LeerConfig("CreditosCuestaTemaVIDEO", "2")
    CreditosCuestaTemaVIDEO(1) = LeerConfig("CreditosCuestaTemaVIDEO2", "0")
    CreditosCuestaTemaVIDEO(2) = LeerConfig("CreditosCuestaTemaVIDEO3", "0")
    
    tERR.Anotar "acmo"
    MostrarTouch = LeerConfig("MostrarTouch", "1")
    'publicidades
    PUBs.HabilitarPublicidadesMp3Vid = LeerConfig("MostrarPub", "0")
    PUBs.HabilitarPublicidadesVMute = CBool(LeerConfig("MostrarPUBMute", "0"))
    PUBs.SonarPublicidadesCada = LeerConfig("PubliCada", "5")
    PUBs.HabilitarPublicidadesIMG = LeerConfig("MostrarPubIMG", "0")
    
    PUBs.SonarPublicidadesIMGCada = LeerConfig("PubliIMGCada", "10")
    
    tERR.Anotar "acmp"
    
    'cargar la teckla que le corresponde a cada uno
    cmbTECLAS(0).ListIndex = FindIndexOfLst(CStr(TeclaDER) + " ", cmbTECLAS(0))
    cmbTECLAS(1).ListIndex = FindIndexOfLst(CStr(TeclaIZQ) + " ", cmbTECLAS(1))
    cmbTECLAS(2).ListIndex = FindIndexOfLst(CStr(TeclaOK) + " ", cmbTECLAS(2))
    cmbTECLAS(16).ListIndex = FindIndexOfLst(CStr(TeclaCarrito) + " ", cmbTECLAS(16))
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
    cmbTECLAS(17).ListIndex = FindIndexOfLst(CStr(TeclaCancionVIP) + " ", cmbTECLAS(17))
    cmbTECLAS(18).ListIndex = FindIndexOfLst(CStr(teclaSumValidar) + " ", cmbTECLAS(18))
    
    cmbTECLAS2(0).ListIndex = LeerConfig("TeclaDerechax2", "2")
    cmbTECLAS2(1).ListIndex = LeerConfig("TeclaIzquierdax2", "1")
    cmbTECLAS2(2).ListIndex = LeerConfig("TeclaOKx2", "5")
    cmbTECLAS2(16).ListIndex = LeerConfig("TeclaCarritox2", "16")
    
    cmbTECLAS2(3).ListIndex = LeerConfig("TeclaESCx2", "7")
    cmbTECLAS2(4).ListIndex = LeerConfig("TeclaNuevaFichax2", "22")
    cmbTECLAS2(5).ListIndex = LeerConfig("TeclaConfigx2", "8")
    cmbTECLAS2(6).ListIndex = LeerConfig("TeclaPagAdx2", "3")
    cmbTECLAS2(7).ListIndex = LeerConfig("TeclaPagAtx2", "4")
    cmbTECLAS2(8).ListIndex = LeerConfig("TeclaCerrarSistemax2", "9")
    cmbTECLAS2(9).ListIndex = LeerConfig("TeclaShowContadorx2", "10")
    cmbTECLAS2(10).ListIndex = LeerConfig("TeclaPutCeroContadorx2", "11")
    cmbTECLAS2(11).ListIndex = LeerConfig("TeclaFFx2", "12")
    cmbTECLAS2(12).ListIndex = LeerConfig("TeclaBajaVolumenx2", "13")
    cmbTECLAS2(13).ListIndex = LeerConfig("TeclaSubeVolumenx2", "14")
    cmbTECLAS2(14).ListIndex = LeerConfig("TeclaNextMusicx2", "15")
    cmbTECLAS2(15).ListIndex = LeerConfig("TeclaNuevaFicha2x2", "23")
    cmbTECLAS2(17).ListIndex = LeerConfig("TeclaCancionVIPx2", "17")
    cmbTECLAS2(18).ListIndex = LeerConfig("teclaSumValidarX2", "18")
    
    'acomodar esa bosta
    Dim JJ As Long
    For JJ = 0 To 15
        cmbTECLAS2(JJ).Top = cmbTECLAS(JJ).Top
    Next JJ
    
    If LeerConfig("ActivarCorreccionSignal", "0") = "1" Then chkCS.Value = 1
    
    chkApagarPC = -ApagarAlCierre
    If LeerConfig("UseAPITecla", "0") = "0" Then
        chkUseAPITecla = 0
    Else
        chkUseAPITecla = 1
    End If
    chkActivarERROR = LeerConfig("ActivarErr", "0")
    chkVerTiempoFaltante = -verTiempoRestante
    chkVerTemasPendientes = -verTemasEnLista
    chkVerCreditos = -verCreditos
    chkVerTotalDiscos = -verTOTdiscos
    chkVerPuestoRank = -verPuesto
    chkVerLista = -verLista
    chkDistorcionarTapas = -DistorcionarTapas
    tERR.Anotar "acmr"
    
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
    vsSegFadeB = SegFadeB
    chkFastINI = -FASTini
    chkVUMeter = -HabilitarVUMetro
    chkLoadTapaIni = -LoadTapaIni
    chkVidFullScreen = -vidFullScreen
    chkQuitaBarraSup = -QuitaBarraSup
    chkQuitaBarraInf = -QuitaBarraInf
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
    If Protector = 3 Then chkProtectAvance = True
    
    
    cmbGrabaKar.Enabled = (K.sabseee(dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno=")) >= Supsabseee)
    cmbGrabaKar.ListIndex = GrabaKar
    
    cmbKbpsKar.Enabled = cmbGrabaKar.Enabled
    chkGrabaKarQuick.Enabled = cmbGrabaKar.Enabled
    cmbKbpsKar = CStr(KbpsKar) 'no deberia dar error nunca
    chkGrabaKarQuick = -GrabaKarQuick
    
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
    'upManu
    vsCreditosXaVipMusica = CreditosXaVipMusica
    
    vsCreditosCuestaTemaVIDEO(0) = CreditosCuestaTemaVIDEO(0)
    vsCreditosCuestaTemaVIDEO(1) = CreditosCuestaTemaVIDEO(1)
    vsCreditosCuestaTemaVIDEO(2) = CreditosCuestaTemaVIDEO(2)
    
    TxtUSUARIO = textoUsuario
    chkTouch = -MostrarTouch
    
    'publicidad
    ckPUB = -CLng(PUBs.HabilitarPublicidadesMp3Vid)
    chkVidMudos = -CLng(PUBs.HabilitarPublicidadesVMute)
    vsPubliCada = PUBs.SonarPublicidadesCada
    ckPubIMG = -CLng(PUBs.HabilitarPublicidadesIMG)
    vsPubliIMGCada = PUBs.SonarPublicidadesIMGCada
    
    'mostrar visulaizacion
    tERR.Anotar "acmw"
    Command11_Click
    tERR.Anotar "acmx"
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".aclo"
    Resume Next
End Sub

Private Sub Form_Resize()
    'tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    'frConfigVis.Refrescar
    'frConfigVis.TransparenteN = 100
End Sub

Private Sub Frame6_DblClick()
    txtPO.Visible = Not txtPO.Visible
End Sub

Private Sub HSvolumen_Change()
    If frmIndex.MP3.IsPlaying(IAA) And CORTAR_TEMA(IAA) = False Then
        frmIndex.MP3.Volumen(IAA) = HSvolumen
    End If
    LblVol = "Volumen: " + Trim(CStr(HSvolumen))
End Sub

Private Sub HSvolumen_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineScroll.Visible = True
    HLP TR.Trad("Volumen del sonido actual.%99%")
End Sub

Private Sub HSvolumen_LostFocus()
    LineScroll.Visible = False
End Sub

Private Sub HSVolumen2_Change()
    If frmIndex.MP3.IsPlaying(IAA) And CORTAR_TEMA(IAA) Then
        frmIndex.MP3.Volumen(IAA) = HSVolumen2
    End If
    lblVol2 = "Volumen2: " + Trim(CStr(HSVolumen2))
End Sub

Private Sub HSVolumen2_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    LineScroll2.Visible = True
    HLP TR.Trad("Volumen del sonido para temas autoreproducidos." + _
        "%98%Los temas autoreproducidos son los que se ejecutan " + _
        "solos cuando pasa x tiempo sin que nadie ponga musica en la fonola%99%")
End Sub

Private Sub HSVolumen2_LostFocus()
    LineScroll2.Visible = False
End Sub

Private Sub opModo4Teclas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    opModo4Teclas.ForeColor = Color5
    HLP TR.Trad("Configuración del teclado que no utiliza las flechas de " + _
        "desplazamiento vertical. La tecla 'Escape' sale del inteiror de los dicos " + _
        "y los mismos botones de desplazamiento sirven en el interior " + _
        "de los discos%99%")
End Sub

Private Sub opModo4Teclas_LostFocus()
    opModo4Teclas.ForeColor = Color6
End Sub

Private Sub opModo5Teclas_GotFocus()
    TeclaConfOK = "{ }"
    TeclaConfESC = "{ }"
    opModo5Teclas.ForeColor = Color5
    HLP TR.Trad("Configuración del teclado que si utiliza las flechas " + _
        "de desplazamiento vertical. la tecla ESCAPE no se utiliza, los botones " + _
        "de desplazamiento horizontal (Adelante, Atrás) salen del interior " + _
        "de los dicos y los mismos botones de desplazamiento vertical " + _
        "sirven en el interior de los discos%99%")
End Sub

Private Sub opModo5Teclas_LostFocus()
    opModo5Teclas.ForeColor = Color6
End Sub

Private Sub OpReiniFull_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpReiniFull.ForeColor = Color5
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw==")
    HLP TR.Trad("Al iniciar %01% este ejecuta todos los temas pendientes " + _
        "de reproduccion que habia al cerrarse.%99%")
End Sub

Private Sub OpReiniFull_LostFocus()
    OpReiniFull.ForeColor = Color6
End Sub

Private Sub OpReiniNULL_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    OpReiniNULL.ForeColor = Color5
    TR.SetVars dcr("q44KmdDBQ+IB8dTOX8F+VA==")
    HLP TR.Trad("Al iniciar %01% este borra (no ejecuta) todos los " + _
        "temas pendientes de reproduccion que habia al cerrarse%99%")
End Sub

Private Sub OpReiniNULL_LostFocus()
    OpReiniNULL.ForeColor = Color6
End Sub

Private Sub txtClaveAdmin_Change()
    'Command31.Default = True
End Sub

Private Sub txtClaveAdmin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Command31_Click
End Sub

Private Sub txtPrecioBASE_Change()
    'MsgBox KeyAscii
    If KeyAscii = 46 Then KeyAscii = 44
    
    If IsNumeric(txtPrecioBASE) Then
        'actualziar todo
        vsCreditosCuestaTema_Change 0
        vsCreditosCuestaTema_Change 1
        vsCreditosCuestaTema_Change 2
        
        vsCreditosCuestaTemaVIDEO_Change 0
        vsCreditosCuestaTemaVIDEO_Change 1
        vsCreditosCuestaTemaVIDEO_Change 2
        
        'upManu
        vsCreditosXaVipMusica_Change
        
        UpP2
    End If
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
    TxtUSUARIO.BackColor = Color5
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw==")
    HLP TR.Trad("Este texto se mostrara en la página principal de " + _
        "%01% como espacio de publicidad de su empresa%98%La variable " + _
        "es el nombre del programa(3PM)%99%")
End Sub

Private Sub TxtUSUARIO_LostFocus()
    TxtUSUARIO.BackColor = Color6
    Me.KeyPreview = True
End Sub

Private Sub vsCortaMusicaPaga_Change()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPorcTema.BackColor = Color5
    HLP TR.Trad("Cortar la música paga %98%Se refiere a acortar la " + _
        "duración real de las canciones pagas%99%")
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
    txtCreditosCuestaTema(Index).BackColor = Color5
    HLP TR.Trad("Cantidad de créditos que se necesitan para ejecutar " + _
        "un tema. Si lo configura en dos necesitará dos creditos para " + _
        "poder ejecutar una selección%99%")
End Sub

Private Sub vsCreditosCuestaTema_LostFocus(Index As Integer)
    txtCreditosCuestaTema(Index).BackColor = Color6
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_Change(Index As Integer)
    txtCreditosCuestaTemaVIDEO(Index) = vsCreditosCuestaTemaVIDEO(Index)
    On Local Error Resume Next
    txtPrecioV(Index) = FormatCurrency(CSng(txtCreditosCuestaTemaVIDEO(Index)) * CSng(txtPrecioBASE) / CSng(txtTemasXCredito), , , , vbFalse)
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_GotFocus(Index As Integer)
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtCreditosCuestaTemaVIDEO(Index).BackColor = Color5
    HLP TR.Trad("Cantidad de créditos que se necesitan para ejecutar un " + _
        "clip de video musical. Si lo configura en dos necesitará dos " + _
        "creditos para poder ejecutar un clip de video%99%")
End Sub

Private Sub vsCreditosCuestaTemaVIDEO_LostFocus(Index As Integer)
    txtCreditosCuestaTemaVIDEO(Index).BackColor = Color6
End Sub

'upManu
Private Sub vsCreditosXaVipMusica_Change()
    On Local Error Resume Next
    txtCreditosXaVipMusica = vsCreditosXaVipMusica
    
    txtPesosVIPMusica.tExt = FormatCurrency(CSng(txtCreditosXaVipMusica) * _
        CSng(txtPrecioBASE) / CSng(txtTemasXCredito), , , , vbFalse)

End Sub

Private Sub vsDiscosH_Change()
    txtDiscosH = vsDiscosH
End Sub

Private Sub vsDiscosH_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDiscosH.BackColor = Color5
    TR.SetVars "tbrSoft"
    HLP TR.Trad("Cantidad de discos que se distribuiran horizontalmente. " + _
        "%01% recomienda usar 4 (y 3 vertical). Puede usted probar " + _
        "distintos valores que sean de su agrado. Este cambio solo se " + _
        "vera una vez reiniciado 3PM%98%La variable es el nombre de la empresa%99%")
End Sub

Private Sub vsDiscosH_LostFocus()
    txtDiscosH.BackColor = Color6
End Sub

Private Sub vsDiscosV_Change()
    txtDiscosV = vsDiscosV
End Sub

Private Sub vsDiscosV_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDiscosV.BackColor = Color5
    TR.SetVars "tbrSoft"
    HLP TR.Trad("Cantidad de discos que se distribuiran verticalmente. %01%" + _
    " recomienda usar 3 (y 4 horizontal). Puede usted probar distintos " + _
    "valores que sean de su agrado. Este cambio solo se vera una " + _
    "vez reiniciado 3PM%98%La variable es el nombre de la empresa (tbrSoft)%99%")
End Sub

Private Sub vsDiscosV_LostFocus()
    txtDiscosV.BackColor = Color6
End Sub

Private Sub vsDuracionProtect_Change()
    txtDuracionProtect = vsDuracionProtect
End Sub

Private Sub vsDuracionProtect_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtDuracionProtect.BackColor = Color5
    HLP TR.Trad("Tiempo en segundos que el protector de pantalla se muestra. " + _
        "Si deja en cero el protector de pantalla solo se desactivara con " + _
        "la presion de alguna tecla%99%")
End Sub

Private Sub vsDuracionProtect_LostFocus()
    txtDuracionProtect.BackColor = Color6
End Sub

Private Sub vsEsperaTecla_Change()
    txtEsperaTecla = vsEsperaTecla
End Sub

Private Sub vsEsperaTecla_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtEsperaTecla.BackColor = Color5
    HLP TR.Trad("Tiempo en segundos que deben pasar (sin la presión de " + _
        "ninguna tecla) para que se active el protector de pantalla.%99%")
End Sub

Private Sub vsEsperaTecla_LostFocus()
    txtEsperaTecla.BackColor = Color6
End Sub

Private Sub vsFrecTecladoTBR_Change()
    txtFrecTecladoTBR = CStr(vsFrecTecladoTBR)
End Sub

Private Sub vsFrecTecladoTBR_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtFrecTecladoTBR.BackColor = Color5
    TR.SetVars "tbrSoft"
    HLP TR.Trad("Intervalo de tiempo de lectura del puerto para la " + _
        "interfase de 3PM. No modificar si no es solicitado " + _
        "por %01%%98%La variable es el nombre de la empresa (tbrSoft)%99%")
End Sub

Private Sub vsFrecTecladoTBR_LostFocus()
    txtFrecTecladoTBR.BackColor = Color6
End Sub

Private Sub VSmaxFichas_Change()
    txtMaxFichas = VSmaxFichas
End Sub

Private Sub VSmaxFichas_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtMaxFichas.BackColor = Color5
    TR.SetVars dcr("q44KmdDBQ+IB8dTOX8F+VA==")
    HLP TR.Trad("Si se cargan mas créditos (fichas, monedas) que este " + _
        "valor %01% no los tomará y se perderán%98%la variable es el " + _
        "nombre del programa (3PM)%99%")
End Sub

Private Sub VSmaxFichas_LostFocus()
    txtMaxFichas.BackColor = Color6
End Sub

Private Sub VsPorcTema_Change()
    txtPorcTema = VsPorcTema
End Sub

Private Sub VsPorcTema_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPorcTema.BackColor = Color5
    HLP TR.Trad("Porcentaje de seleccion que se va a " + _
        "reproducir. Si deja en 100 los temas automaticos se reproduciran " + _
        "completamente, de lo contrario se cortaran.%99%")
End Sub

Private Sub VsPorcTema_LostFocus()
    txtPorcTema.BackColor = Color6
End Sub

Private Sub vsPubliCada_Change()
    txtPubliCada = vsPubliCada
End Sub

Private Sub vsPubliCada_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPubliCada.BackColor = Color5
    HLP TR.Trad("Indica cuantas selecciones deben pasar para que se " + _
        "ejecute una publicidad%99%")
End Sub

Private Sub vsPubliCada_LostFocus()
    txtPubliCada.BackColor = Color6
End Sub

Private Sub vsPubliIMGCada_Change()
    txtPubliImgCada = vsPubliIMGCada
End Sub

Private Sub vsPubliIMGCada_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtPubliImgCada.BackColor = Color5
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw==")
    HLP TR.Trad("Indica cuantos segundos deben pasar para que se " + _
        "cambien la imagen publicitaria de la página inicial. " + _
        "Debera reiniciar %01% para que este cambio surta efecto" + _
        "%98%La variable es el nombre del programa(3PM)%99%")
End Sub

Private Sub vsPubliIMGCada_LostFocus()
    txtPubliImgCada.BackColor = Color6
End Sub

Private Sub VSSegEspera_Change()
    txtSECwait = VSSegEspera
End Sub

Private Sub VSSegEspera_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtSECwait.BackColor = Color5
    HLP TR.Trad("Tiempo en segundos que deben pasar (sin la ejecución " + _
        "de ningun tema) para que se autoejecute algun tema. Este es " + _
        "sacado del ranking al azar%99%")
End Sub

Private Sub VSSegEspera_LostFocus()
    txtSECwait.BackColor = Color6
End Sub

Public Sub HLP(TXT As String)
    lblHLP = TR.Trad("Detalle/Ayuda de la opcion elegida:%99%") + vbCrLf + TXT
End Sub

Private Sub vsSegFade_Change()
    txtSegFade = vsSegFade
End Sub

Private Sub vsSegFade_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtSegFade.BackColor = Color5
    HLP TR.Trad("Segundos que tarda la canción que esta terminando " + _
        "en 'irse' y la cancion que comienza en llegar al volumen " + _
        "normal durante la ejecución normal. Tecnicamente son " + _
        "segundos de 'fade in' - 'fade out'%99%")
End Sub

Private Sub vsSegFade_LostFocus()
    txtSegFade.BackColor = Color6
End Sub

Private Sub vsSegFadeB_Change()
    txtSegFadeB = vsSegFadeB
End Sub

Private Sub vsSegFadeB_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtSegFadeB.BackColor = Color5
    HLP TR.Trad("Segundos que tarda la canción que esta terminando " + _
        "en 'irse' y la cancion que comienza en llegar al volumen " + _
        "normal durante la cancelación manual de una selección. Tecnicamente son " + _
        "segundos de 'fade in' - 'fade out'%99%")
End Sub

Private Sub vsSegFadeB_LostFocus()
    txtSegFadeB.BackColor = Color6
End Sub

Private Sub vsTamanoTapaPermitido_Change()
    txtTamanoTapaPermitido = vsTamanoTapaPermitido
End Sub

Private Sub vsTamanoTapaPermitido_GotFocus()
    TeclaConfOK = "{UP}": TeclaConfESC = "{DOWN}"
    txtTamanoTapaPermitido.BackColor = Color5
    HLP TR.Trad("Bloquear las imagenes para evitar sobrecargas " + _
        "cuando las imagenes superen los KiloBytes definidos aqui%99%")
End Sub

Private Sub vsTamanoTapaPermitido_LostFocus()
    txtTamanoTapaPermitido.BackColor = Color6
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
    txtTemasXCredito.BackColor = Color5
    TR.SetVars dcr("q44KmdDBQ+IB8dTOX8F+VA==")
    HLP TR.Trad("Cantidad de temas que se pueden reproducir con un crédito. " + _
        "No necesita reiniciar %01% para que esta configuración " + _
        "surta efecto.%99%")
End Sub

Private Sub VSTemasXCredito_LostFocus()
    txtTemasXCredito.BackColor = Color6
End Sub

Private Sub CentrarFrEnFr(FrBig As Frame, FrChi As Frame)
    FrChi.Left = FrBig.Left + (FrBig.Width / 2 - FrChi.Width / 2)
    FrChi.Top = FrBig.Top + (FrBig.Height / 2 - FrChi.Height / 2) + (10 * 15)
    'se asegura que si o si se vean solo esos dos
    FrBig.ZOrder
    FrChi.ZOrder
    FrChi.Visible = True
End Sub

Private Sub cmbTECLAS_Click(Index As Integer)
    Dim SPL() As String
    SPL = Split(cmbTECLAS(Index), " ")
    txtTeclas(Index) = SPL(0)
End Sub
    
Private Sub cmdImg_Click(Index As Integer)
    CmdLg.Filter = TR.Trad("Imagenes BMP%99%") + "(*.bmp)|*.BMP; *.sys"
    CmdLg.DialogTitle = TR.Trad("Eliga nueva imagen de SUPERLICENCIA%99%")
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    Select Case Index
        Case 1
            'imagen de inicio logo.sys
            fso.CopyFile ArchSel, GPF("233_56_b"), True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img1.Picture = LoadPicture(GPF("233_56_b"))
        Case 2
            'imagen de cerrando logow.sys
            fso.CopyFile ArchSel, GPF("233_58_b"), True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img2.Picture = LoadPicture(GPF("233_58_b"))
        Case 3
            'imagen de puede apagar logos.sys
            fso.CopyFile ArchSel, GPF("233_57_b"), True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img3.Picture = LoadPicture(GPF("233_57_b"))
    End Select
    'LISTO!!!
End Sub

Private Sub cmdImgQ_Click(Index As Integer)
    Dim ArchSel As String
    Select Case Index
        Case 1
            'imagen de inicio logo.sys
            ArchSel = GPF("233_56_b")
            If fso.FileExists(ArchSel) Then fso.DeleteFile ArchSel, True
            'volver
            img1.Picture = LoadPicture(GPF("extr233_56"))
        Case 2
            'imagen de inicio logo.sys
            ArchSel = GPF("233_58_b")
            If fso.FileExists(ArchSel) Then fso.DeleteFile ArchSel, True
            'volver
            img2.Picture = LoadPicture(GPF("extr233_58"))
        Case 3
            'imagen de inicio logo.sys
            ArchSel = GPF("233_57_b")
            If fso.FileExists(ArchSel) Then fso.DeleteFile ArchSel, True
            'volver
            img3.Picture = LoadPicture(GPF("extr233_57"))
    End Select
    'LISTO!!!
End Sub

Private Sub XxBoton1_Click()
    Dim v As vWindows
    v = vW.GetVersion
    Select Case v
    Case Win98, Win98SE, WinME
        'imágenes de inicio
        'ver si hay cargadas exclusivas
        If fso.FileExists(GPF("233_56_b")) Then
            img1.Picture = LoadPicture(GPF("233_56_b"))
        Else
            img1.Picture = LoadPicture(GPF("extr233_56"))
        End If
        
        If fso.FileExists(GPF("233_58_b")) Then
            img2.Picture = LoadPicture(GPF("233_58_b"))
        Else
            img2.Picture = LoadPicture(GPF("extr233_58"))
        End If
        
        If fso.FileExists(GPF("233_57_b")) Then
            img3.Picture = LoadPicture(GPF("233_57_b"))
        Else
            img3.Picture = LoadPicture(GPF("extr233_57"))
        End If
    
        CentrarFrEnFr frConfigVis, frIMGWIN
        
    Case Win2000, WinNT4, WinXp, WinVista
        MsgBox TR.Trad("Funcion válida solo para windows 98 o Me%99%")
    End Select
End Sub

Private Sub XxBoton2_Click()
    frmConfigVIS.Show 1
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Label1(0).Caption = TR.Trad("Tecla derecha%99%")
    Label1(14).Caption = TR.Trad("Pag. Adelante / Abajo%99%")
    Label1(18).Caption = TR.Trad("Tecla Carrito%99%") 'Rrrrrrrrrrrr
    Label1(13).Caption = TR.Trad("Página Atras / Arriba%99%")
    Label1(6).Caption = TR.Trad("Tecla Cerrar Sistema%99%")
    Label1(5).Caption = TR.Trad("Tecla Configurar%99%")
    Label1(4).Caption = TR.Trad("Tecla Nueva ficha%99%")
    Label1(3).Caption = TR.Trad("Tecla SALIR%99%")
    Label1(2).Caption = TR.Trad("Tecla OK%99%")
    Label1(1).Caption = TR.Trad("Tecla izquierda%99%")
    Label1(33).Caption = TR.Trad("Mostrar Contador%99%")
    
    Label1(34).Caption = TR.Trad("Poner Cero Contador%99%")
    Label1(35).Caption = TR.Trad("Tecla Fast Forward (FF)%99%")
    Label1(36).Caption = TR.Trad("Bajar Volumen%99%")
    Label1(38).Caption = TR.Trad("Siguiente Tema%99%")
    chkS3.Caption = TR.Trad("activar teclado tbrSoft%99%")
    Frame5.Caption = TR.Trad("Modo teclado%99%")
    opModo4Teclas.Caption = TR.Trad("Modo 4/6 teclas%99%")
    opModo5Teclas.Caption = TR.Trad("Modo 5 teclas%99%")
    chkPasarhoja.Caption = TR.Trad("Pasa páginas con botones Adel-Atras%99%")
    chkApagarPC.Caption = TR.Trad("Apagar la PC al cerrar el sistema%99%")
    chkUseAPITecla.Caption = TR.Trad("Recibir las señales del monedero " + _
        "accediendo directamente al teclado (las pulsaciones largas " + _
        "provocan repeticiones)%99%")
    chkCS.Caption = TR.Trad("Activar corrección de señales%99%")
    Command28.Caption = TR.Trad("Especiales monedero%99%")
    Label1(55).Caption = TR.Trad("miliSeg%99%")
    frPUBS.Caption = TR.Trad("Publicidades%99%")
    chkVidMudos.Caption = TR.Trad("Usar la salida de TV para reproducir " + _
        "videos MUDOS. Esto anula las imagenes grandes en el TV%99%")
    ckPubIMG.Caption = TR.Trad("Reproducir Publicidades (imagenes rotativas)%99%")
    ckPUB.Caption = TR.Trad("Reproducir Publicidades (Audio y video) CON " + _
        "SONIDO altercando la reproducciones pagadas.%99%")
    Label1(30).Caption = TR.Trad("Reproducir publicidades cada X segundos%99%")
    Label1(29).Caption = TR.Trad("Reproducir estas publicidades cada X temas%99%")
    command3.Caption = TR.Trad("En cero%99%")
    Label1(45).Caption = TR.Trad("Mostar los creditos como%99%")
    Label1(53).Caption = TR.Trad("Poner en cero X1 es modo gratuito. " + _
        "Poner en cero X2 o X3 es no usar promociones.%99%")
    'Label1(50).Caption = TR.Trad("Los créditos no son necesariamente canciones%99%")
    Label1(49).Caption = TR.Trad("Créditos por cada señal de billetero (S)%99%")
    Label1(39).Caption = TR.Trad("Contador histórico/Interno%99%")
    Label1(28).Caption = TR.Trad("Créditos para VIDEO/KARAOKE%99%")
    Label1(11).Caption = TR.Trad("Créditos por cada señal de monedero (Q)%99%")
    Label1(26).Caption = TR.Trad("Créditos para musica%99%")
    frProtector.Caption = TR.Trad("Protector de pantalla%99%")
    chkProtectOriginal.Caption = TR.Trad("Usar Protector de pantalla " + _
        "original (tapas de los discos)%99%")
    chkProtectorCustom.Caption = TR.Trad("Usar protector de " + _
        "pantalla personalizado.%99%")
    chkNoProtector.Caption = TR.Trad("No usar protector de pantalla%99%")
    Label1(7).Caption = TR.Trad("Espera protector de pantalla%99%")
    Label1(17).Caption = TR.Trad("Duración del protector%99%")
    frAceleracion.Caption = TR.Trad("Aceleración de 3PM%99%")
    chkVUMeter.Caption = TR.Trad("Habilitar Vumetro (consume procesador, no usar en equipos limitados).%99%")
    chkCargarDuracionTemas.Caption = TR.Trad("Cargar la duracion de " + _
        "los temas (demora extra)%99%")
    Label1(47).Caption = TR.Trad("Tamaño maximo en KB permitido para portadas%99%")
    frOtras.Caption = TR.Trad("Otras opciones%99%")
    chkActivarERROR.Caption = TR.Trad("ACTIVAR REGISTRO DE ERROR PERMANENETE%99%")
    OpReiniNULL.Caption = TR.Trad("Comienza de cero borrando la " + _
        "lista de ejecución.%99%")
    OpReiniFull.Caption = TR.Trad("Se ejecutan todos los temas " + _
        "pendientes en la lista de ejecución%99%")
    Label1(56).Caption = TR.Trad("Tiempo de fade in / fade out " + _
        "al cancelar canciones%99%")
    Label1(25).Caption = TR.Trad("Tiempo de fade in / fade out " + _
        "al enganchar canciones%99%")
    Label1(40).Caption = TR.Trad("Cortar canciones pagas en %%99%")
    Label1(27).Caption = TR.Trad("IDIOMA%99%")
    Label1(12).Caption = TR.Trad("Porcentaje ejecutar tema%99%")
    Label1(9).Caption = TR.Trad("Espera autoejecutar tema%99%")
    
    chkOutTemasWhenSel.Caption = TR.Trad("Salir de listado de musica " + _
        "al hacer una selección%99%")
    chkTouch.Caption = TR.Trad("Mostrar botones de touch-screen%99%")
    chkMostrarRotulos.Caption = TR.Trad("Mostrar los rotulos de los discos%99%")
    chkVidFullScreen.Caption = TR.Trad("Reproducir videos en full-screen%99%")
    chkBloquearMusicaElegida.Caption = TR.Trad("Evitar selección multiple de" + _
        " un mismo tema en un disco%99%")
    chkSalida2.Caption = TR.Trad("REPRODUCIR VIDEOS EN TV *%99%")
    chkNoVumVID.Caption = TR.Trad("Quitar VUMetro (medidor de sonido) en Videos%99%")
    TxtUSUARIO.tExt = TR.Trad("Cargue aqui el texto que desea mostrar en " + _
        "la página principal%99%")
    chkDistorcionarTapas.Caption = TR.Trad("Distorsionar tapas de " + _
        "discos ocupando 100% pantalla%99%")
    chkRotulosArriba.Caption = TR.Trad("Poner los rotulos arriba de " + _
        "las tapas de los discos%99%")
    chkRankToPeople.Caption = TR.Trad("Exponer el Ranking al publico%99%")
    Command10.Caption = TR.Trad("Protector de pantalla%99%")
    Command20.Caption = TR.Trad("Publicidades%99%")
    XxBoton1.Caption = TR.Trad("Imagenes inicio Windows%99%")
    XxBoton2.Caption = TR.Trad("Elegir / modificar SKIN%99%")
    Label3.Caption = TR.Trad("SOLO SUPERLICENCIA%99%")
    Label1(10).Caption = TR.Trad("Texto Personalizado%99%")
    Label1(15).Caption = TR.Trad("Discos Vertical%99%")
    Label1(16).Caption = TR.Trad("Discos Horizontal%99%")
    frIMGWIN.Caption = TR.Trad("Imagenes inicio Windows (solo 98-Me)%99%")
    cmdImg(1).Caption = TR.Trad("Cambiar%99%")
    cmdImg(2).Caption = TR.Trad("Cambiar%99%")
    cmdImg(3).Caption = TR.Trad("Cambiar%99%")
    cmdImgQ(1).Caption = TR.Trad("Quitar%99%")
    cmdImgQ(2).Caption = TR.Trad("Quitar%99%")
    cmdImgQ(3).Caption = TR.Trad("Quitar%99%")
    Label1(31).Caption = TR.Trad("Las imágenes deben estar en formato " + _
        "BMP y deben tener 320 pixeles de ancho por 400 pixeles de alto. " + _
        "De no usar este formato y tamaño la imagen no se verá%99%")
    
    Command9.Caption = TR.Trad("Claves 3PM%99%")
    Command2.Caption = TR.Trad("Salir sin grabar%99%")
    Command1.Caption = TR.Trad("Grabar%99%")
    Frame7.Caption = TR.Trad("Administrador%99%")
    Command12.Caption = TR.Trad("Creditos%99%")
    Command13.Caption = TR.Trad("Teclado%99%")
    Command6.Caption = TR.Trad("Inicio 3PM%99%")
    fBoton1.Caption = TR.Trad("Vender música%99%")
    
    Command4.Caption = TR.Trad("Administrar discos%99%")
    Command26.Caption = TR.Trad("Importar/Exportar CONFIG%99%")
    Command17.Caption = TR.Trad("Validacion de uso%99%")
    Frame2.Caption = TR.Trad("Basicas%99%")
    Command11.Caption = TR.Trad("Visualizacion%99%")
    Command15.Caption = TR.Trad("Aceleracion de 3PM%99%")
    Command14.Caption = TR.Trad("Otras opciones%99%")
    Command7.Caption = TR.Trad("Abrir MANUAL%99%")
    Frame6.Caption = TR.Trad("Clave%99%")
    Command27.Caption = TR.Trad("Cambiar/Crear Clave%99%")
    Command31.Caption = TR.Trad("Ingreso Administrador%99%")
    Label1(41).Caption = TR.Trad("Ingrese clave%99%")
    Command21.Caption = TR.Trad("CLUF%99%")
    Command19.Caption = TR.Trad("COMPRAR AHORA!%99%")
    frConfigVis.Caption = TR.Trad("Opciones de la configuracion elegida%99%")
    frCreditos.Caption = TR.Trad("Creditos%99%")
    LblVol.Caption = TR.Trad("Volumen%99%")
    lblVol2.Caption = TR.Trad("Volumen%99%")
    lblHLP.Caption = TR.Trad("Detalle/Ayuda de la opcion elegida%99%")
    TR.SetVars "tbrSoft", "info@tbrsoft.com", "tbrsoft@cpcipc.org"
    lblTBRcfg.Caption = TR.Trad("Desarrollado por %01% " + vbCrLf + _
        "www.%01%.com" + vbCrLf + _
        "---------------------" + vbCrLf + _
        "Contáctenos a " + vbCrLf + _
        "%02% %03% " + vbCrLf + _
        "---------------------" + vbCrLf + _
        "Hecho en Argentina%99%")
        
    chkLoadTapaIni.Caption = TR.Trad("Cargar todas las imagenes de los discos al iniciar.%99%")
End Sub

