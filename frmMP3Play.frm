VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmMP3Play 
   Caption         =   "MegaAmp -Hugo Gratz- MP3 Player"
   ClientHeight    =   8595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmMP3Play.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11400
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer6 
      Left            =   870
      Top             =   1860
   End
   Begin VB.Timer Timer5 
      Interval        =   1000
      Left            =   285
      Top             =   75
   End
   Begin VB.PictureBox piccuadro 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   2745
      Left            =   1410
      ScaleHeight     =   2745
      ScaleWidth      =   8580
      TabIndex        =   15
      Top             =   45
      Width           =   8580
      Begin VB.PictureBox Picture5 
         Height          =   30
         Left            =   825
         ScaleHeight     =   30
         ScaleWidth      =   540
         TabIndex        =   47
         Top             =   2760
         Width           =   540
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   2490
         Left            =   300
         ScaleHeight     =   2490
         ScaleWidth      =   7965
         TabIndex        =   16
         Top             =   75
         Width           =   7965
         Begin VB.PictureBox Picture4 
            Height          =   645
            Left            =   0
            ScaleHeight     =   585
            ScaleWidth      =   7845
            TabIndex        =   19
            Top             =   1785
            Width           =   7905
            Begin VB.CommandButton cmdLoad 
               Height          =   540
               Left            =   255
               Picture         =   "frmMP3Play.frx":0E42
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Abrir lista de Reproducción"
               Top             =   60
               Width           =   615
            End
            Begin VB.CommandButton cmdPlay 
               Height          =   510
               Left            =   2625
               Picture         =   "frmMP3Play.frx":1284
               Style           =   1  'Graphical
               TabIndex        =   25
               ToolTipText     =   "Play"
               Top             =   15
               Width           =   495
            End
            Begin VB.CommandButton cmdPause 
               Height          =   510
               Left            =   3225
               Picture         =   "frmMP3Play.frx":16C6
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Pausa"
               Top             =   15
               Width           =   495
            End
            Begin VB.CommandButton cmdStop 
               Height          =   510
               Left            =   3825
               Picture         =   "frmMP3Play.frx":1B08
               Style           =   1  'Graphical
               TabIndex        =   23
               ToolTipText     =   "Parar"
               Top             =   15
               Width           =   495
            End
            Begin VB.CommandButton cmdClose 
               Height          =   435
               Left            =   5520
               Picture         =   "frmMP3Play.frx":1F4A
               Style           =   1  'Graphical
               TabIndex        =   22
               ToolTipText     =   "Cerrar"
               Top             =   75
               Width           =   525
            End
            Begin VB.CommandButton cmdAbout 
               Caption         =   "Acerca de:"
               Height          =   360
               Left            =   6585
               TabIndex        =   21
               Top             =   120
               Width           =   1155
            End
            Begin VB.CommandButton mcdgrabar 
               Height          =   600
               Left            =   1110
               Picture         =   "frmMP3Play.frx":238C
               Style           =   1  'Graphical
               TabIndex        =   20
               ToolTipText     =   "Grabar Lista de Reproducción"
               Top             =   15
               Width           =   525
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   4530
            ScaleHeight     =   270
            ScaleWidth      =   2880
            TabIndex        =   18
            Top             =   285
            Width           =   2880
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000007&
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   4530
            ScaleHeight     =   270
            ScaleWidth      =   2880
            TabIndex        =   17
            Top             =   780
            Width           =   2880
         End
         Begin Proyecto1.MP3Play MP3Play1 
            Height          =   480
            Left            =   90
            TabIndex        =   27
            Top             =   150
            Width           =   480
            _ExtentX        =   847
            _ExtentY        =   847
         End
         Begin VB.Label lblhora 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   1395
            TabIndex        =   46
            Top             =   30
            Width           =   75
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H0080C0FF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   6
            Left            =   7050
            Top             =   1050
            Width           =   315
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   5
            Left            =   6645
            Top             =   1050
            Width           =   315
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H00C0C000&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   4
            Left            =   6240
            Top             =   1050
            Width           =   315
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   3
            Left            =   5835
            Top             =   1050
            Width           =   315
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H00800080&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   2
            Left            =   5430
            Top             =   1050
            Width           =   315
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H00FF80FF&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   1
            Left            =   5025
            Top             =   1050
            Width           =   330
         End
         Begin VB.Shape sp2 
            BorderColor     =   &H0080C0FF&
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   240
            Index           =   0
            Left            =   4620
            Top             =   1035
            Width           =   315
         End
         Begin VB.Label lblPosition 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   450
            Left            =   1890
            TabIndex        =   45
            Top             =   690
            Width           =   255
         End
         Begin VB.Label lblDuration 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Duración:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   720
            TabIndex        =   44
            Top             =   330
            Width           =   1005
         End
         Begin VB.Label lblFile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Título:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   240
            Left            =   675
            TabIndex        =   43
            Top             =   1395
            Width           =   660
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            X1              =   4470
            X2              =   7515
            Y1              =   1290
            Y2              =   1290
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1200
            Index           =   0
            Left            =   4620
            TabIndex        =   35
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   1
            Left            =   5025
            TabIndex        =   34
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   2
            Left            =   5430
            TabIndex        =   33
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   3
            Left            =   5835
            TabIndex        =   32
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   4
            Left            =   6240
            TabIndex        =   31
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   5
            Left            =   6645
            TabIndex        =   30
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblmovi 
            BackColor       =   &H8000000D&
            Height          =   1185
            Index           =   6
            Left            =   7050
            TabIndex        =   29
            Top             =   75
            Width           =   315
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Tiempo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   345
            Left            =   720
            TabIndex        =   28
            Top             =   765
            Width           =   990
         End
         Begin VB.Label lblbase 
            BackColor       =   &H000000FF&
            Height          =   1185
            Index           =   0
            Left            =   4620
            TabIndex        =   42
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H0000FFFF&
            Height          =   1185
            Index           =   1
            Left            =   5025
            TabIndex        =   41
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H00FF0000&
            Height          =   1185
            Index           =   2
            Left            =   5430
            TabIndex        =   40
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H0080FF80&
            Height          =   1185
            Index           =   3
            Left            =   5835
            TabIndex        =   39
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H00FFFFC0&
            Height          =   1185
            Index           =   4
            Left            =   6240
            TabIndex        =   38
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H00404080&
            Height          =   1185
            Index           =   5
            Left            =   6645
            TabIndex        =   37
            Top             =   75
            Width           =   315
         End
         Begin VB.Label lblbase 
            BackColor       =   &H00008000&
            Height          =   1185
            Index           =   6
            Left            =   7050
            TabIndex        =   36
            Top             =   75
            Width           =   315
         End
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   2625
         Left            =   135
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   8280
      End
   End
   Begin VB.Timer Timer4 
      Left            =   270
      Top             =   450
   End
   Begin VB.Timer Timer3 
      Left            =   255
      Top             =   885
   End
   Begin VB.PictureBox cuadro2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      Height          =   5100
      Left            =   315
      ScaleHeight     =   5100
      ScaleWidth      =   10875
      TabIndex        =   0
      Top             =   4035
      Width           =   10875
      Begin ComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   7230
         TabIndex        =   5
         Top             =   4245
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         _Version        =   327681
         Value           =   2490
         Alignment       =   0
         BuddyControl    =   "List1"
         BuddyDispid     =   196643
         OrigLeft        =   7545
         OrigTop         =   4125
         OrigRight       =   10245
         OrigBottom      =   4425
         Increment       =   10
         Max             =   3500
         Min             =   2400
         Orientation     =   1
         SyncBuddy       =   -1  'True
         Wrap            =   -1  'True
         BuddyProperty   =   65542
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdpasaauto 
         Caption         =   " Pasar Automático"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5505
         TabIndex        =   11
         ToolTipText     =   "Pasar en forma automática"
         Top             =   2160
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuitaTodo 
         Caption         =   "Quita Todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5505
         TabIndex        =   10
         ToolTipText     =   "Quitar todo de la Lista"
         Top             =   3645
         Width           =   1215
      End
      Begin VB.CommandButton CmdAgretodo 
         Caption         =   "Agrega todo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5505
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Agregar todo a la Lista"
         Top             =   2970
         Width           =   1215
      End
      Begin VB.CommandButton CmdQuitar 
         Height          =   495
         Left            =   5505
         Picture         =   "frmMP3Play.frx":27CE
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Quitar de a uno de la Lista"
         Top             =   1455
         Width           =   1215
      End
      Begin VB.CommandButton CmdAgrega 
         Height          =   495
         Left            =   5490
         Picture         =   "frmMP3Play.frx":2C10
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Pasar de uno a la Lista"
         Top             =   735
         Width           =   1215
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00404040&
         ForeColor       =   &H0000FFFF&
         Height          =   3345
         IntegralHeight  =   0   'False
         ItemData        =   "frmMP3Play.frx":3052
         Left            =   7215
         List            =   "frmMP3Play.frx":3054
         TabIndex        =   4
         ToolTipText     =   "Boton derecho del ratón selecciona"
         Top             =   885
         Width           =   2490
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         ItemData        =   "frmMP3Play.frx":3056
         Left            =   2460
         List            =   "frmMP3Play.frx":3063
         Sorted          =   -1  'True
         TabIndex        =   6
         Text            =   "Lista Mp3"
         Top             =   4170
         Width           =   2550
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   360
         Left            =   240
         TabIndex        =   3
         Top             =   825
         Width           =   2160
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000000&
         Height          =   2790
         Left            =   225
         TabIndex        =   2
         Top             =   1260
         Width           =   2175
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H80000001&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   3210
         Left            =   2460
         Pattern         =   "*.jpg;*.bmp;*.ico;*.wmf;*.gif;*.rle;*.cur"
         System          =   -1  'True
         TabIndex        =   1
         Top             =   825
         Width           =   2595
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   9960
         Top             =   4575
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "De Reproducción"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   7695
         TabIndex        =   50
         Top             =   480
         Width           =   1770
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   7170
         Picture         =   "frmMP3Play.frx":3080
         Stretch         =   -1  'True
         Top             =   345
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000008&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   9480
         TabIndex        =   14
         Top             =   405
         Width           =   255
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo del &Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   510
         TabIndex        =   13
         Top             =   4275
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo del &Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Index           =   0
         Left            =   525
         TabIndex        =   12
         Top             =   4275
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         FillStyle       =   0  'Solid
         Height          =   4950
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   10590
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   255
      Top             =   1350
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10635
      Top             =   375
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   255
      Top             =   1845
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   90
   End
   Begin VB.Label lbltitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Left            =   2355
      TabIndex        =   48
      Top             =   2835
      Width           =   90
   End
   Begin VB.Menu archi 
      Caption         =   "Archivo"
      Visible         =   0   'False
      Begin VB.Menu smenu1 
         Caption         =   "&Abrir"
         Index           =   0
      End
      Begin VB.Menu smenu1 
         Caption         =   "&Grabar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMP3Play"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public voltata As Long
Dim X As New CommonDialog
Dim mover, mtitulo, p, a, b, c, t As Integer
Dim pausa, pausa2 As Boolean
Dim Mytitulo, MyText As String

Const conInterval = 50
Const conIntervalPlus = 55

Dim CurrentValue As Double

Private Sub cmdAbout_Click()
    frmCtlAbout.Show vbModal, Me
End Sub

Private Sub cmdClose_Click()
    MP3Play1.mmStop
    Unload Me
End Sub

Private Sub cmdLoad_Click()
smenu1_Click 0
End Sub
Private Sub cmdPlay_Click()
Timer1.Interval = 500

If pausa2 = True Then
    MP3Play1.mmPlay
Else
'MP3Play1.Enabled = True
If pausa2 = False Then MP3Play1.FileName = List1.List(p)
      MP3Play1.mmPlay
End If
   pausa = False
   Timer3.Interval = 10
   cmdPlay.Enabled = False

pausa2 = False
End Sub
Private Sub cmdPause_Click()
    Timer3.Interval = 0
    pausa2 = True
    MP3Play1.mmPause
cmdPlay.Enabled = True
End Sub

Private Sub cmdStop_Click()
 Timer3.Interval = 0
MP3Play1.Enabled = False
    MP3Play1.mmStop
cmdPlay.Enabled = True
End Sub



Private Sub Form_Unload(Cancel As Integer)
    MP3Play1.mmStop
End Sub

Private Sub List1_DblClick()
 On Error GoTo imageErrs

     If List1.ListCount > 0 Then
          
      If MP3Play1.FileName = "" Then
          p = List1.ListIndex
          cmdPlay_Click
       cmdPlay.Enabled = False
       Else
          
           MP3Play1.FileName = ""
           MP3Play1.mmStop
           p = List1.ListIndex
          MP3Play1.FileName = List1.List(p)
           MP3Play1.Enabled = True
          MP3Play1.mmPlay
           Timer3.Interval = 10
          cmdPlay.Enabled = False
      End If
   
      Exit Sub
    End If
imageErrs:
errores
End Sub
Private Sub mcdgrabar_Click()
smenu1_Click 1
End Sub
Private Sub Timer1_Timer()
On Error GoTo errores
   If MP3Play1.IsPlaying = True Then
        lblDuration = "Duración: " & MP3Play1.Length
         lblPosition = MP3Play1.Position
 If pausa2 = False Then
   MP3Play1.FileName = List1.List(p)
 End If
    AUX = Len(MP3Play1.FileName) - 1
        ct = 1
        Do While Mid(MP3Play1.FileName, AUX, 1) <> "\"
            ct = ct + 1
            AUX = AUX - 1
        Loop
     lblFile = "Título: " & Right(MP3Play1.FileName, ct)
      
   End If
   


  If MP3Play1.IsPlaying = False Then
   
   If pausa2 = False Then
   
     If MP3Play1.Enabled = False Then
       lblFile = "Reproducción Detenida o Parada "
      Else
         p = p + 1
         MP3Play1.FileName = ""
         MP3Play1.mmStop
         If p > List1.ListCount - 1 Then p = 0
         List1.ListIndex = p
         MP3Play1.FileName = List1.List(p)
          MP3Play1.Enabled = True
         MP3Play1.mmPlay
      End If
    End If

 End If
errores:
'cmdStop_Click
 
End Sub



Private Sub cuadro2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu archi
 End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu archi
 End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Dim Msg ' Declara la variable.
    If UnloadMode > 0 Then
        ' Si sale de la aplicación.
        Msg = "¿Realmente desea salir de la aplicación?"
    Else
        ' Si sólo se cierra el formulario.
        Msg = "¿Realmente desea cerrar el formulario?"
    End If
    ' Si el usuario hace clic en el botón No, se detiene QueryUnload.
    If MsgBox(Msg, vbQuestion + vbYesNo, Me.Caption) = vbNo Then Cancel = True
End Sub
Private Sub Form_Load()

a = 1200
b = 1200
Combo1.AddItem "*.mp3;*.wma"
Dir1.path = App.path
If Combo1.Text = "Lista Mp3" Then
File1.pattern = "*.mp3;*.wma"
End If

cmdPlay.Enabled = False
cmdPause.Enabled = False
cmdStop.Enabled = False
 For l = 0 To 6
   lblmovi(l).BackColor = &H0&
  Next l
For t = 0 To 6
   sp2(t).Visible = False
  Next t
pantalla Me
Centrar2 cuadro2, Me
Timer4.Interval = 10
End Sub

'Private Sub Form_Resize()
'fondocolor Me, 0, 0, 1
'
'End Sub

Private Sub Drive1_Change()
 On Error GoTo DriveErrs
    Dir1.path = Drive1.drive
   Exit Sub
DriveErrs:
errores
End Sub
Private Sub Dir1_Change()
File1.path = Dir1.path

End Sub
Private Sub File1_Click()
On Error GoTo imageErrs

    If Right(File1.path, 1) = "\" Then
        titulo File1.path + File1.FileName
    Else
        titulo File1.path + "\" + File1.FileName
    End If
   
     Exit Sub

imageErrs:

errores


End Sub
Sub titulo(titu As String)
 
 AUX = Len(titu) - 1
        ct = 1
        Do While Mid(titu, AUX, 1) <> "\"
            ct = ct + 1
            AUX = AUX - 1
        Loop

     lbltitulo = Right(titu, ct)
 

End Sub
Private Sub Combo1_Click()
File1.pattern = Combo1.Text
If Combo1.Text = "Lista Mp3" Then
File1.pattern = "*.mp3;*.wma"
End If
End Sub

Private Sub CmdAgrega_Click()
canti = List1.ListCount + 1
If File1.FileName <> "" Then
    If Right(File1.path, 1) = "\" Then
        List1.AddItem File1.path + File1.FileName
      Label2.Caption = canti
    cmdPlay.Enabled = True
cmdPause.Enabled = True
cmdStop.Enabled = True
    Else
      List1.AddItem File1.path + "\" + File1.FileName
      Label2.Caption = canti
   cmdPlay.Enabled = True
cmdPause.Enabled = True
cmdStop.Enabled = True
   End If


End If
canti = canti + 1

End Sub

Private Sub CmdAgretodo_Click()
With File1
If .ListCount > 0 Then
    For i = 0 To .ListCount - 1
        If Right(.path, 1) = "\" Then
            List1.AddItem .path + .List(i)
         Label2.Caption = i
        Else
         List1.AddItem .path + "\" + .List(i)
        Label2.Caption = i + 1
        End If
    Next i
End If
End With
cmdPlay.Enabled = True
cmdPause.Enabled = True
cmdStop.Enabled = True
' mcimp3.FileName = List1.List(p)
'    On Error GoTo mci_error
'    mcimp3.Command = "Open"
'mci_error:
'    On Error GoTo 0

End Sub

Private Sub CmdQuitar_Click()


With List1
canti = .ListCount
If .ListIndex >= 0 Then
    AUX = .ListIndex
    .RemoveItem .ListIndex
    canti = canti - 1
    Label2.Caption = canti
    If .ListCount > 0 Then
        If .ListCount = AUX Then
            .ListIndex = .ListCount - 1
        Else
            .ListIndex = AUX
        End If
    End If
End If
If .ListCount = 0 Then
cmdStop_Click
cmdPlay.Enabled = False
cmdPause.Enabled = False
cmdStop.Enabled = False
End If

End With

End Sub

Private Sub CmdQuitaTodo_Click()
cmdStop_Click
List1.Clear
Label2.Caption = "0 "
cmdPlay.Enabled = False
cmdPause.Enabled = False
p = 0

End Sub

Private Sub File1_DblClick()

With File1
 canti = List1.ListCount + 1
If .FileName <> "" Then
    If Right(.path, 1) = "\" Then
        List1.AddItem .path + .FileName
          cmdPlay.Enabled = True
          cmdPause.Enabled = True
          cmdStop.Enabled = True
    Else
        List1.AddItem .path + "\" + .FileName
        Label2.Caption = canti
           cmdPlay.Enabled = True
           cmdPause.Enabled = True
           cmdStop.Enabled = True
    
    End If
End If
End With
canti = canti + 1

End Sub

Private Sub list1_ItemCheck(índice As Integer)
 On Error GoTo imageErrs

     If List1.ListCount > 0 Then
         
     Exit Sub
    End If
imageErrs:
 errores
End Sub

'Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Label2.Caption = List1.ListCount
'End Sub

Private Sub cmdpasaauto_Click()
   If mover = 0 Then
       cmdpasaauto.Caption = "Parar"
       mover = 1
    Else
       cmdpasaauto.Caption = "   Pasar Automático"
       mover = 0
   
    End If
 
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
   PopupMenu archi
 End If
End Sub

Private Sub mquitar_Click(Index As Integer)
Select Case Index
   Case 0
       CmdQuitar_Click
   Case 1
      CmdQuitaTodo_Click
End Select

End Sub

Private Sub smenu1_Click(Index As Integer)
Select Case Index
   Case 0
       CommonDialog1.CancelError = True
        On Error GoTo errhandler
         With CommonDialog1
           .Filter = "Lista(*.ml3)|*.ml3"
           .ShowOpen
           List1.Clear
           i = 0
           Open .FileName For Input As #1
        
           Do While Not EOF(1)
             Input #1, auxiliar
             List1.List(i) = auxiliar
             i = i + 1
            Loop
           Close #1
          Label2 = List1.ListCount
          End With
       cmdPlay.Enabled = True
           cmdPause.Enabled = True
           cmdStop.Enabled = True
    
   
   Case 1
     If List1.ListCount > 0 Then
        CommonDialog1.CancelError = True
        On Error GoTo errhandler
         With CommonDialog1
           .Filter = "Lista(*.ml3)|*.ml3"
           .flags = cdlOFNOverwritePrompt
           .ShowSave
           Open .FileName For Output As #1
           
           For i = 0 To List1.ListCount - 1
            Print #1, List1.List(i)
            Next
            
            Close #1
          End With
      Else
      
      MsgBox "Debe establecer una lista de reproducciòn"
      End If
Exit Sub
 End Select
errhandler:
     
End Sub

Private Sub Timer2_Timer()

   If mover = 1 Then pasar
End Sub

Private Sub pasar()

On Error GoTo imageErrs
  If File1.ListCount - 1 = File1.ListIndex Then
         File1.ListIndex = 0
   canti = 0
  End If
                  File1_DblClick
                  
                  File1.ListIndex = File1.ListIndex + 1
                  Label2.Caption = List1.ListCount
         
        If File1.ListCount - 1 = File1.ListIndex Then
          File1_DblClick
           mover = 0
            cmdpasaauto.Caption = "   Pasar Automático"
        End If
      
 
imageErrs:
errores
canti = canti + 1

End Sub
Private Sub cmdvolver_Click()
End
End Sub



Private Sub Timer3_Timer()
a = a - 200: If a = 0 Then a = 1200
b = b - 100: If b = 0 Then b = 1200
   For g = 0 To 6
   lblmovi(g).Height = a
   
   Next g
Timer4.Interval = 10
 For u = 0 To 6 Step 2
   lblmovi(u).Height = b
  
 Next u
End Sub

Private Sub Timer4_Timer()
 
If c < 7 Then
 sp2(c).Visible = True
 Else
   For t = 0 To 6
   sp2(t).Visible = False
  Next t

 End If
 
 c = c + 1: If c = 8 Then c = 0

End Sub



Private Sub Timer5_Timer()
lblhora = Format(Now)
End Sub
