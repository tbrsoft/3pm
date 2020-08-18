VERSION 5.00
Begin VB.Form frmTemasDeDisco 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer RelojTDD 
      Enabled         =   0   'False
      Left            =   30
      Top             =   60
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   8985
      Left            =   150
      TabIndex        =   1
      Top             =   30
      Width           =   11805
      Begin VB.TextBox lstAgregados 
         BackColor       =   &H00000080&
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
         Height          =   960
         Left            =   30
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   7050
         Width           =   7080
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Touch"
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
         Height          =   1305
         Left            =   7200
         TabIndex        =   9
         Top             =   7620
         Width           =   4515
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
            Height          =   950
            Left            =   2280
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdDiscoAd 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmTemasDeDisco.frx":0000
            Height          =   950
            Left            =   1200
            Picture         =   "frmTemasDeDisco.frx":0CFD
            Style           =   1  'Graphical
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdDiscoAt 
            BackColor       =   &H00C0C0C0&
            DownPicture     =   "frmTemasDeDisco.frx":15D5
            Height          =   950
            Left            =   120
            Picture         =   "frmTemasDeDisco.frx":2347
            Style           =   1  'Graphical
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "SALIR"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   950
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.ListBox lstEXT 
         BackColor       =   &H00404080&
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
         ForeColor       =   &H00C0E0FF&
         Height          =   1605
         IntegralHeight  =   0   'False
         ItemData        =   "frmTemasDeDisco.frx":2C8A
         Left            =   8010
         List            =   "frmTemasDeDisco.frx":2C9D
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   4905
         Visible         =   0   'False
         Width           =   2610
      End
      Begin VB.ListBox lstTIME 
         BackColor       =   &H00404080&
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
         ForeColor       =   &H00C0E0FF&
         Height          =   6555
         IntegralHeight  =   0   'False
         Left            =   45
         TabIndex        =   4
         Top             =   480
         Width           =   1185
      End
      Begin VB.ListBox lstTemas 
         BackColor       =   &H00404080&
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
         ForeColor       =   &H00C0E0FF&
         Height          =   6555
         IntegralHeight  =   0   'False
         Left            =   1260
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   480
         Width           =   5865
      End
      Begin VB.Label lblCOMOSALIR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         Caption         =   "PRESIONE ESC PARA SALIR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   15
         Top             =   8040
         Width           =   7065
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7980
         TabIndex        =   14
         Top             =   3060
         Visible         =   0   'False
         Width           =   345
      End
      Begin VB.Label lblPrecios 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "1 coin = 8 creditos / 8 creditos = 1 tema / 8 creditos = 1 VIDEO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   495
         Left            =   30
         TabIndex        =   13
         Top             =   8340
         Width           =   7065
      End
      Begin VB.Label lblNoEjecuta 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "NO HAY CREDITO PARA EJECUTAR"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   7200
         TabIndex        =   7
         Top             =   6840
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   4515
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         Caption         =   "TEMAS EN ESTE DISCO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Width           =   7065
      End
      Begin VB.Label lblDataDisco 
         Alignment       =   2  'Center
         BackColor       =   &H000040C0&
         Caption         =   "No hay datos adicionales del disco"
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
         Height          =   3225
         Left            =   7200
         TabIndex        =   3
         Top             =   4200
         UseMnemonic     =   0   'False
         Width           =   4500
      End
      Begin VB.Label lblDisco 
         Alignment       =   2  'Center
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Titulo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7200
         TabIndex        =   2
         Top             =   3660
         UseMnemonic     =   0   'False
         Width           =   4545
      End
      Begin VB.Image TapaCD 
         BorderStyle     =   1  'Fixed Single
         Height          =   3300
         Left            =   7740
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmTemasDeDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SegSinTecla As Long
Dim YaInicio As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    SegSinTecla = 0 'protector para salir de esta frm
    SecSinTecla = 0 'preteccion global de pantalla
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    On Local Error GoTo FallaKD
    'ver detalle mas abajo de que mierda es esto y en el gral de este frm
    YaInicio = YaInicio + 1
    'puede no escuchar el coin!!!!!!
    'esto se pone mas abajo!!!!
    'If YaInicio <= 1 Then Exit Sub
        
    
    Select Case RealKeyCode
        
        
        Case TeclaOK
            If YaInicio <= 1 Then Exit Sub
            
            If (frmIndex.MP3.IsPlaying(0) Or frmIndex.MP3.IsPlaying(1)) Then
                    
                    If BloquearMusicaElegida Then
                        lstTemas.List(lstTemas.ListIndex) = "----------"
                        lstTIME.List(lstTIME.ListIndex) = "---"
                    End If
                    SaltarEspaciosLstTemas True
                    
                    '************ NOV 2006
                    'si era un gratuito pasarse al que sigue
                    If CORTAR_TEMA(IAA) Then EMPEZAR_SIGUIENTE 6
                    '************
                    If OutTemasWhenSel Then Unload Me
    
End Sub

Private Sub Form_Load()
    If MostrarTouch = False Then Frame2.Visible = False
End Sub

Private Sub lstTemas_Click()
    On Local Error Resume Next
    If CargarDuracionTemas Then lstTIME.ListIndex = lstTemas.ListIndex
End Sub

Private Sub RelojTDD_Timer()
    'relojTemasDeDisco
    SegSinTecla = SegSinTecla + 1
    Label2 = SegSinTecla
    If SegSinTecla = 20 Then
        RelojTDD.Enabled = False
        Unload Me
    End If
End Sub
Private Sub SaltarEspaciosLstTemas(HaciaAdelante As Boolean)
    'cuando eligo un tema lo saco para que no haga macana
    'el secreto es no generar el listindex salvo que se haya encontrado...
    'uso la prop LIST() que puede ver sin tocar!!!!!!!
    Dim a As Long
    Dim CC As Long
    Dim Ahora As Long
    Ahora = lstTemas.ListIndex
    
    Dim nINI As Long, nFin As Long, StepMio As Long
    If HaciaAdelante Then
        nINI = Ahora
        nFin = lstTemas.ListCount - 1
        StepMio = 1
    Else
        nINI = Ahora
        nFin = 0
        StepMio = -1
    End If
    Dim Vueltas As Long
    Vueltas = 0
ReiniLST:
    Vueltas = Vueltas + 1
    'si da 4 vueltas es que no hay!!
    If Vueltas = 4 Then
        Unload Me
        Exit Sub
    End If
    For a = nINI To nFin Step StepMio
        If lstTemas.List(a) <> "----------" Then
            'ya esta lo encontro!!!!!!!
            'ir ahi!!!
            lstTemas.ListIndex = a
            Exit For
        Else
            'si es el ultimo......!!
            If HaciaAdelante Then
                If a = nFin Then 'este es lstTemas.ListCount - 1
                    'voy al primero
                    nINI = 0
                    GoTo ReiniLST
                End If
            Else
                If a = nFin Then 'este es 0
                    'voy al ultimo
                    nINI = lstTemas.ListCount - 1
                    GoTo ReiniLST
                End If
            End If
        End If
    Next
End Sub
