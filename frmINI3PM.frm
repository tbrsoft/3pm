VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmINI3PM 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de 3PM"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   Icon            =   "frmINI3PM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton Command2 
      Height          =   405
      Left            =   2070
      TabIndex        =   28
      Top             =   5940
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   405
      Left            =   2070
      TabIndex        =   27
      Top             =   5460
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar configuración"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Height          =   2475
      Left            =   6240
      TabIndex        =   19
      Top             =   5250
      Width           =   5055
      Begin VB.OptionButton OpApagar3PM 
         BackColor       =   &H00000000&
         Caption         =   "Imagen de 3PM"
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
         Height          =   315
         Left            =   1380
         TabIndex        =   21
         Top             =   660
         Width           =   2055
      End
      Begin VB.OptionButton OpApagarWIN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Imagen original"
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
         Height          =   315
         Left            =   1710
         TabIndex        =   20
         Top             =   1740
         Width           =   1875
      End
      Begin VB.Label lblNoImgApagar 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Imagen no disponible. De todas formas se puede habilitar"
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
         Left            =   3660
         TabIndex        =   25
         Top             =   570
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image imgApagar3PM 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   120
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1200
      End
      Begin VB.Image imgApagarWIN 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   3660
         Stretch         =   -1  'True
         Top             =   570
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00533422&
         Caption         =   "Imagen ""Ahora puede apagar el equipo"""
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
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   4785
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   2475
      Left            =   6240
      TabIndex        =   15
      Top             =   2700
      Width           =   5055
      Begin VB.OptionButton OpCerrando3PM 
         BackColor       =   &H00000000&
         Caption         =   "Imagen de 3PM"
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
         Height          =   315
         Left            =   1410
         TabIndex        =   17
         Top             =   750
         Width           =   2055
      End
      Begin VB.OptionButton OpCerrandoWIN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Imagen original"
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
         Height          =   315
         Left            =   1740
         TabIndex        =   16
         Top             =   1830
         Width           =   1875
      End
      Begin VB.Label lblNoImgCerrando 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Imagen no disponible. De todas formas se puede habilitar"
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
         Left            =   3690
         TabIndex        =   24
         Top             =   690
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Image imgCerrando3PM 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   150
         Stretch         =   -1  'True
         Top             =   690
         Width           =   1200
      End
      Begin VB.Image imgCerrandoWIN 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   3690
         Stretch         =   -1  'True
         Top             =   690
         Width           =   1200
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00533422&
         Caption         =   "Imagen ""Windows se esta cerrando"""
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
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   18
         Top             =   360
         Width           =   4785
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   2625
      Left            =   150
      TabIndex        =   11
      Top             =   2130
      Width           =   5805
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   $"frmINI3PM.frx":0442
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
         Height          =   855
         Left            =   330
         TabIndex        =   14
         Top             =   1440
         Width           =   5385
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "No cargar 3PM al iniciar Windows. Recomendado si el equipo no es de uso exclusivo de 3PM."
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
         Height          =   705
         Left            =   330
         TabIndex        =   13
         Top             =   660
         Width           =   5355
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00533422&
         Caption         =   "Cargar 3PM al Iniciar el equipo"
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
         Height          =   225
         Index           =   3
         Left            =   330
         TabIndex        =   12
         Top             =   330
         Width           =   5385
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
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
      Height          =   2565
      Left            =   6240
      TabIndex        =   7
      Top             =   60
      Width           =   5055
      Begin VB.OptionButton opIniWIN 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "Imagen original"
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
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   1830
         Width           =   1875
      End
      Begin VB.OptionButton OpIni3PM 
         BackColor       =   &H00000000&
         Caption         =   "Imagen de 3PM"
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
         Height          =   315
         Left            =   1350
         TabIndex        =   8
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label lblNoImgIni 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Imagen no disponible. De todas formas se puede habilitar"
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
         Left            =   3630
         TabIndex        =   23
         Top             =   660
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H00533422&
         Caption         =   "Imagen Inicio de Windows"
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
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   330
         Width           =   4785
      End
      Begin VB.Image imgIniWin 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   3630
         Stretch         =   -1  'True
         Top             =   660
         Width           =   1200
      End
      Begin VB.Image imgIni3PM 
         BorderStyle     =   1  'Fixed Single
         Height          =   1500
         Left            =   90
         Stretch         =   -1  'True
         Top             =   690
         Width           =   1200
      End
   End
   Begin VB.TextBox Text1 
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
      Index           =   2
      Left            =   1950
      TabIndex        =   2
      Text            =   "c:\windows\logos.sys"
      Top             =   7710
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.TextBox Text1 
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
      Index           =   1
      Left            =   1950
      TabIndex        =   1
      Text            =   "c:\windows\logow.sys"
      Top             =   7410
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.TextBox Text1 
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
      Index           =   0
      Left            =   1950
      TabIndex        =   0
      Text            =   "c:\logo.sys"
      Top             =   7110
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmINI3PM.frx":04EC
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
      Height          =   1245
      Index           =   4
      Left            =   900
      TabIndex        =   26
      Top             =   735
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "La imagen debe ser de 320 x 400 px y con 256 colores (paleta de 8 bits)."
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
      Height          =   825
      Index           =   3
      Left            =   4020
      TabIndex        =   6
      Top             =   7170
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Puede Apagar"
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
      Height          =   225
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   7740
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Se esta cerrando"
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
      Height          =   225
      Index           =   1
      Left            =   60
      TabIndex        =   4
      Top             =   7440
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   225
      Index           =   0
      Left            =   480
      TabIndex        =   3
      Top             =   7140
      Visible         =   0   'False
      Width           =   1425
   End
End
Attribute VB_Name = "frmINI3PM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MostrarMensajeWarning As Boolean

Private Sub Command1_Click()
    On Error GoTo MiErr
    'leer el archivo de configuracion para guardar los datos
    tERR.Anotar "acnv"
    Set TE = fso.OpenTextFile(GPF("iit17222"))
    Dim Ls() As String, C As Long
    C = 1
    Do While Not TE.AtEndOfStream
       
       ReDim Preserve Ls(C)
       Ls(C) = TE.ReadLine
       tERR.Anotar "acnw", C, Ls(C)
       C = C + 1
    Loop
    TE.Close
    If opIniWIN Then
        tERR.Anotar "acnx"
        Ls(4) = "LoadImgIni=w"
        If fso.FileExists("c:\logo.sys") Then fso.DeleteFile "c:\logo.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(1), 1, "=") = "1" Then
            'volver a cargarla
            fso.CopyFile GPF("ildw9m"), "c:\logo.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    
    If OpIni3PM Then
        tERR.Anotar "acny"
        Ls(4) = "LoadImgIni=3"
        If fso.FileExists("c:\logo.sys") Then fso.DeleteFile "c:\logo.sys", True
        'volver a cargarla
        fso.CopyFile GPF("ild3pm"), "c:\logo.sys", True
    End If
    
    If OpCerrandoWIN Then
        Ls(5) = "LoadImgCerrando=w"
        If fso.FileExists(WINfolder + "logow.sys") Then fso.DeleteFile WINfolder + "logow.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(2), 1, "=") = "1" Then
            'volver a cargarla
            fso.CopyFile GPF("ildw9m3"), WINfolder + "logow.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    tERR.Anotar "acnz"
    If OpCerrando3PM Then
        Ls(5) = "LoadImgCerrando=3"
        If fso.FileExists(WINfolder + "logow.sys") Then fso.DeleteFile WINfolder + "logow.sys", True
        'volver a cargarla
        fso.CopyFile GPF("ild3pm3"), WINfolder + "logow.sys", True
    End If
    
    If OpApagarWIN Then
        Ls(6) = "LoadImgApagar=w"
        tERR.Anotar "acoa"
        If fso.FileExists(WINfolder + "logos.sys") Then fso.DeleteFile WINfolder + "logos.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(3), 1, "=") = "1" Then
            'volver a cargarla
            fso.CopyFile GPF("ildw9m2"), WINfolder + "logos.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    
    If OpApagar3PM Then
        Ls(6) = "LoadImgApagar=3"
        If fso.FileExists(WINfolder + "logos.sys") Then fso.DeleteFile WINfolder + "logos.sys", True
        'volver a cargarla
        fso.CopyFile GPF("ild3pm2"), WINfolder + "logos.sys", True
    End If
        
    'volver a escribir el archivo de inicio con los nuevos datos
    fso.DeleteFile (GPF("iit17222"))
    Set TE = fso.CreateTextFile(GPF("iit17222"), True)
    TE.WriteLine Ls(1)
    TE.WriteLine Ls(2)
    TE.WriteLine Ls(3)
    TE.WriteLine Ls(4)
    TE.WriteLine Ls(5)
    TE.WriteLine Ls(6)
    TE.Close
    
    tERR.Anotar "acob"
    'leer el system.ini y ver si estamos con PROGMAN o EXPLORER
    'copiarlo para no echar moco
    If fso.FileExists(AP + "system.ini") Then fso.DeleteFile AP + "system.ini", True
    fso.CopyFile WINfolder + "system.ini", AP + "system.ini", True
    Set TE = fso.OpenTextFile(AP + "system.ini")
    Dim TodoSystem() As String
    Dim ActualShell As String, UbicShell As Long
    C = 1
    Do While Not TE.AtEndOfStream
        ReDim Preserve TodoSystem(C)
        TodoSystem(C) = TE.ReadLine
        tERR.Anotar "acoc", C, TodoSystem(C)
        If LCase(txtInLista(TodoSystem(C), 0, "=")) = "shell" Then
            UbicShell = C
            ActualShell = txtInLista(TodoSystem(C), 1, "=")
            'no salir para que se copie todo
        End If
        C = C + 1
    Loop
    TE.Close
    
    If Option6 Then TodoSystem(UbicShell) = "Shell=explorer.exe"
    If Option2 Then TodoSystem(UbicShell) = "Shell=progman.exe"
    'volver a escribir el archivo
    If fso.FileExists(AP + "system.ini") Then fso.DeleteFile AP + "system.ini", True
    Set TE = fso.CreateTextFile(AP + "system.ini", True)
    For A = 1 To UBound(TodoSystem)
        TE.WriteLine TodoSystem(A)
    Next
    TE.Close
    If fso.FileExists(WINfolder + "OLDsystem.ini") Then fso.DeleteFile WINfolder + "OLDsystem.ini", True
    If fso.FileExists(WINfolder + "system.ini") Then fso.MoveFile WINfolder + "system.ini", WINfolder + "OLDsystem.ini"
    fso.MoveFile AP + "system.ini", WINfolder + "system.ini"
    tERR.Anotar "acod"
    Unload Me
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acnu"
    Resume Next
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            Unload Me
            YaCerrar3PM
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            SetKeyState vbKeyScrollLock, True
            VarCreditos CSng(TemasPorCredito)
        Else
            'apagar el fichero electronico
            SetKeyState vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    On Error GoTo MiErr
    tERR.Anotar "acoe"
    MostrarMensajeWarning = False
    AjustarFRM Me, 12000, 9000
    
    'cargar las imágenes de 3pm
    If fso.FileExists(GPF("ild3pm")) Then imgIni3PM.Picture = LoadPicture(GPF("ild3pm"))
    If fso.FileExists(GPF("ild3pm3")) Then imgCerrando3PM.Picture = LoadPicture(GPF("ild3pm3"))
    If fso.FileExists(GPF("ild3pm2")) Then imgApagar3PM.Picture = LoadPicture(GPF("ild3pm2"))
    
    'cargar las imágenes de windows
    If fso.FileExists(GPF("ildw9m")) Then
        imgIniWin.Picture = LoadPicture(GPF("ildw9m"))
    Else
        lblNoImgIni.Visible = True
    End If
    
    If fso.FileExists(GPF("ildw9m3")) Then
        imgCerrandoWIN.Picture = LoadPicture(GPF("ildw9m3"))
    Else
        lblNoImgCerrando.Visible = True
    End If
    tERR.Anotar "acof"
    If fso.FileExists(GPF("ildw9m2")) Then
        imgApagarWIN.Picture = LoadPicture(GPF("ildw9m2"))
    Else
        lblNoImgApagar.Visible = True
    End If
    
    'leer el archivo de configuracion para saber
    'cual imagen se esta usando
    
    Set TE = fso.OpenTextFile(GPF("iit17222"))
    Dim Ls() As String, C As Long
    C = 1
    Do While Not TE.AtEndOfStream
        ReDim Preserve Ls(C)
        Ls(C) = TE.ReadLine
        tERR.Anotar "acog", C, Ls(C)
        C = C + 1
    Loop
    TE.Close
    Dim LoadImgIni As String
    LoadImgIni = txtInLista(Ls(4), 1, "=")
    Dim LoadImgCerrando As String
    LoadImgCerrando = txtInLista(Ls(5), 1, "=")
    Dim LoadImgApagar As String
    LoadImgApagar = txtInLista(Ls(6), 1, "=")
    tERR.Anotar "acoh"
    If LoadImgIni = "w" Then opIniWIN = True
    If LoadImgIni = "3" Then OpIni3PM = True
    
    If LoadImgCerrando = "w" Then OpCerrandoWIN = True
    If LoadImgCerrando = "3" Then OpCerrando3PM = True
    
    If LoadImgApagar = "w" Then OpApagarWIN = True
    If LoadImgApagar = "3" Then OpApagar3PM = True
    
    'leer el system.ini y ver si estamos con PROGMAN o EXPLORER
    'copiarlo para no echar moco
    If fso.FileExists(AP + "system.ini") Then fso.DeleteFile AP + "system.ini", True
    fso.CopyFile WINfolder + "system.ini", AP + "system.ini", True
    Set TE = fso.OpenTextFile(AP + "system.ini")
    Dim TodoSystem() As String
    Dim ActualShell As String, UbicShell As Long
    C = 1
    
    Do While Not TE.AtEndOfStream
        ReDim Preserve TodoSystem(C)
        TodoSystem(C) = TE.ReadLine
        tERR.Anotar "acoi", C, TodoSystem(C)
        If LCase(txtInLista(TodoSystem(C), 0, "=")) = "shell" Then
            UbicShell = C
            ActualShell = txtInLista(TodoSystem(C), 1, "=")
            Exit Do
        End If
        C = C + 1
    Loop
    TE.Close
    If UCase(ActualShell) = "EXPLORER.EXE" Then Option6 = True
    If UCase(ActualShell) = "PROGMAN.EXE" Then Option2 = True
    tERR.Anotar "acoj"
    MostrarMensajeWarning = True
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acnt2"
    Resume Next
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub Option2_Click()
    If MostrarMensajeWarning Then
        'MsgBox "¡¡¡Solo active esta opción sobre " + _
        tr.trad("Windows 98. Puede no ser compatible ") + _
        tr.trad("con otros sistemas operativos!!!")
    End If
End Sub

'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command1.Caption = TR.Trad("Grabar configuración%99%")
    Command2.Caption = TR.Trad("Salir sin grabar%99%")
    OpApagar3PM.Caption = TR.Trad("Imagen de 3PM%99%")
    OpApagarWIN.Caption = TR.Trad("Imagen original%99%")
    lblNoImgApagar.Caption = TR.Trad("Imagen no disponible. De todas formas " + _
        "se puede habilitar%99%")
    Label2(2).Caption = TR.Trad("Imagen 'Ahora puede apagar el equipo'%99%")
    OpCerrando3PM.Caption = TR.Trad("Imagen de 3PM%99%")
    OpCerrandoWIN.Caption = TR.Trad("Imagen original%99%")
    lblNoImgCerrando.Caption = TR.Trad("Imagen no disponible. De todas formas " + _
        "se puede habilitar%99%")
    Label2(0).Caption = TR.Trad("Imagen 'Windows se esta cerrando'%99%")
    Option2.Caption = TR.Trad("3PM se carga siempre al inicio omitiendo la " + _
        "carga del entorno de exploracion de Windows. Carga mínima de " + _
        "Windows. Recomendado para la instalación de 3PM en la fonola%99%")
    Option6.Caption = TR.Trad("No cargar 3PM al iniciar Windows. Recomendado " + _
        "si el equipo no es de uso exclusivo de 3PM%99%")
    Label2(3).Caption = TR.Trad("Cargar 3PM al Iniciar el equipo%99%")
    opIniWIN.Caption = TR.Trad("Imagen original%99%")
    OpIni3PM.Caption = TR.Trad("Imagen de 3PM%99%")
    lblNoImgIni.Caption = TR.Trad("Imagen no disponible. De todas formas se " + _
        "puede habilitar%99%")
    Label2(1).Caption = TR.Trad("Imagen de Inicio de Windows%99%")
    Label1(4).Caption = TR.Trad("De esta página de Configuración podrá " + _
        "administrar el modo de inicio de Windows así como también las " + _
        "imagenes mostrar al iniciar y cerrar el sistema. (solo Windows 98/Me).%99%")
    Label1(3).Caption = TR.Trad("La imagen debe ser de 320 x 400 y con " + _
        "256 colores (paleta de 8 bits)%99%")
    Label1(2).Caption = TR.Trad("Puede Apagar%99%")
    Label1(1).Caption = TR.Trad("Se esta cerrando%99%")
    Label1(0).Caption = TR.Trad("Inicio%99%")
End Sub
