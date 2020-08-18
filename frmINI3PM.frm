VERSION 5.00
Begin VB.Form frmINI3PM 
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Inicio de 3PM"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   Icon            =   "frmINI3PM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grabar configuracion"
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
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5550
      Width           =   2350
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salir sin grabar"
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
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6000
      Width           =   2350
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00400000&
      Height          =   2475
      Left            =   6240
      TabIndex        =   19
      Top             =   5250
      Width           =   5055
      Begin VB.OptionButton OpApagar3PM 
         BackColor       =   &H00400000&
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
         BackColor       =   &H00400000&
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
         BackColor       =   &H00404000&
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
         BackColor       =   &H00808000&
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
      BackColor       =   &H00400000&
      Height          =   2475
      Left            =   6240
      TabIndex        =   15
      Top             =   2700
      Width           =   5055
      Begin VB.OptionButton OpCerrando3PM 
         BackColor       =   &H00400000&
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
         BackColor       =   &H00400000&
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
         BackColor       =   &H00404000&
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
         BackColor       =   &H00808000&
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
      BackColor       =   &H00400000&
      Height          =   2625
      Left            =   150
      TabIndex        =   11
      Top             =   2400
      Width           =   5805
      Begin VB.OptionButton Option2 
         BackColor       =   &H00400000&
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
         Top             =   1410
         Width           =   5385
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00400000&
         Caption         =   "No cargar 3PM al iniciar Windows. Recomendado si el equipo no es de uso exclusivo de 3PM"
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
         BackColor       =   &H00808000&
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
      BackColor       =   &H00400000&
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
         BackColor       =   &H00400000&
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
         BackColor       =   &H00400000&
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
         BackColor       =   &H00404000&
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
         BackColor       =   &H00808000&
         Caption         =   "Imagen de Inicio de Windows"
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
      Caption         =   $"frmINI3PM.frx":04EB
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
      Height          =   1065
      Index           =   4
      Left            =   900
      TabIndex        =   26
      Top             =   990
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "La imagen debe ser de 320 x 400 y con 256 colores (paleta de 8 bits)"
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
      Height          =   825
      Index           =   3
      Left            =   4020
      TabIndex        =   6
      Top             =   7170
      Visible         =   0   'False
      Width           =   1815
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
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
    'leer el archivo de configuracion para guardar
    'los datos
    
    Set TE = FSO.OpenTextFile(AP + "imgini.tbr")
    Dim Ls() As String, c As Long
    c = 1
    Do While Not TE.AtEndOfStream
       ReDim Preserve Ls(c)
       Ls(c) = TE.ReadLine
       c = c + 1
    Loop
    TE.Close
    If opIniWIN Then
        Ls(4) = "LoadImgIni=w"
        If FSO.FileExists("c:\logo.sys") Then FSO.DeleteFile "c:\logo.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(1), 1, "=") = "1" Then
            'volver a cargarla
            FSO.CopyFile WINfolder + "img3pm\w\logo.sys", "c:\logo.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    
    If OpIni3PM Then
        Ls(4) = "LoadImgIni=3"
        If FSO.FileExists("c:\logo.sys") Then FSO.DeleteFile "c:\logo.sys", True
        'volver a cargarla
        FSO.CopyFile WINfolder + "img3pm\3\logo.sys", "c:\logo.sys", True
    End If
    
    If OpCerrandoWIN Then
        Ls(5) = "LoadImgCerrando=w"
        If FSO.FileExists(WINfolder + "logow.sys") Then FSO.DeleteFile WINfolder + "logow.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(2), 1, "=") = "1" Then
            'volver a cargarla
            FSO.CopyFile WINfolder + "img3pm\w\logow.sys", WINfolder + "logow.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    
    If OpCerrando3PM Then
        Ls(5) = "LoadImgCerrando=3"
        If FSO.FileExists(WINfolder + "logow.sys") Then FSO.DeleteFile WINfolder + "logow.sys", True
        'volver a cargarla
        FSO.CopyFile WINfolder + "img3pm\3\logow.sys", WINfolder + "logow.sys", True
    End If
    
    If OpApagarWIN Then
        Ls(6) = "LoadImgApagar=w"
        If FSO.FileExists(WINfolder + "logos.sys") Then FSO.DeleteFile WINfolder + "logos.sys", True
        'ver si la imagen estaba
        If txtInLista(Ls(3), 1, "=") = "1" Then
            'volver a cargarla
            FSO.CopyFile WINfolder + "img3pm\w\logos.sys", WINfolder + "logos.sys", True
        Else
            'como no estaba se queda sin imagen
        End If
    End If
    
    If OpApagar3PM Then
        Ls(6) = "LoadImgApagar=3"
        If FSO.FileExists(WINfolder + "logos.sys") Then FSO.DeleteFile WINfolder + "logos.sys", True
        'volver a cargarla
        FSO.CopyFile WINfolder + "img3pm\3\logos.sys", WINfolder + "logos.sys", True
    End If
        
    'volver a escribir el archivo de inicio con los nuevos datos
    FSO.DeleteFile (AP + "imgini.tbr")
    Set TE = FSO.CreateTextFile(AP + "imgini.tbr", True)
    TE.WriteLine Ls(1)
    TE.WriteLine Ls(2)
    TE.WriteLine Ls(3)
    TE.WriteLine Ls(4)
    TE.WriteLine Ls(5)
    TE.WriteLine Ls(6)
    TE.Close
    
    'leer el system.ini y ver si estamos con PROGMAN o EXPLORER
    'copiarlo para no echar moco
    If FSO.FileExists(AP + "system.ini") Then FSO.DeleteFile AP + "system.ini", True
    FSO.CopyFile WINfolder + "system.ini", AP + "system.ini", True
    Set TE = FSO.OpenTextFile(AP + "system.ini")
    Dim TodoSystem() As String
    Dim ActualShell As String, UbicShell As Long
    c = 1
    Do While Not TE.AtEndOfStream
        ReDim Preserve TodoSystem(c)
        TodoSystem(c) = TE.ReadLine
        If LCase(txtInLista(TodoSystem(c), 0, "=")) = "shell" Then
            UbicShell = c
            ActualShell = txtInLista(TodoSystem(c), 1, "=")
            'no salir para que se copie todo
        End If
        c = c + 1
    Loop
    TE.Close
    If Option6 Then TodoSystem(UbicShell) = "Shell=explorer.exe"
    If Option2 Then TodoSystem(UbicShell) = "Shell=progman.exe"
    'volver a escribir el archivo
    If FSO.FileExists(AP + "system.ini") Then FSO.DeleteFile AP + "system.ini", True
    Set TE = FSO.CreateTextFile(AP + "system.ini", True)
    For A = 1 To UBound(TodoSystem)
        TE.WriteLine TodoSystem(A)
    Next
    TE.Close
    If FSO.FileExists(WINfolder + "OLDsystem.ini") Then FSO.DeleteFile WINfolder + "OLDsystem.ini", True
    If FSO.FileExists(WINfolder + "system.ini") Then FSO.MoveFile WINfolder + "system.ini", WINfolder + "OLDsystem.ini"
    FSO.MoveFile AP + "system.ini", WINfolder + "system.ini"
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
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
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            
            ShowCredits
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
            
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    MostrarMensajeWarning = False
    AjustarFRM Me, 12000
    
    'cargar las imágenes de 3pm
    If FSO.FileExists(WINfolder + "img3pm\3\logo.sys") Then imgIni3PM.Picture = LoadPicture(WINfolder + "img3pm\3\logo.sys")
    If FSO.FileExists(WINfolder + "img3pm\3\logow.sys") Then imgCerrando3PM.Picture = LoadPicture(WINfolder + "img3pm\3\logow.sys")
    If FSO.FileExists(WINfolder + "img3pm\3\logos.sys") Then imgApagar3PM.Picture = LoadPicture(WINfolder + "img3pm\3\logos.sys")
    
    'cargar las imágenes de windows
    If FSO.FileExists(WINfolder + "img3pm\w\logo.sys") Then
        imgIniWin.Picture = LoadPicture(WINfolder + "img3pm\w\logo.sys")
    Else
        lblNoImgIni.Visible = True
    End If
    
    If FSO.FileExists(WINfolder + "img3pm\w\logow.sys") Then
        imgCerrandoWIN.Picture = LoadPicture(WINfolder + "img3pm\w\logow.sys")
    Else
        lblNoImgCerrando.Visible = True
    End If
    
    If FSO.FileExists(WINfolder + "img3pm\w\logos.sys") Then
        imgApagarWIN.Picture = LoadPicture(WINfolder + "img3pm\w\logos.sys")
    Else
        lblNoImgApagar.Visible = True
    End If
    
    'leer el archivo de configuracion para saber
    'cual imagen se esta usando
    
    Set TE = FSO.OpenTextFile(AP + "imgini.tbr")
    Dim Ls() As String, c As Long
    c = 1
    Do While Not TE.AtEndOfStream
       ReDim Preserve Ls(c)
       Ls(c) = TE.ReadLine
       c = c + 1
    Loop
    TE.Close
    Dim LoadImgIni As String
    LoadImgIni = txtInLista(Ls(4), 1, "=")
    Dim LoadImgCerrando As String
    LoadImgCerrando = txtInLista(Ls(5), 1, "=")
    Dim LoadImgApagar As String
    LoadImgApagar = txtInLista(Ls(6), 1, "=")
    
    If LoadImgIni = "w" Then opIniWIN = True
    If LoadImgIni = "3" Then OpIni3PM = True
    
    If LoadImgCerrando = "w" Then OpCerrandoWIN = True
    If LoadImgCerrando = "3" Then OpCerrando3PM = True
    
    If LoadImgApagar = "w" Then OpApagarWIN = True
    If LoadImgApagar = "3" Then OpApagar3PM = True
    
    'leer el system.ini y ver si estamos con PROGMAN o EXPLORER
    'copiarlo para no echar moco
    If FSO.FileExists(AP + "system.ini") Then FSO.DeleteFile AP + "system.ini", True
    FSO.CopyFile WINfolder + "system.ini", AP + "system.ini", True
    Set TE = FSO.OpenTextFile(AP + "system.ini")
    Dim TodoSystem() As String
    Dim ActualShell As String, UbicShell As Long
    c = 1
    
    Do While Not TE.AtEndOfStream
        ReDim Preserve TodoSystem(c)
        TodoSystem(c) = TE.ReadLine
        If LCase(txtInLista(TodoSystem(c), 0, "=")) = "shell" Then
            UbicShell = c
            ActualShell = txtInLista(TodoSystem(c), 1, "=")
            Exit Do
        End If
        c = c + 1
    Loop
    TE.Close
    If UCase(ActualShell) = "EXPLORER.EXE" Then Option6 = True
    If UCase(ActualShell) = "PROGMAN.EXE" Then Option2 = True
    
    MostrarMensajeWarning = True
End Sub

Private Sub Option2_Click()
    If MostrarMensajeWarning Then
        'MsgBox "¡¡¡Solo active esta opción sobre " + _
        "Windows 98. Puede no ser compatible " + _
        "con otros sistemas operativos!!!"
    End If
End Sub

