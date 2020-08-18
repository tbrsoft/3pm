VERSION 5.00
Begin VB.Form frmREPAIR 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reparar 3PM"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5520
   Icon            =   "frmREPAIR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDO 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   10
      Text            =   "frmREPAIR.frx":0442
      Top             =   4020
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   3630
      Width           =   2595
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reiniciar Archivos. (tecla DER (x) de fonola)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   1
      Left            =   2850
      TabIndex        =   4
      Top             =   2460
      Width           =   2625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Recuperar archivos (tecla IZQ (z) de fonola)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   2460
      Width           =   2595
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   5220
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "REPARAR BORRANDO"
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
      Height          =   225
      Index           =   2
      Left            =   2910
      TabIndex        =   8
      Top             =   900
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "RECUPERAR ARCHIVOS"
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
      Height          =   225
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   900
      Width           =   2565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmREPAIR.frx":0448
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
      Index           =   1
      Left            =   2850
      TabIndex        =   6
      Top             =   1110
      Width           =   2595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmREPAIR.frx":04E0
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
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Top             =   1110
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Herramienta de recuperacion y reparacion de archivos externos de 3PM."
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
      Height          =   465
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   90
      Width           =   5115
   End
   Begin VB.Label lblPBAR 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   3180
      Width           =   15
   End
   Begin VB.Label lblBAR 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   3180
      Width           =   4965
   End
End
Attribute VB_Name = "frmREPAIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AP As String
Dim WinFolder As String
Dim SysFolder As String
Dim Fso As New Scripting.FileSystemObject

Private Sub Command1_Click(Index As Integer)
    'borrar todos los archivos que no se instalaron y _
        que forman parte de 3PM. Buscar copias de seguridad antes
    
    '****************************
    'parte de recuperacion
    If Index = 0 Then
        'ARCHIVO LICENCIA
        'ORIGINAL: SYSfolder + "dciLib22.dll"
        'COPIA EN: SysFolder "c2LK.dll"
        'lo de la licencia va mas alla de cualuier opción
        If Fso.FileExists(SysFolder + "c2LK.dll") Then
            txtDO = txtDO + vbCrLf + "RECUPERADA LICENCIA!!"
            Fso.CopyFile SysFolder + "c2LK.dll", SysFolder + "dciLib22.dll", True
        End If
        '-------------------------------------
        'CONFIGURACION:
        'ORIGINAL: SYSfolder + "3pmcfg.tbr"
        'COPIA: SYSfolder + "autoSave3PM.cfg"
        tbrDelete SysFolder + "3pmcfg.tbr", 10
        If Index = 0 Then
            If Fso.FileExists(SysFolder + "autoSave3PM.cfg") Then
                Fso.CopyFile SysFolder + "autoSave3PM.cfg", SysFolder + "3pmcfg.tbr", True
                txtDO = txtDO + vbCrLf + "RECUPERADA CONFGURACION!!"
            End If
        End If
        GoTo FIN
        '-------------------------------------
    Else
        'si no los borra
        tbrDelete SysFolder + "dciLib22.dll", 5
        tbrDelete SysFolder + "3pmcfg.tbr", 10
    End If
    'ORIGENES DISCOS:
    'ORIGINAL: SYSfolder + "oddtb.jut"
    'COPIA: NO HAY
    tbrDelete SysFolder + "oddtb.jut", 15
       
    'En Frm Reg una chiquita tipo la chica = index = tapa _
        f61.dlw
    tbrDelete SysFolder + "f61.dlw", 17
    
    'En frmIni: _
        una grande: f52.dlw
    tbrDelete SysFolder + "f52.dlw", 20
    'En frmIndex se necesita _
        'El fondo grande: f53.dlw
    tbrDelete SysFolder + "f53.dlw", 25
        'El fondo chico de abajo: f55.dlw (para exclusivo el mismo!!!)
    tbrDelete SysFolder + "f55.dlw", 30
        'tbrPassImg: es el mismo f61.dlw !!!
    'en frmTop10-RANK: el mismo f61.dlw
    
    'En frmSuperLic se necesitan: _
        los 3 archivos de Windows _
        logo.sys = f56.dlw
    tbrDelete SysFolder + "f56.dlw", 35
        'logos.sys = f57.dlw
    tbrDelete SysFolder + "f57.dlw", 40
        'logow.sys = f58.dlw
    tbrDelete SysFolder + "f58.dlw", 45
        'las imagenes del frmINI _
        f52.dlw _
        Imagen del index en tbrPassIMG _
        tapa.jpg = f61.dlw _
        TOP10.jpg = f54.dlw
    tbrDelete SysFolder + "f54.dlw", 50
            
    'claves de creditos gratuitos
    tbrDelete WinFolder + "sevalc.dll", 55
    'imagen grande al inicio
    tbrDelete WinFolder + "SL\imgbig.tbr", 60
    'imagen tbr al inicio
    tbrDelete WinFolder + "SL\imgtbr.tbr", 65
    'logito en principal (index)
    tbrDelete WinFolder + "SL\indexchi.tbr", 70
    ' texto principal en index
    tbrDelete WinFolder + "SL\txtIDX.tbr", 75
    'texto en configuracion
    tbrDelete WinFolder + "SL\txtCFG.tbr", 80
    'azar guid si es K5
    tbrDelete SysFolder + "razaGUID.dll", 85
    'codigo a pedir en validacion
    tbrDelete SysFolder + "codped.cfg", 87
    'contador de usos y fechas
    tbrDelete SysFolder + "daily.cfg", 88
    
    'creditos en validacion
    tbrDelete SysFolder + "radilav.cfg", 91
    'ranking
    tbrDelete AP + "ranking.tbr", 92
    'temas a reinicar ejecutando
    tbrDelete AP + "reini.tbr", 93
    'creditos
    tbrDelete AP + "creditos.tbr", 94
    'temporal de la config
    tbrDelete AP + "tmp.tbr", 94
    'protector de pantalla
    tbrDelete AP + "protect.tbr", 99
    'imagenes con que se inicia
    tbrDelete AP + "imgini.tbr", 100

FIN:
    txtDO = txtDO + vbCrLf + "Se ha terminado la reparacion. Se iniciara 3PM" + vbCrLf + "espere..."
    txtDO.Refresh
    Dim T As Single
    T = Timer
    Do While T + 2 > Timer
    
    Loop
    
    
    Shell App.path + "\3pm.exe"
    Unload Me
    End
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyZ Then Command1_Click 0
    If KeyCode = vbKeyX Then Command1_Click 1
End Sub

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    SysFolder = Fso.GetSpecialFolder(SystemFolder)
    WinFolder = Fso.GetSpecialFolder(WindowsFolder)
    If Right(SysFolder, 1) <> "\" Then SysFolder = SysFolder + "\"
    If Right(WinFolder, 1) <> "\" Then WinFolder = WinFolder + "\"
    txtDO = "Acciones realizadas:"
End Sub

Public Function tbrDelete(Arch As String, PorcPasado As Long) As Boolean
    If Fso.FileExists(Arch) Then
        Fso.DeleteFile Arch, True
        tbrDelete = True
    Else
        tbrDelete = False
    End If
    lblPBAR.Width = lblBAR.Width * PorcPasado / 100
    txtDO = txtDO + vbCrLf + "BORRADO...(" + CStr(PorcPasado) + ")"
End Function

