VERSION 5.00
Begin VB.Form frmREPAIR 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reparar 3PM"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7380
   Icon            =   "frmREPAIR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "CERRAR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1560
      TabIndex        =   4
      Top             =   2250
      Width           =   2745
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reparar 3PM"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1590
      TabIndex        =   1
      Top             =   1050
      Width           =   2745
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   29
      Left            =   6420
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   28
      Left            =   6660
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   27
      Left            =   6900
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   26
      Left            =   7140
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   25
      Left            =   5220
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   24
      Left            =   5460
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   23
      Left            =   5700
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   22
      Left            =   5940
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   21
      Left            =   6180
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   20
      Left            =   4950
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   19
      Left            =   3750
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   18
      Left            =   3990
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   17
      Left            =   4230
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   16
      Left            =   4470
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   15
      Left            =   4710
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   14
      Left            =   2520
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   13
      Left            =   2760
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   12
      Left            =   3000
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   11
      Left            =   3240
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   10
      Left            =   3480
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   9
      Left            =   1290
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   8
      Left            =   1530
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   7
      Left            =   1770
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   6
      Left            =   2010
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   5
      Left            =   2250
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   4
      Left            =   1020
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   3
      Left            =   780
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   2
      Left            =   540
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   1
      Left            =   300
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Shape OKdelete 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   195
      Index           =   0
      Left            =   60
      Shape           =   3  'Circle
      Top             =   1560
      Width           =   225
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREPAIR.frx":0442
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
      Height          =   765
      Left            =   120
      TabIndex        =   3
      Top             =   150
      Width           =   7215
   End
   Begin VB.Label lblPBAR 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   1770
      Width           =   15
   End
   Begin VB.Label lblBAR 
      BackColor       =   &H000000C0&
      Height          =   375
      Left            =   60
      TabIndex        =   0
      Top             =   1770
      Width           =   7275
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
Dim FSO As New Scripting.FileSystemObject

Private Sub Command1_Click()
    'borrar todos los archivos que no se instalaron y _
        que forman parte de 3PM
    
    'primero RMLVF.DLL (indicador de licencia)
    'tbrDelete SysFolder + "\rmlvf.dll", 10, 0
    'tbrDelete SysFolder + "\rmlvf.tlb", 15, 1
    
    'origernes de discos
    tbrDelete SysFolder + "oddtb.jut", 10, 0
        
    'En Frm Reg una chiquita tipo la chica = index = tapa _
        f61.dlw
    tbrDelete SysFolder + "f61.dlw", 15, 1
    
    'En frmIni: _
        una grande: f52.dlw
    tbrDelete SysFolder + "f52.dlw", 20, 2
    
    'En frmIndex se necesita _
        'El fondo grande: f53.dlw
    tbrDelete SysFolder + "f53.dlw", 25, 3
        'El fondo chico de abajo: f55.dlw (para exclusivo el mismo!!!)
    tbrDelete SysFolder + "f55.dlw", 30, 4
        'tbrPassImg: es el mismo f61.dlw !!!
    'en frmTop10-RANK: el mismo f61.dlw
    
    'En frmSuperLic se necesitan: _
        los 3 archivos de Windows _
        logo.sys = f56.dlw
    tbrDelete SysFolder + "f56.dlw", 35, 5
        'logos.sys = f57.dlw
    tbrDelete SysFolder + "f57.dlw", 40, 6
        'logow.sys = f58.dlw
    tbrDelete SysFolder + "f58.dlw", 45, 7
        'las imagenes del frmINI _
        f52.dlw _
        Imagen del index en tbrPassIMG _
        tapa.jpg = f61.dlw _
        TOP10.jpg = f54.dlw
    tbrDelete SysFolder + "f54.dlw", 50, 8
            
    'claves de creditos gratuitos
    tbrDelete WinFolder + "sevalc.dll", 55, 9
    'imagen grande al inicio
    tbrDelete WinFolder + "SL\imgbig.tbr", 60, 10
    'imagen tbr al inicio
    tbrDelete WinFolder + "SL\imgtbr.tbr", 65, 11
    'logito en principal (index)
    tbrDelete WinFolder + "SL\indexchi.tbr", 70, 12
    ' texto principal en index
    tbrDelete WinFolder + "SL\txtIDX.tbr", 75, 13
    'texto en configuracion
    tbrDelete WinFolder + "SL\txtCFG.tbr", 80, 14
    'azar guid si es K5
    tbrDelete SysFolder + "razaGUID.dll", 85, 15
    'codigo a pedir en validacio
    tbrDelete SysFolder + "codped.cfg", 87, 16
    'contador de usos y fechas
    tbrDelete SysFolder + "daily.cfg", 88, 17
    'configuracion
    tbrDelete SysFolder + "3pmcfg.tbr", 90, 18
    'creditos en validacion
    tbrDelete SysFolder + "radilav.cfg", 91, 19
    'ranking
    tbrDelete AP + "ranking.tbr", 92, 20
    'temas a reinicar ejecutando
    tbrDelete AP + "reini.tbr", 93, 21
    'creditos
    tbrDelete AP + "creditos.tbr", 94, 22
    'temporal de la config
    tbrDelete AP + "tmp.tbr", 94, 23
    'log error
    tbrDelete AP + "TBRlog.txt", 95, 24
    'error acumulado
    tbrDelete AP + "OLDtbrlog.txt", 96, 25
    'log error
    tbrDelete AP + "log.txt", 97, 26
    'error acumulado
    tbrDelete AP + "OLDlog.txt", 98, 27
    'protector de pantalla
    tbrDelete AP + "protect.tbr", 99, 28
    'imagenes con que se inicia
    tbrDelete AP + "imgini.tbr", 100, 29
    
    MsgBox "Se ha terminado la reparacion"
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    AP = App.Path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    SysFolder = FSO.GetSpecialFolder(SystemFolder)
    WinFolder = FSO.GetSpecialFolder(WindowsFolder)
    If Right(SysFolder, 1) <> "\" Then SysFolder = SysFolder + "\"
    If Right(WinFolder, 1) <> "\" Then WinFolder = WinFolder + "\"
End Sub

Public Function tbrDelete(Arch As String, PorcPasado As Long, IndiceBola As Integer) As Boolean
    If FSO.FileExists(Arch) Then
        FSO.DeleteFile Arch, True
        OKdelete(IndiceBola).BackColor = vbGreen
        tbrDelete = True
    Else
        OKdelete(IndiceBola).BackColor = vbRed
        tbrDelete = False
    End If
    lblPBAR.Width = lblBAR.Width * PorcPasado / 100
End Function

