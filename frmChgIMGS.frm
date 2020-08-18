VERSION 5.00
Object = "{AC1ACB77-BE60-49F4-BE38-2F9A87F5E5E4}#2.0#0"; "tbrX_Boton II.ocx"
Begin VB.Form frmChgIMGS 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar componentes graficos"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Imagen de inicio de 3PM"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   270
      TabIndex        =   4
      Top             =   2280
      Width           =   4725
      Begin tbrX_Boton2.XxBoton Command23 
         Height          =   465
         Left            =   270
         TabIndex        =   5
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Elegir imagen"
         xEnabled        =   0   'False
      End
      Begin tbrX_Boton2.XxBoton XxBoton5 
         Height          =   465
         Left            =   270
         TabIndex        =   6
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Quitar"
         xEnabled        =   0   'False
      End
      Begin VB.Image IMG2 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Imagen de fondo de las portadas"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   270
      TabIndex        =   1
      Top             =   300
      Width           =   4725
      Begin tbrX_Boton2.XxBoton Command22 
         Height          =   465
         Left            =   270
         TabIndex        =   2
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Elegir imagen"
         xEnabled        =   0   'False
      End
      Begin tbrX_Boton2.XxBoton XxBoton4 
         Height          =   465
         Left            =   270
         TabIndex        =   3
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Quitar"
         xEnabled        =   0   'False
      End
      Begin VB.Image IMG1 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
   Begin tbrX_Boton2.XxBoton XxBoton1 
      Height          =   345
      Left            =   4290
      TabIndex        =   0
      Top             =   4260
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   609
      xFColor         =   16777215
      xBColor         =   64
      xCapt           =   "SALIR"
      xEnabled        =   -1  'True
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Tapa predetermianda de discos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   5220
      TabIndex        =   7
      Top             =   300
      Width           =   4725
      Begin tbrX_Boton2.XxBoton Command24 
         Height          =   465
         Left            =   270
         TabIndex        =   8
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Elegir imagen"
         xEnabled        =   0   'False
      End
      Begin tbrX_Boton2.XxBoton XxBoton6 
         Height          =   465
         Left            =   270
         TabIndex        =   9
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Quitar"
         xEnabled        =   0   'False
      End
      Begin VB.Image IMG3 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2610
         Stretch         =   -1  'True
         Top             =   270
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000000&
      Caption         =   "Tapa predetermianda de ranking"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   5190
      TabIndex        =   10
      Top             =   2280
      Width           =   4725
      Begin tbrX_Boton2.XxBoton Command25 
         Height          =   465
         Left            =   270
         TabIndex        =   11
         Top             =   510
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Elegir imagen"
         xEnabled        =   0   'False
      End
      Begin tbrX_Boton2.XxBoton XxBoton7 
         Height          =   465
         Left            =   270
         TabIndex        =   12
         Top             =   990
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   820
         xFColor         =   16777215
         xBColor         =   6263909
         xCapt           =   "Quitar"
         xEnabled        =   0   'False
      End
      Begin VB.Image IMG4 
         BorderStyle     =   1  'Fixed Single
         Height          =   1440
         Left            =   2640
         Stretch         =   -1  'True
         Top             =   300
         Width           =   1995
      End
   End
End
Attribute VB_Name = "frmChgIMGS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetArch(sTitle As String, sFilter As String) As String
    
    Dim CM As New CommonDialog
    CM.DialogPrompt = sTitle
    CM.DialogTitle = sTitle
    CM.Filter = sFilter
    
    CM.ShowOpen
    Dim F As String
    F = CM.FileName
    
    Set CM = Nothing
    
    GetArch = F
    
End Function

Private Sub Form_Load()
    XxBoton4.Enabled = (K.LICENCIA = HSuperLicencia)
    Command22.Enabled = (K.LICENCIA = HSuperLicencia)
    
    XxBoton5.Enabled = (K.LICENCIA = HSuperLicencia)
    Command23.Enabled = (K.LICENCIA = HSuperLicencia)
    
    XxBoton6.Enabled = (K.LICENCIA = HSuperLicencia)
    Command24.Enabled = (K.LICENCIA = HSuperLicencia)
    
    XxBoton7.Enabled = (K.LICENCIA = HSuperLicencia)
    Command25.Enabled = (K.LICENCIA = HSuperLicencia)

    'mostrar las imagenes usadas actualemente
    If FSO.FileExists(GPF("iischu")) Then
        img1.Picture = LoadPicture(GPF("iischu"))
    Else
        img1.Picture = LoadPicture(ExtraData.GetImagePath("FondoDeLasTapas"))
    End If
    
    If FSO.FileExists(GPF("iisl67")) Then
        img2.Picture = LoadPicture(GPF("iisl67"))
    Else
        img2.Picture = LoadPicture(ExtraData.GetImagePath("iniciasys"))
    End If
    
    If FSO.FileExists(GPF("tddp322")) Then
        img3.Picture = LoadPicture(GPF("tddp322"))
    Else
        img3.Picture = LoadPicture(ExtraData.GetImagePath("tapapredeterminada"))
    End If
    
    If FSO.FileExists(GPF("tddp323")) Then
        IMG4.Picture = LoadPicture(GPF("tddp323"))
    Else
        IMG4.Picture = LoadPicture(ExtraData.GetImagePath("taparanking"))
    End If
    
End Sub

Private Sub XxBoton1_Click()
    Unload Me
End Sub

Private Sub Command22_Click()
    Dim F As String
    F = GetArch("Elegir imagen", "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    FSO.CopyFile F, GPF("iischu"), True
    img1.Picture = LoadPicture(GPF("iischu"))
End Sub

Private Sub XxBoton4_Click()
    FSO.DeleteFile GPF("iischu"), True
    'vuelve a la del skin original que si o si existe
    img1.Picture = LoadPicture(ExtraData.GetImagePath("FondoDeLasTapas"))
End Sub

Private Sub Command23_Click()
    Dim F As String
    F = GetArch("Elegir imagen", "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    'lo pongo en la ubicacion
    FSO.CopyFile F, GPF("iisl67"), True
    img2.Picture = LoadPicture(GPF("iisl67"))
End Sub

Private Sub XxBoton5_Click()
    FSO.DeleteFile GPF("iisl67"), True
    'vuelve a la del skin original que si o si existe
    img2.Picture = LoadPicture(ExtraData.GetImagePath("iniciasys"))
End Sub

Private Sub Command24_Click()
    Dim F As String
    F = GetArch("Elegir imagen", "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    FSO.CopyFile F, GPF("tddp322"), True
    img3.Picture = LoadPicture(GPF("tddp322"))
End Sub

Private Sub XxBoton6_Click()
    FSO.DeleteFile GPF("tddp322"), True
    'vuelve a la del skin original que si o si existe
    img3.Picture = LoadPicture(ExtraData.GetImagePath("tapapredeterminada"))
End Sub

Private Sub Command25_Click()
    Dim F As String
    F = GetArch("Elegir imagen", "Imagenes (*.jpg, *.bmp, *.gif)|*.jpg; *.gif;*.bmp")
    If F = "" Then Exit Sub
    'lo pongo en la ubicacion
    FSO.CopyFile F, GPF("tddp323"), True
    IMG4.Picture = LoadPicture(GPF("tddp323"))
End Sub

Private Sub XxBoton7_Click()
    FSO.DeleteFile GPF("tddp323"), True
    'vuelve a la del skin original que si o si existe
    IMG4.Picture = LoadPicture(ExtraData.GetImagePath("taparanking"))
End Sub


