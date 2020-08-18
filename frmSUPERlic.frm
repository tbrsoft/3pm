VERSION 5.00
Begin VB.Form frmSUPERlic 
   BackColor       =   &H00000080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7950
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5580
      Width           =   1215
   End
   Begin VB.CommandButton cmdImgQ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6180
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdImgQ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3750
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdImgQ 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2820
      Width           =   1215
   End
   Begin VB.TextBox lblTBR 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   20
      Text            =   "frmSUPERlic.frx":0000
      Top             =   4620
      Width           =   5865
   End
   Begin VB.TextBox txtCFG 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00C0C000&
      Height          =   2235
      Left            =   9000
      MultiLine       =   -1  'True
      TabIndex        =   19
      Text            =   "frmSUPERlic.frx":000A
      Top             =   5625
      Width           =   2865
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5220
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7950
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4650
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2070
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5190
      Width           =   1215
   End
   Begin VB.CommandButton cmdImgPresP 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2430
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   10650
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8190
      Width           =   1215
   End
   Begin VB.CommandButton cmdImg 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   4980
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdImg 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3120
      Width           =   1100
   End
   Begin VB.CommandButton cmdImg 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3120
      Width           =   1100
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Esta imagen se utilizara para: 1) La página principal. 2) La página de ranking. 3) Como portada de Cd predeterminada."
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
      Height          =   915
      Index           =   9
      Left            =   60
      TabIndex        =   21
      Top             =   6900
      Width           =   3345
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Modifique el texto libremente. Para grabar los cambios presione el boton cambiar"
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
      Height          =   705
      Index           =   8
      Left            =   8370
      TabIndex        =   17
      Top             =   4455
      Width           =   3480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto en configuración"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   7
      Left            =   8370
      TabIndex        =   16
      Top             =   4050
      Width           =   3510
   End
   Begin VB.Image TapaRank 
      Height          =   1245
      Left            =   7155
      Stretch         =   -1  'True
      Top             =   6660
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":00B6
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
      Height          =   1545
      Index           =   5
      Left            =   3600
      TabIndex        =   14
      Top             =   6975
      Width           =   3390
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Portada del ranking"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   6
      Left            =   3600
      TabIndex        =   13
      Top             =   6570
      Width           =   3405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos en la página principal"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   4
      Left            =   30
      TabIndex        =   10
      Top             =   3540
      Width           =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":017D
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
      Height          =   675
      Index           =   3
      Left            =   30
      TabIndex        =   9
      Top             =   3900
      Width           =   7215
   End
   Begin VB.Image imgIndexCHI 
      BorderStyle     =   1  'Fixed Single
      Height          =   1665
      Left            =   45
      Stretch         =   -1  'True
      Top             =   5190
      Width           =   1995
   End
   Begin VB.Image imgPRESp 
      BorderStyle     =   1  'Fixed Single
      Height          =   1890
      Left            =   8010
      Stretch         =   -1  'True
      Top             =   1650
      Width           =   2475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imagenes de presentacion del soft"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   2
      Left            =   7290
      TabIndex        =   7
      Top             =   60
      Width           =   4545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":0244
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
      Height          =   1125
      Index           =   1
      Left            =   7290
      TabIndex        =   6
      Top             =   420
      Width           =   4545
   End
   Begin VB.Image img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   2000
      Left            =   4890
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   2000
      Left            =   2460
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2000
      Left            =   60
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":031D
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
      Height          =   675
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Imagenes de inicio y cierre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Index           =   18
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7200
   End
End
Attribute VB_Name = "frmSUPERlic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CmdLg As New CommonDialog

Private Sub cmdImg_Click(Index As Integer)
    CmdLg.Filter = "Imagenes BMP (*.bmp)|*.BMP; *.sys"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    Select Case Index
        Case 1
            'imagen de inicio logo.sys
            FSO.CopyFile ArchSel, SYSfolder + "f5yaSL.nam", True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img1.Picture = LoadPicture(SYSfolder + "f5yaSL.nam")
        Case 2
            'imagen de cerrando logow.sys
            FSO.CopyFile ArchSel, SYSfolder + "f7yaSL.nam", True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img2.Picture = LoadPicture(SYSfolder + "f7yaSL.nam")
        Case 3
            'imagen de puede apagar logos.sys
            FSO.CopyFile ArchSel, SYSfolder + "f6yaSL.nam", True
            'se grtaba con otro nombre (igual pero con el SL)
            'luego al usarlo reviso, si existe el SL entonces lo uso con prioridad
            img3.Picture = LoadPicture(SYSfolder + "f6yaSL.nam")
    End Select
    'LISTO!!!
End Sub

Private Sub cmdImgPresP_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    If FSO.FolderExists(WINfolder + "SL") = False Then FSO.CreateFolder (WINfolder + "SL")
    If FSO.FileExists(WINfolder + "SL\imgBig.tbr") Then FSO.DeleteFile WINfolder + "SL\imgBig.tbr", True
    'grabar la imagen elegida
    FSO.CopyFile ArchSel, WINfolder + "SL\imgbig.tbr", True
    'mostrar que se cambio
    imgPRESp.Picture = LoadPicture(ArchSel)
    'LISTO!!!
End Sub

Private Sub cmdImgQ_Click(Index As Integer)
    Dim ArchSel As String
    Select Case Index
        Case 1
            'imagen de inicio logo.sys
            ArchSel = SYSfolder + "f5yaSL.nam"
            If FSO.FileExists(ArchSel) Then FSO.DeleteFile ArchSel, True
            'volver
            img1.Picture = LoadPicture(SYSfolder + "f56.dlw")
        Case 2
            'imagen de inicio logo.sys
            ArchSel = SYSfolder + "f7yaSL.nam"
            If FSO.FileExists(ArchSel) Then FSO.DeleteFile ArchSel, True
            'volver
            img2.Picture = LoadPicture(SYSfolder + "f58.dlw")
        Case 3
            'imagen de inicio logo.sys
            ArchSel = SYSfolder + "f6yaSL.nam"
            If FSO.FileExists(ArchSel) Then FSO.DeleteFile ArchSel, True
            'volver
            img3.Picture = LoadPicture(SYSfolder + "f57.dlw")
    End Select
    'LISTO!!!
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    If FSO.FolderExists(WINfolder + "SL") = False Then FSO.CreateFolder (WINfolder + "SL")
    If FSO.FileExists(WINfolder + "SL\indexCHI.tbr") Then FSO.DeleteFile WINfolder + "SL\indexCHI.tbr", True
    'grabar la imagen elegida
    FSO.CopyFile ArchSel, WINfolder + "SL\indexCHI.tbr", True
    'mostrar que se cambio
    imgIndexCHI.Picture = LoadPicture(ArchSel)
    'LISTO!!!
End Sub

Private Sub Command3_Click()
    If FSO.FolderExists(WINfolder + "SL") = False Then FSO.CreateFolder (WINfolder + "SL")
    If FSO.FileExists(WINfolder + "SL\txtIDX.tbr") Then FSO.DeleteFile WINfolder + "SL\txtIDX.tbr", True
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(WINfolder + "SL\txtIDX.tbr", True)
    'si deja en blanco jode!!!!!!
    If lblTBR = "" Then lblTBR = " "
    TE.Write lblTBR
    TE.Close
End Sub

Private Sub Command4_Click()
    'borrar eñl archivo!
    If FSO.FileExists(WINfolder + "SL\imgbig.tbr") Then _
        FSO.DeleteFile WINfolder + "SL\imgbig.tbr", True
    
    'mostrar el original
    imgPRESp.Picture = LoadPicture(SYSfolder + "f52.dlw")
End Sub

Private Sub Command5_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    'imagen de rank
    'XXXX
    'poner en otro archivo o ver porque sino se reemplazara
    'con el inicio que se redescomprime
    FSO.CopyFile ArchSel, SYSfolder + "f9yaSL.nam", True
    TapaRank.Picture = LoadPicture(SYSfolder + "f9yaSL.nam")
    'LISTO!!!
End Sub

Private Sub Command6_Click()
    If FSO.FolderExists(WINfolder + "SL") = False Then FSO.CreateFolder (WINfolder + "SL")
    If FSO.FileExists(WINfolder + "SL\txtCFG.tbr") Then FSO.DeleteFile WINfolder + "SL\txtCFG.tbr", True
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(WINfolder + "SL\txtCFG.tbr", True)
    If txtCFG = "" Then txtCFG = " "
    TE.Write txtCFG
    TE.Close
End Sub

Private Sub Command7_Click()
    Dim ArchSel As String
    ArchSel = WINfolder + "SL\indexCHI.tbr"
    If FSO.FileExists(ArchSel) Then FSO.DeleteFile ArchSel, True
    'mostrar original
    imgIndexCHI.Picture = LoadPicture(SYSfolder + "f61.dlw")
    'LISTO!!!
End Sub

Private Sub Command8_Click()
    Dim ArchSel As String
    ArchSel = SYSfolder + "f9yaSL.nam"
    If FSO.FileExists(ArchSel) Then FSO.DeleteFile ArchSel, True
    TapaRank.Picture = LoadPicture(SYSfolder + "f54.dlw")
    'LISTO!!!
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    'imágenes de inicio
    'ver si hay cargadas exclusivas
    If FSO.FileExists(SYSfolder + "f5yaSL.nam") Then
        img1.Picture = LoadPicture(SYSfolder + "f5yaSL.nam")
    Else
        img1.Picture = LoadPicture(SYSfolder + "f56.dlw")
    End If
    
    If FSO.FileExists(SYSfolder + "f7yaSL.nam") Then
        img2.Picture = LoadPicture(SYSfolder + "f7yaSL.nam")
    Else
        img2.Picture = LoadPicture(SYSfolder + "f58.dlw")
    End If
    
    If FSO.FileExists(SYSfolder + "f6yaSL.nam") Then
        img3.Picture = LoadPicture(SYSfolder + "f6yaSL.nam")
    Else
        img3.Picture = LoadPicture(SYSfolder + "f57.dlw")
    End If
    'la tapa de CD es la misma que la de rank que la del index que la del reg
    imgIndexCHI.Picture = LoadPicture(SYSfolder + "f61.dlw")
    TapaRank.Picture = LoadPicture(SYSfolder + "f54.dlw")
    'si hay Sl mostrar!
    If FSO.FileExists(WINfolder + "SL\indexCHI.tbr") Then
        imgIndexCHI.Picture = LoadPicture(WINfolder + "SL\indexCHI.tbr")
    End If
    If FSO.FileExists(SYSfolder + "f9yaSL.nam") Then
        TapaRank.Picture = LoadPicture(SYSfolder + "f9yaSL.nam")
    End If
    
    'cargar originales
    imgPRESp.Picture = LoadPicture(SYSfolder + "f52.dlw")
    'si existen reemplazan a las originales...
    If FSO.FileExists(WINfolder + "SL\imgbig.tbr") Then imgPRESp.Picture = LoadPicture(WINfolder + "SL\imgbig.tbr")
    If FSO.FileExists(WINfolder + "SL\imgtbr.tbr") Then imgIniTBR.Picture = LoadPicture(WINfolder + "SL\imgtbr.tbr")
    
    If FSO.FileExists(WINfolder + "SL\txtcfg.tbr") Then
        Set TE = FSO.OpenTextFile(WINfolder + "SL\txtcfg.tbr", ForReading, False)
        txtCFG = TE.ReadAll
        TE.Close
    End If
    If FSO.FileExists(WINfolder + "SL\txtIDX.tbr") Then
        Set TE = FSO.OpenTextFile(WINfolder + "SL\txtIDX.tbr", ForReading, False)
        lblTBR = TE.ReadAll
        TE.Close
    Else
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
    End If
    
End Sub
