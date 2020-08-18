VERSION 5.00
Begin VB.Form frmSUPERlic 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Left            =   1320
      MultiLine       =   -1  'True
      TabIndex        =   22
      Text            =   "frmSUPERlic.frx":0000
      Top             =   5580
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
      Left            =   7560
      MultiLine       =   -1  'True
      TabIndex        =   21
      Text            =   "frmSUPERlic.frx":000A
      Top             =   5640
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
      Left            =   10590
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5550
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
      Left            =   5100
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8130
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
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
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8130
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
      Left            =   5970
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6180
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
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6390
      Width           =   1215
   End
   Begin VB.CommandButton cmdImgPresTbr 
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
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3510
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
      Left            =   9570
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1920
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
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4140
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
      Index           =   2
      Left            =   3030
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4140
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
      Index           =   1
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4140
      Width           =   1215
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
      Height          =   615
      Index           =   8
      Left            =   7260
      TabIndex        =   19
      Top             =   4920
      Width           =   4605
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
      Left            =   7260
      TabIndex        =   18
      Top             =   4530
      Width           =   4590
   End
   Begin VB.Image TapaRank 
      Height          =   1245
      Left            =   5040
      Stretch         =   -1  'True
      Top             =   6870
      Width           =   1320
   End
   Begin VB.Image TapaCD 
      Height          =   1245
      Left            =   3660
      Stretch         =   -1  'True
      Top             =   6870
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":00A5
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
      Height          =   1275
      Index           =   5
      Left            =   30
      TabIndex        =   15
      Top             =   7290
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tapas por defecto de los discos"
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
      Left            =   30
      TabIndex        =   14
      Top             =   6900
      Width           =   3630
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
      Left            =   0
      TabIndex        =   11
      Top             =   4530
      Width           =   7200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":016C
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
      Left            =   0
      TabIndex        =   10
      Top             =   4890
      Width           =   7215
   End
   Begin VB.Image imgIndexCHI 
      BorderStyle     =   1  'Fixed Single
      Height          =   1260
      Left            =   30
      Picture         =   "frmSUPERlic.frx":0233
      Stretch         =   -1  'True
      Top             =   5610
      Width           =   1230
   End
   Begin VB.Image imgIniTBR 
      BorderStyle     =   1  'Fixed Single
      Height          =   690
      Left            =   9570
      Picture         =   "frmSUPERlic.frx":153B
      Stretch         =   -1  'True
      Top             =   2760
      Width           =   2280
   End
   Begin VB.Image imgPRESp 
      BorderStyle     =   1  'Fixed Single
      Height          =   1890
      Left            =   7320
      Picture         =   "frmSUPERlic.frx":2FA9
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   2250
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
      Caption         =   $"frmSUPERlic.frx":4A73
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
      Height          =   1455
      Index           =   1
      Left            =   7290
      TabIndex        =   6
      Top             =   420
      Width           =   4545
   End
   Begin VB.Image img3 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   4890
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Image img2 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   2460
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Image img1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3000
      Left            =   60
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   2400
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmSUPERlic.frx":4B76
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
            FSO.CopyFile ArchSel, AP + "logo.sys", True
            img1.Picture = LoadPicture(AP + "logo.sys")
        Case 2
            'imagen de cerrando logow.sys
            FSO.CopyFile ArchSel, AP + "logow.sys", True
            img1.Picture = LoadPicture(AP + "logow.sys")
        Case 3
            'imagen de puede apagar logos.sys
            FSO.CopyFile ArchSel, AP + "logos.sys", True
            img1.Picture = LoadPicture(AP + "logos.sys")
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
    If FSO.FolderExists(WINfolder + "\SL") = False Then FSO.CreateFolder (WINfolder + "\SL")
    If FSO.FileExists(WINfolder + "\SL\imgBig.tbr") Then FSO.DeleteFile WINfolder + "\SL\imgBig.tbr", True
    'grabar la imagen elegida
    FSO.CopyFile ArchSel, WINfolder + "\SL\imgbig.tbr", True
    'mostrar que se cambio
    imgPRESp.Picture = LoadPicture(ArchSel)
    'LISTO!!!
End Sub

Private Sub cmdImgPresTbr_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    If FSO.FolderExists(WINfolder + "\SL") = False Then FSO.CreateFolder (WINfolder + "\SL")
    If FSO.FileExists(WINfolder + "\SL\imgTBR.tbr") Then FSO.DeleteFile WINfolder + "\SL\imgTBR.tbr", True
    'grabar la imagen elegida
    FSO.CopyFile ArchSel, WINfolder + "\SL\imgTBR.tbr", True
    'mostrar que se cambio
    imgIniTBR.Picture = LoadPicture(ArchSel)
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
    If FSO.FolderExists(WINfolder + "\SL") = False Then FSO.CreateFolder (WINfolder + "\SL")
    If FSO.FileExists(WINfolder + "\SL\indexCHI.tbr") Then FSO.DeleteFile WINfolder + "\SL\indexCHI.tbr", True
    'grabar la imagen elegida
    FSO.CopyFile ArchSel, WINfolder + "\SL\indexCHI.tbr", True
    'mostrar que se cambio
    imgIndexCHI.Picture = LoadPicture(ArchSel)
    'LISTO!!!
End Sub

Private Sub Command3_Click()
    If FSO.FolderExists(WINfolder + "\SL") = False Then FSO.CreateFolder (WINfolder + "\SL")
    If FSO.FileExists(WINfolder + "\SL\txtIDX.tbr") Then FSO.DeleteFile WINfolder + "\SL\txtIDX.tbr", True
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(WINfolder + "\SL\txtIDX.tbr", True)
    TE.Write lblTBR
    TE.Close
End Sub

Private Sub Command4_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    'imagen de inicio logo.sys
    FSO.CopyFile ArchSel, AP + "tapa.jpg", True
    TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
    'LISTO!!!
End Sub

Private Sub Command5_Click()
    CmdLg.Filter = "Imagenes JPG GIF |*.jpg; *.gif; *.jpeg"
    CmdLg.DialogTitle = "Eliga nueva imagen de SUPERLICENCIA"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    Dim ArchSel As String
    ArchSel = CmdLg.FileName
    'imagen de inicio logo.sys
    FSO.CopyFile ArchSel, AP + "top10.jpg", True
    TapaCD.Picture = LoadPicture(AP + "top10.jpg")
    'LISTO!!!
End Sub

Private Sub Command6_Click()
    If FSO.FolderExists(WINfolder + "\SL") = False Then FSO.CreateFolder (WINfolder + "\SL")
    If FSO.FileExists(WINfolder + "\SL\txtCFG.tbr") Then FSO.DeleteFile WINfolder + "\SL\txtCFG.tbr", True
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(WINfolder + "\SL\txtCFG.tbr", True)
    TE.Write txtCFG
    TE.Close
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    'imágenes de inicio
    img1.Picture = LoadPicture(AP + "logo.sys")
    img2.Picture = LoadPicture(AP + "logow.sys")
    img3.Picture = LoadPicture(AP + "logos.sys")
    TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
    TapaRank.Picture = LoadPicture(AP + "top10.jpg")
    If FSO.FileExists(WINfolder + "\SL\txtcfg.tbr") Then
        Set TE = FSO.OpenTextFile(WINfolder + "\SL\txtcfg.tbr", ForReading, False)
        txtCFG = TE.ReadAll
        TE.Close
    End If
    If FSO.FileExists(WINfolder + "\SL\txtIDX.tbr") Then
        Set TE = FSO.OpenTextFile(WINfolder + "\SL\txtIDX.tbr", ForReading, False)
        lblTBR = TE.ReadAll
        TE.Close
    Else
        lblTBR = "Software desarrollado por tbrSoft www.tbrsoft.com - info@tbrsoft.com - tbrsoft@cpcipc.org."
    End If
    
End Sub

