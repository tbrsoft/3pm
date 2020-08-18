VERSION 5.00
Begin VB.Form frmINDEX 
   BackColor       =   &H00404040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Actualizacion de 3PM"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5130
   Icon            =   "frmINDEX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1155
      Left            =   90
      Picture         =   "frmINDEX.frx":0442
      ScaleHeight     =   1095
      ScaleWidth      =   3570
      TabIndex        =   7
      Top             =   3540
      Width           =   3630
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reubicar..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3780
      TabIndex        =   6
      Top             =   2040
      Width           =   1305
   End
   Begin VB.TextBox txtUbic3PM 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmINDEX.frx":1EB0
      Top             =   1290
      Width           =   5025
   End
   Begin VB.CommandButton cmdUP 
      Caption         =   "Actualizar!"
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
      Height          =   495
      Left            =   60
      TabIndex        =   3
      Top             =   2730
      Width           =   1425
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "ASEGURESE QUE 3PM NO SE ESTE EJECUTANDO EN ESTE MOMENTO!!!"
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
      Height          =   405
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicacion de la instalacion de 3PM"
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
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   1080
      Width           =   3525
   End
   Begin VB.Label lblBAR 
      BackColor       =   &H0000FFFF&
      Height          =   225
      Left            =   1560
      TabIndex        =   2
      Top             =   2880
      Width           =   285
   End
   Begin VB.Label lblTODOBar 
      BackColor       =   &H00000080&
      Height          =   225
      Left            =   1590
      TabIndex        =   1
      Top             =   2880
      Width           =   2115
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Actualizacion de 3PM a la versión Kabalin 6.5"
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
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   5025
   End
End
Attribute VB_Name = "frmINDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CmDlg As New CommonDialog
Dim FSO As New Scripting.FileSystemObject
Dim CarpInst As String
Dim Ap As String
Dim TotFiles As Integer

Private Sub cmdUP_Click()
    'ver si estn los archivos para actulizar
    'para 3.4.820
    Dim NoEsta As Integer
    NoEsta = 0
    If FSO.FileExists(Ap + "source\1.tbr") = False Then NoEsta = NoEsta + 1 '3pm.exe
    'If FSO.FileExists(Ap + "source\2.tbr") = False Then NoEsta = NoEsta + 1 'tapa.jpg
    'If FSO.FileExists(Ap + "source\3.tbr") = False Then NoEsta = NoEsta + 1 'top10.jpg
    'If FSO.FileExists(Ap + "source\4.tbr") = False Then NoEsta = NoEsta + 1 'logo.sys
    'If FSO.FileExists(Ap + "source\5.tbr") = False Then NoEsta = NoEsta + 1 'logos.sys
    'If FSO.FileExists(Ap + "source\6.tbr") = False Then NoEsta = NoEsta + 1 'logow.sys
    'If FSO.FileExists(Ap + "source\7.tbr") = False Then NoEsta = NoEsta + 1 'ini.exe
    'If FSO.FileExists(Ap + "source\8.tbr") = False Then NoEsta = NoEsta + 1 'manual.doc
    TotFiles = 1
    
    If NoEsta > 0 Then
        MsgBox "Faltan algunos archivos para actualzar! No se puede realizar la tarea"
        Exit Sub
    End If
    'OK todo bien
    On Error GoTo NoPuede
    UpdateFile CarpInst + "3pm.exe", 1
    'UpdateFile CarpInst + "tapa.jpg", 2
    'UpdateFile CarpInst + "top10.jpg", 3
    'UpdateFile CarpInst + "logo.sys", 4
    'UpdateFile CarpInst + "logos.sys", 5
    'UpdateFile CarpInst + "logow.sys", 6
    'UpdateFile CarpInst + "ini.exe", 7
    'UpdateFile CarpInst + "manual.doc", 8
    MsgBox "La actualizacion se ha realizado correctamente"
    End
NoPuede:
    MsgBox "Error al actualizar 3pm. N° " + CStr(Err.Number) + ". Descripcion: " + Err.Description
End Sub

Public Sub UpdateFile(Arch As String, nSource As Integer)
    If FSO.FileExists(Arch) Then FSO.DeleteFile Arch, True
    FSO.CopyFile Ap + "source\" + CStr(nSource) + ".tbr", Arch, True
    lblBAR.Width = nSource / TotFiles * lblTODOBar.Width
    lblBAR.Refresh
End Sub

Private Sub Command2_Click()
    CmDlg.Filter = "Ejecutable 3PM|3pm.exe"
    CmDlg.ShowOpen
    If CmDlg.FileName = "" Then Exit Sub
    If UCase(CmDlg.FileTitle) <> "3PM.EXE" Then
        MsgBox "No es válido el archivo elegido!. Pruebe de nuevo"
        Exit Sub
    End If
    'si llego aca esta todo OK
    CarpInst = Left(CmDlg.FileName, Len(CmDlg.FileName) - Len(CmDlg.FileTitle))
    'si o si tiene la "\"
    txtUbic3PM = CarpInst + vbCrLf + _
        "Se ha encontrado OK!!"
    cmdUP.Enabled = True
End Sub

Private Sub Form_Load()
    Ap = App.path
    If Right(Ap, 1) <> "\" Then Ap = Ap + "\"
    Me.Caption = "Actualizacion de 3PM tbrSoft version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
    'ver si esta en la carpeta por defecto
    If FSO.FileExists("c:\archivos de programa\3pm\3pm.exe") Then
        'todo OK
        cmdUP.Enabled = True
        CarpInst = "c:\archivos de programa\3pm\" 'si o si con la barra
        txtUbic3PM = "c:\archivos de programa\3pm\3pm.exe" + vbCrLf + _
        "Se ha encontrado OK!!"
    Else
        txtUbic3PM = "No se ha encontrado 3PM en su ubicacion por defecto " + _
            "utilize el boton REUBICAR para indicar la ubicacion de 3PM"
        cmdUP.Enabled = False
    End If
End Sub

