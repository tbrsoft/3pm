VERSION 5.00
Begin VB.Form frmImpExpCONFIG 
   BackColor       =   &H00000040&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Importar / Exportar Configuración"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   3345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1740
      Width           =   2650
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Exportar Configuracion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   2650
   End
   Begin VB.CommandButton Command26 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Importar Configuracion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   380
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   900
      Width           =   2650
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puede usar EXPORTAR para hacer copias de seguridad de su archivo de configuración."
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
      Height          =   615
      Index           =   18
      Left            =   180
      TabIndex        =   3
      Top             =   150
      Width           =   3015
   End
End
Attribute VB_Name = "frmImpExpCONFIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click() 'EPORTAR
    ExportarCFG
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command26_Click() 'IMPORTAR
    Dim CmdLg As New CommonDialog
    CmdLg.DialogTitle = "Importar Archivo de configuración de 3PM"
    CmdLg.ShowOpen
    Dim F As String
    F = CmdLg.FileName
    If F = "" Then Exit Sub
    If FSO.FileExists(F) Then
        If MsgBox("¿Esta seguro que desea reemplazar el archivo de " + _
            "configuracion actual? ", vbQuestion + vbYesNo) = _
            vbNo Then Exit Sub
    End If
    FSO.CopyFile F, SYSfolder + "3pmcfg.tbr", True
    MsgBox "El archivo se importo correctamente. 3PM se cerrará"
    Unload Me
    frmConfig.SendW
End Sub
