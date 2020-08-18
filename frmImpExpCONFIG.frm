VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmImpExpCONFIG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
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
   Begin tbrFaroButton.fBoton command26 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   900
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Importar Configuración"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1290
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Exportar Configuración"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1680
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Puede utilizar EXPORTAR para realizar copias de seguridad de su archivo de configuración."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   615
      Index           =   18
      Left            =   180
      TabIndex        =   0
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
    CmdLg.DialogTitle = TR.Trad("Importar Archivo de configuración de 3PM%99%")
    CmdLg.ShowOpen
    Dim F As String
    F = CmdLg.FileName
    If F = "" Then Exit Sub
    If fso.FileExists(F) Then
        If MsgBox(TR.Trad("¿Esta seguro que desea reemplazar el archivo de " + _
            "configuracion actual?%99%"), vbQuestion + vbYesNo) = _
            vbNo Then Exit Sub
    End If
    fso.CopyFile F, GPF("config"), True
    MsgBox TR.Trad("El archivo se importo correctamente. 3PM se cerrará%99%")
    Unload Me
    frmConfig.SendW
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command2.Caption = TR.Trad("SALIR%99%")
    Command1.Caption = TR.Trad("Exportar Configuracion%99%")
    command26.Caption = TR.Trad("Importar Configuracion%99%")
    Label1(18).Caption = TR.Trad("Puede utilizar EXPORTAR para realizar copias de " + _
        "seguridad de su archivo de configuración.%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

