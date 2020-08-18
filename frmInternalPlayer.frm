VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmInternalPlayer 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalles internos del reproductor"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   375
      Left            =   4950
      TabIndex        =   2
      Top             =   3270
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "SALIR"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.CheckBox chkValidarDriverVideo 
      BackColor       =   &H00000000&
      Caption         =   "Forzar los valores recomendados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1140
      TabIndex        =   1
      Top             =   3300
      Width           =   3015
   End
   Begin VB.Label lInfo1 
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxx xxxxxxxxxxxx xx xx xxx xxx xxxxx xxxxx xxxx xxxxx xxxxxxxxxxxx xx xx xxx xxx xxxxx xxxxx xxxx"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5655
   End
End
Attribute VB_Name = "frmInternalPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fBoton1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    chkValidarDriverVideo.Value = CLng(LeerConfig("ValidarDriverVideo", "1"))
    
    Dim TMP As String
    
    TMP = "Detalle de drivers internos de reproducción de video"
    TMP = TMP + vbCrLf
    TMP = TMP + vbCrLf
    
    'frmIndex.MP3.GetDefaultDevice("MPEGVideo") DEBE SER "mciqtz.drv"
    TMP = TMP + vbCrLf + "Driver para MPEG: " + frmIndex.MP3.GetDefaultDevice("MPEGVideo")
    TMP = TMP + vbCrLf + "tbrSoft recomienda: " + "mciqtz.drv"
    
    TMP = TMP + vbCrLf
    
    'frmIndex.MP3.GetDefaultDevice("avivideo") DEBE SER  "mciavi.drv" Then
    TMP = TMP + vbCrLf + "Driver para AVI: " + frmIndex.MP3.GetDefaultDevice("avivideo")
    TMP = TMP + vbCrLf + "tbrSoft recomienda: " + "mciavi.drv"
    
    TMP = TMP + vbCrLf
    TMP = TMP + vbCrLf + "Puede forzar a usar los valores recomendados por tbrSoft"
    
    lInfo1.Caption = TMP
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    ChangeConfig "ValidarDriverVideo", CStr(Abs(chkValidarDriverVideo.Value))
End Sub
