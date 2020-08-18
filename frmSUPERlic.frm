VERSION 5.00
Begin VB.Form frmSUPERlic 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SuperLicencia"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808000&
      Caption         =   "Definir SKIN. Cambio grafico completo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   330
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5910
      Width           =   2715
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
      Height          =   1245
      Left            =   3030
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmSUPERlic.frx":0000
      Top             =   540
      Width           =   2895
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
      Left            =   3060
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmSUPERlic.frx":000A
      Top             =   2700
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
      Height          =   735
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4980
      Width           =   1095
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
      Height          =   705
      Left            =   2550
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1005
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
      Height          =   525
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6180
      Width           =   2625
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto en pantalla principal: Modifique el texto libremente. Para grabar los cambios presione el boton cambiar"
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
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   570
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto en la configuracion: Modifique el texto libremente. Para grabar los cambios presione el boton cambiar"
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
      Height          =   2235
      Index           =   8
      Left            =   240
      TabIndex        =   2
      Top             =   2700
      Width           =   2850
   End
End
Attribute VB_Name = "frmSUPERlic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CmdLg As New CommonDialog

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    frmConfigVIS.Show 1
End Sub

Private Sub Command3_Click()
    
    If FSO.FileExists(GPF("tslpri112")) Then FSO.DeleteFile GPF("tslpri112"), True
    
    'si deja en blanco jode!!!!!!
    If lblTBR = "" Then lblTBR = " "
    
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(GPF("tslpri112"), True)
        TE.Write lblTBR
    TE.Close
End Sub

Private Sub Command6_Click()
    'texto en config No Tbr Wf + "SL\txtCFG.tbr"
    If FSO.FileExists(GPF("telcnot")) Then FSO.DeleteFile GPF("telcnot"), True
    'grabar el texto como un nuevo archivo
    Set TE = FSO.CreateTextFile(GPF("telcnot"), True)
    If txtCFG = "" Then txtCFG = " "
    TE.Write txtCFG
    TE.Close
End Sub

Private Sub Form_Load()
    'texto en config No Tbr Wf + "SL\txtCFG.tbr"
    If FSO.FileExists(GPF("telcnot")) Then
        Set TE = FSO.OpenTextFile(GPF("telcnot"), ForReading, False)
        txtCFG.Text = TE.ReadAll
        TE.Close
    End If
    'texto de SL
    If FSO.FileExists(GPF("tslpri112")) Then
        Set TE = FSO.OpenTextFile(GPF("tslpri112"), ForReading, False)
        lblTBR = TE.ReadAll
        TE.Close
    Else
        lblTBR = "Software desarrollado" + vbCrLf + _
                "por tbrSoft " + vbCrLf + _
                "www.tbrsoft.com" + vbCrLf + _
                "info@tbrsoft.com" + vbCrLf + _
                "tbrsoft@cpcipc.org."
    End If
    
End Sub
