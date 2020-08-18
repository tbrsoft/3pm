VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmSUPERlic 
   AutoRedraw      =   -1  'True
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
   Begin tbrFaroButton.fBoton Command2 
      Height          =   585
      Left            =   300
      TabIndex        =   6
      Top             =   6150
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Definir SKIN. Cambio gráfico completo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command3 
      Height          =   645
      Left            =   2340
      TabIndex        =   4
      Top             =   1980
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "cambiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.TextBox lblTBR 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1245
      Left            =   3030
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2235
      Left            =   3060
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmSUPERlic.frx":000A
      Top             =   2700
      Width           =   2865
   End
   Begin tbrFaroButton.fBoton command6 
      Height          =   645
      Left            =   2370
      TabIndex        =   5
      Top             =   5040
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "cambiar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   585
      Left            =   3270
      TabIndex        =   7
      Top             =   6150
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto en pantalla principal: Modifique el texto libremente. Para grabar los cambios presione el botón cambiar."
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
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   3
      Top             =   570
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Texto en la configuración: Modifique el texto libremente. Para grabar los cambios presione el botón cambiar."
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
      Height          =   2235
      Index           =   8
      Left            =   240
      TabIndex        =   0
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
    
    If fso.FileExists(GPF("tslpri112")) Then fso.DeleteFile GPF("tslpri112"), True
    
    'si deja en blanco jode!!!!!!
    If lblTBR = "" Then lblTBR = " "
    
    'grabar el texto como un nuevo archivo
    Set TE = fso.CreateTextFile(GPF("tslpri112"), True)
        TE.Write lblTBR
    TE.Close
End Sub

Private Sub Command6_Click()
    'texto en config No Tbr Wf + "SL\txtCFG.tbr"
    If fso.FileExists(GPF("telcnot")) Then fso.DeleteFile GPF("telcnot"), True
    'grabar el texto como un nuevo archivo
    Set TE = fso.CreateTextFile(GPF("telcnot"), True)
    If txtCFG = "" Then txtCFG = " "
    TE.Write txtCFG.Text
    TE.Close
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    'texto en config No Tbr Wf + "SL\txtCFG.tbr"
    If fso.FileExists(GPF("telcnot")) Then
        Set TE = fso.OpenTextFile(GPF("telcnot"), ForReading, False)
        txtCFG.Text = TE.ReadAll
        TE.Close
    End If
    'texto de SL
    If fso.FileExists(GPF("tslpri112")) Then
        Set TE = fso.OpenTextFile(GPF("tslpri112"), ForReading, False)
        lblTBR = TE.ReadAll
        TE.Close
    Else
        TR.SetVars "tbrSoft", "info@tbrsoft.com", "tbrsoft@cpcipc.org"
        lblTBR = TR.Trad("Software desarrollado" + vbCrLf + _
                "por %01% " + vbCrLf + _
                "www.%01%.com" + vbCrLf + _
                "%02%" + vbCrLf + _
                "%03%%98% la variable 1 dice tbrsoft y 2 y 3 son emails%99%")
    End If
    
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command2.Caption = TR.Trad("Definir SKIN. Cambio grafico completo%99%")
    
    TR.SetVars "tbrSoft", _
               "info@tbrsoft.com", _
               "tbrsoft@cpcipc.org", _
               "Argentina"
               
    txtCFG.Text = TR.Trad("Desarrollado por %01%" + vbCrLf + _
        "www.%01%.com" + vbCrLf + _
        "----------------" + vbCrLf + _
        "Contáctenos a %02%" + vbCrLf + _
        "%03%" + vbCrLf + _
        "----------------" + vbCrLf + _
        "Hecho en %04%%99%")
        
    Command6.Caption = TR.Trad("Cambiar%99%")
    Command3.Caption = TR.Trad("Cambiar%99%")
    Command1.Caption = TR.Trad("SALIR%99%")
    Label1(1).Caption = TR.Trad("Texto en pantalla principal: " + _
        "Modifique el texto libremente. Para grabar los cambios " + _
        "presione el boton Cambiar%99%")
    Label1(8).Caption = TR.Trad("Texto en la configuracion: Modifique " + _
        "el texto libremente. Para grabar los cambios presione el boton cambiar%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
