VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmClaves 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Claves personales de 3PM"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6960
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command8 
      Height          =   375
      Left            =   4110
      TabIndex        =   12
      Top             =   4110
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command6 
      Height          =   375
      Left            =   4110
      TabIndex        =   11
      Top             =   3690
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar claves"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.TextBox txtLenCredit 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3060
      TabIndex        =   10
      Text            =   "19"
      Top             =   4230
      Width           =   495
   End
   Begin VB.TextBox txtLenCLose 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3150
      TabIndex        =   9
      Text            =   "20"
      Top             =   2490
      Width           =   495
   End
   Begin VB.TextBox txtLenConfig 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3120
      TabIndex        =   8
      Text            =   "20"
      Top             =   1830
      Width           =   495
   End
   Begin VB.TextBox txtClaveCredit 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Text            =   "4444444444444444444"
      Top             =   4230
      Width           =   2985
   End
   Begin VB.TextBox txtClaveCLose 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "12345612345612345612"
      Top             =   2490
      Width           =   3075
   End
   Begin VB.TextBox txtClaveConfig 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   30
      TabIndex        =   1
      Text            =   "88888888888888888888"
      Top             =   1830
      Width           =   3075
   End
   Begin VB.Label lblIDteclas 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1995
      Left            =   3930
      TabIndex        =   7
      Top             =   1380
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClaves.frx":0000
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1275
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   3030
      Width           =   3525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de ingreso a la configuración."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   2280
      Width           =   3795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave de ingreso a la configuración."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1500
      Width           =   6465
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClaves.frx":00A5
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1035
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   6495
   End
End
Attribute VB_Name = "frmClaves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command6_Click()
    'ver que tengan el largo correspondiente
    Dim LenClave As Integer
    LenClave = Len(txtClaveConfig)
    TR.SetVars LenClave
    msg1 = TR.Trad("La clave de configuración tiene %01% caracteres. Debe tener 20 para poder grabar%99%")
    TR.SetVars LenClave
    msg2 = TR.Trad("La clave de cerrado tiene %01% caracteres. Debe tener 20 para poder grabar%99%")
    TR.SetVars LenClave
    msg3 = TR.Trad("La clave de creditos tiene %01% caracteres. Debe tener 19 para poder grabar%99%")
    
    If LenClave <> 20 Then MsgBox msg1: Exit Sub
    LenClave = Len(txtClaveCLose)
    If LenClave <> 20 Then MsgBox msg2: Exit Sub
    LenClave = Len(txtClaveCredit)
    If LenClave <> 19 Then MsgBox msg3: Exit Sub
    
    'ok todas las claves estan bien
    Set TE = fso.CreateTextFile(GPF("sequeda32"), True)
        TE.WriteLine "Config:" + txtClaveConfig
        TE.WriteLine "Close:" + txtClaveCLose
        TE.WriteLine "Credit:" + txtClaveCredit
    TE.Close
    Unload Me
End Sub

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub fBoton1_Click()
    
End Sub

Private Sub Form_Activate()
    Label1(0) = TR.Trad("Modifique sus claves para obtener mayor seguridad. " + _
        "Utilice solo las teclas que usted expone al público para no " + _
        "perder funcionalidad. Si no desea habilitar esta claves podrá " + _
        "cargar algún caracter no válido para que estas claves no puedan " + _
        "ser usadas (ej: el caracter '7')%99%")
        
    Label1(1) = TR.Trad("Clave para ingreso a la configuración%99%")
    Label1(2) = TR.Trad("Clave para cerrar sistema%99%")
    Label1(3) = TR.Trad("Clave para cargar créditos. Son 19 dígitos, " + _
        "el ultimo dependerá de la cantidad de créditos que desee cargar" + _
        ". Recuerde que estos créditos no se suman al contador%99%")
    Command6.Caption = TR.Trad("Grabar claves%99%")
    Command8.Caption = TR.Trad("Salir sin Grabar%99%")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            Unload Me
            YaCerrar3PM
    End Select
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    If fso.FileExists(GPF("sequeda32")) = False Then
        MsgBox TR.Trad("No esta presente el archivo de claves. Reinicie 3PM%99%")
        Exit Sub
    End If
    Set TE = fso.OpenTextFile(GPF("sequeda32"), ForReading, False)
    'config/close/credit es el orden del archivo
    txtClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    txtClaveCLose = txtInLista(TE.ReadLine, 1, ":")
    txtClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    lblIDteclas = TR.Trad("Identificación de teclas%98%Asignacion " + _
        "de valores a diferentes teclas%99%") + vbCrLf + vbCrLf + _
        TR.Trad("1- Izquierda%99%") + vbCrLf + _
        TR.Trad("2- Derecha%99%") + vbCrLf + _
        TR.Trad("3- OK%99%") + vbCrLf + _
        TR.Trad("4- Escape%99%") + vbCrLf + _
        TR.Trad("5- Página adelante%99%") + vbCrLf + _
        TR.Trad("6- Página atras%99%")
    
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub txtClaveCLose_Change()
    txtLenCLose = CStr(Len(txtClaveCLose))
    Command6.Enabled = LargoClavesOK
End Sub

Private Sub txtClaveConfig_Change()
    txtLenConfig = CStr(Len(txtClaveConfig))
    Command6.Enabled = LargoClavesOK
End Sub

Private Sub txtClaveCredit_Change()
    txtLenCredit = CStr(Len(txtClaveCredit))
    Command6.Enabled = LargoClavesOK
End Sub

Public Function LargoClavesOK() As Boolean
    If txtLenConfig = "20" Then
        If txtLenCLose = "20" Then
            If txtLenCredit = "19" Then
                LargoClavesOK = True
                Exit Function
            End If
        End If
    End If
    LargoClavesOK = False
End Function
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Label1(0) = TR.Trad("Modifique sus claves para obtener mayor seguridad. " + _
        "Utilice solo las teclas que usted expone al público para no " + _
        "perder funcionalidad. Si no desea habilitar esta clave podrá " + _
        "cargar algún caracter no válido para que estas claves no puedan " + _
        "ser usadas (Ej.: el caracter '7').%99%")
        
    Label1(1) = TR.Trad("Clave de ingreso a la configuración.%99%")
    Label1(2) = TR.Trad("Clave de ingreso a la configuración.%99%")
    '
    Label1(3) = TR.Trad("Clave para cargar créditos. Son 19 dígitos, " + _
        "el último dependerá de la cantidad de créditos que desee cargar" + _
        ". Recuerde que estos créditos no se suman al contador.%99%")
    Command6.Caption = TR.Trad("Grabar claves%99%")
    Command8.Caption = TR.Trad("Salir sin Grabar%99%")
End Sub
