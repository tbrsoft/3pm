VERSION 5.00
Begin VB.Form frmClaves 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Claves personales de 3PM"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   12
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
      TabIndex        =   11
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
      TabIndex        =   10
      Text            =   "20"
      Top             =   1830
      Width           =   495
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grabar claves"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   2350
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Salir sin Grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4260
      Width           =   2350
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
      BackColor       =   &H00400000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1995
      Left            =   3690
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
      ForeColor       =   &H00400000&
      Height          =   1275
      Index           =   3
      Left            =   60
      TabIndex        =   6
      Top             =   2940
      Width           =   2985
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave para ingreso a la configuracion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   2250
      Width           =   3045
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave para ingreso a la configuracion"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Index           =   1
      Left            =   30
      TabIndex        =   4
      Top             =   1410
      Width           =   3285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClaves.frx":00A4
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
    
    Select Case IDIOMA
        Case "Español"
            msg1 = "La clave de configuración tiene " + _
                CStr(LenClave) + " caracteres. Debe tener 20 para poder grabar"
            msg2 = "La clave de cerrado tiene " + _
                CStr(LenClave) + " caracteres. Debe tener 20 para poder grabar"
            msg3 = "La clave de creditos tiene " + _
                CStr(LenClave) + " caracteres. Debe tener 19 para poder grabar"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    
    If LenClave <> 20 Then MsgBox msg1: Exit Sub
    LenClave = Len(txtClaveCLose)
    If LenClave <> 20 Then MsgBox msg2: Exit Sub
    LenClave = Len(txtClaveCredit)
    If LenClave <> 19 Then MsgBox msg3: Exit Sub
    
    'ok todas las claves estan bien
    Set TE = FSO.CreateTextFile(WINfolder + "/sevalc.dll", True)
    TE.WriteLine "Config:" + txtClaveConfig
    TE.WriteLine "Close:" + txtClaveCLose
    TE.WriteLine "Credit:" + txtClaveCredit
    TE.Close
    Unload Me
End Sub

Private Sub Command8_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Select Case IDIOMA
        Case "Español"
            Label1(0) = "Modifique sus claves para obtener mayor seguridad. Utilize solo las teclas que usted expone al público para no perder funcionalidad. Si no desea habilitar esta claves podrá cargar algún caracter no válido para que estas claves no puedan ser usadas (ej: el caracter '7')"
            Label1(1) = "Clave para ingreso a la configuracion"
            Label1(2) = "Clave para ingreso a la configuracion"
            Label1(3) = "Clave para cargar créditos. Son 19 dígitos, el ultimo dependerá de la cantidad de créditos que desee cargar. Recuerde que estos créditos no se suman al contador"
            Command6.Caption = "Grabar claves"
            Command8.Caption = "Salir sin Grabar"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            End
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
        'si ya hay 9 cargados se traga las fichas
        If CREDITOS <= MaximoFichas Then
            OnOffCAPS vbKeyScrollLock, True
            CREDITOS = CREDITOS + TemasPorCredito
            SumarContadorCreditos TemasPorCredito
            'grabar cant de creditos
            EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
            If CREDITOS >= 10 Then
                Select Case IDIOMA
                    Case "Español"
                        frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                    Case "English"
                    Case "Francois"
                    Case "Italiano"
                End Select
            Else
                Select Case IDIOMA
                    Case "Español"
                        frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                    Case "English"
                    Case "Francois"
                    Case "Italiano"
                End Select
            End If
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
            
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    If FSO.FileExists(WINfolder + "sevalc.dll") = False Then
        MsgBox "No esta presente el archivo de claves. Reinicie 3PM"
        Exit Sub
    End If
    Set TE = FSO.OpenTextFile(WINfolder + "sevalc.dll", ForReading, False)
    'config/close/credit es el orden del archivo
    txtClaveConfig = txtInLista(TE.ReadLine, 1, ":")
    txtClaveCLose = txtInLista(TE.ReadLine, 1, ":")
    txtClaveCredit = txtInLista(TE.ReadLine, 1, ":")
    TE.Close
    
    Select Case IDIOMA
        Case "Español"
            lblIDteclas = "Identificacion de teclas" + vbCrLf + vbCrLf + _
                "1- Izquierda" + vbCrLf + _
                "2- Derecha" + vbCrLf + _
                "3- OK" + vbCrLf + _
                "4- Escape" + vbCrLf + _
                "5- Página adelante" + vbCrLf + _
                "6- Página atras"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    
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
