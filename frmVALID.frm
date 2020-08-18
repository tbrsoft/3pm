VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmVALID 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Validacion de uso del propietario"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   8640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRegistroDiario 
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
      Height          =   1455
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Text            =   "frmVALID.frx":0000
      Top             =   3630
      Width           =   3885
   End
   Begin VB.TextBox txtEstadoValidacion 
      Alignment       =   2  'Center
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
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2730
      Width           =   3885
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1710
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   900
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox Text2 
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
      Height          =   2205
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   7
      Top             =   480
      Width           =   3885
   End
   Begin VB.CheckBox chkValid 
      BackColor       =   &H00000000&
      Caption         =   "Bloquear el equipo según usos."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4620
      TabIndex        =   0
      Top             =   540
      Width           =   3735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4755
      Left            =   4050
      TabIndex        =   1
      Top             =   360
      Width           =   4455
      Begin VB.TextBox tCONT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   270
         TabIndex        =   15
         Text            =   "0"
         Top             =   2820
         Width           =   4005
      End
      Begin VB.TextBox tPC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtPreAviso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1410
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0"
         Top             =   1500
         Width           =   1395
      End
      Begin VB.TextBox txtUSOS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "0"
         Top             =   840
         Width           =   1395
      End
      Begin tbrFaroButton.fBoton XxBoton1 
         Height          =   405
         Left            =   810
         TabIndex        =   6
         Top             =   3360
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Generar un archivo de claves"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton4 
         Height          =   405
         Left            =   810
         TabIndex        =   12
         Top             =   4200
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "GRABAR CAMBIOS"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin tbrFaroButton.fBoton XxBoton2 
         Height          =   405
         Left            =   810
         TabIndex        =   17
         Top             =   3780
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         fFColor         =   16777215
         fBColor         =   14737632
         fCapt           =   "Reiniciar conteo"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   5452834
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Valor Actual del contador histórico"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   540
         TabIndex        =   16
         Top             =   2580
         Width           =   3195
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "texto para recordar equipo"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   510
         TabIndex        =   9
         Top             =   1890
         Width           =   3195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Créditos de preaviso"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1230
         TabIndex        =   5
         Top             =   1260
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "cantidad de créditos a los que se bloqueará"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   510
         TabIndex        =   3
         Top             =   570
         Width           =   3405
      End
   End
   Begin tbrFaroButton.fBoton XxBoton3 
      Height          =   645
      Left            =   2580
      TabIndex        =   11
      Top             =   5220
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
End
Attribute VB_Name = "frmVALID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkValid_Click()
    txtUSOS.Enabled = CBool(chkValid.Value)
    txtPreAviso.Enabled = txtUSOS.Enabled
    tPC.Enabled = txtUSOS.Enabled
    XxBoton1.Enabled = txtUSOS.Enabled
    'XxBoton4.Enabled = txtUSOS.Enabled
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    'validar cada X Creditos
    VALIDAR = LeerConfig("Validar", "0")
    ValidarCada = LeerConfig("ValidarCada", "3000")
    AvisarAntes = LeerConfig("AvisarAntes", "500")
    
    If VALIDAR Then
        chkValid.Value = 1
        TR.SetVars CreditosValidar, ValidarCada, ValidarCada - CreditosValidar, CodigoParaClaveActual
        txtEstadoValidacion = TR.Trad("Estado de validación: " + vbCrLf + _
            "Creditos Usados: %01% de %02%" + vbCrLf + _
            "Quedan: %03%" + vbCrLf + _
            "Codigo Actual: %04% %98%La validacion es una cantidad de canciones que" + _
            " se pueden escuchar antes de que la fonola se bloquee. Esto lo usan los " + _
            "dueños de las fonolas para que no les roben las fonolas las personas" + _
            " que las ponen al publico.%99%")
    Else
        chkValid.Value = 0
        txtEstadoValidacion = TR.Trad("Estado de validación: " + vbCrLf + _
        "  * No hay bloqueos pendientes%99%")
    End If
    
    txtUSOS = ValidarCada
    txtPreAviso = AvisarAntes
    tPC = LeerConfig("IdentPcValid", "No Identificada")
    
    tERR.Anotar "acmv"
    'mostrar el registro diario de contador
    Dim TE2 As TextStream
    Set TE2 = fso.OpenTextFile(GPF("rdcday"), ForReading, False)
        Dim TodoTe2 As String
        TodoTe2 = TE2.ReadAll
    TE2.Close
    
    txtRegistroDiario = TR.Trad("REGISTRO DE ACTIVIDADES DE LOS CONTADORES CADA VEZ " + _
        "QUE INICIA 3PM" + vbCrLf + vbCrLf + _
        "Contador 'R' es el reiniciable y contador 'H' es el historico.%99%") + vbCrLf + vbCrLf + _
        TodoTe2
    
    tCONT.Text = STRceros(CONTADOR2, 11)
    
    Text2.Text = TR.Trad("¿Como proteger mi equipo al rentarlo ?" + vbCrLf + _
        "3PM cuenta con un sistema de bloqueos diferidos según cantidades de " + _
        "creditos cargados." + vbCrLf + _
        "En primer lugar debe activar la opción 'Bloquear el equipo segun usos'. " + _
        "De esta forma si el equipo no esta bajo su administración podrá asegurarse " + _
        "que deban contactarlo periodicamente para validar el uso de la rockola." + vbCrLf + _
        "Este sistema de seguridad funciona a base de creditos usados, la casilla " + _
        "'cantidad de creditos a los que se bloquera' indica cuantos créditos se " + _
        "cargaran antes de que el equipo se bloquee. Tenga en cuenta para esto " + _
        "que cada moneda puede representar más de un crédito. Esto se especifica " + _
        "en la sección precios de la configuración.%99%") + vbCrLf + _
        TR.Trad("Los 'Creditos de preaviso' son los de anticipación al bloqueo del equipo. " + _
        "Aqui le aparecerá al usuario una pantalla indicando que pida a usted la " + _
        "clave. Por ejemplo puede poner 4000 creditos con 400 de preaviso, de esta " + _
        "forma cuando pasen 3600 creditos cada vez que inicie aparecera un cartel " + _
        "solicitando clave. Esta se podrá omitir pero cuando llegue a los 4000 ya " + _
        "no podrá qudará bloqueda.%99%") + vbCrLf + _
        TR.Trad("El boton 'Generar un archivo de claves' creara la lista de codigos y claves " + _
        "correspondientes y le pedirá una ubicación para grabar el archivo de claves " + _
        "desencriptado. Un pen-drive será una buena opción. Es muy recomendable " + _
        "imprimirlo. Es un archivo de y le servirá para responder rápidamente cuando " + _
        "le pidan una clave desde esta pc. Para evitar confusiones el documento " + _
        "incluye el texto que haya escrito 'texto para recordar el equipo' de forma " + _
        "que sabrá que claves dar a cada cliente si tiene más de un equipo con " + _
        "diferentes usuarios.%99%")
    
    chkValid_Click
    
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub XxBoton2_Click()
    CreditosValidar = 0
    EscribirArch1Linea GPF("radliv"), "0"
    
    If VALIDAR Then
        chkValid.Value = 1
        TR.SetVars CreditosValidar, ValidarCada, ValidarCada - CreditosValidar
            
        txtEstadoValidacion.Text = TR.Trad("Estado de validación: " + vbCrLf + _
            "Creditos Usados: %01% de %02%" + vbCrLf + _
            "Quedan: %03%" + vbCrLf + _
            "Codigo Actual: %04% %98%La validacion es una cantidad de canciones que" + _
            " se pueden escuchar antes de que la fonola se bloquee. Esto lo usan los " + _
            "dueños de las fonolas para que no les roben las fonolas las personas" + _
            " que las ponen al publico.%99%")
    Else
        chkValid.Value = 0
        txtEstadoValidacion = TR.Trad("Estado de validación: " + vbCrLf + _
            "  * No hay bloqueos pendientes%99%")
    End If
    
End Sub

Private Sub XxBoton4_Click()
    
    'si no se crea un archivo de claves no permitir!!!
    If fso.FileExists(GPF("dalivmp2")) = False Then
        MsgBox TR.Trad("No creo el archivo de claves! No se grabara!%98%Para " + _
            "empezar la validación primero debe crear el archivo de claves%99%")
        Exit Sub
    End If
    
    'ver cuando se crea el codigo para validar....
    CrearNuevoCodigoValidar
    
    tERR.Anotar "aclw"
    'validacion con clave cada x creditos
    ChangeConfig "Validar", CStr(chkValid)
    ChangeConfig "ValidarCada", txtUSOS
    ChangeConfig "AvisarAntes", txtPreAviso
    
    MsgBox TR.Trad("Los cambios se han guardado ok y se ha creado un numero " + _
        "nuevo de validacion%99%")
    
End Sub

Private Sub XxBoton1_Click()
    
    If tPC.Text = "" Then
        MsgBox TR.Trad("No definio el texto para recordar la PC. No puede seguir%99%")
        Exit Sub
    End If
    
    If IsNumeric(txtUSOS.Text) Then
        If CLng(txtUSOS.Text) = 0 Then
            MsgBox TR.Trad("No puede usar valores en cero%99%")
            Exit Sub
        End If
    Else
        MsgBox TR.Trad("Use numeros !%99%")
        Exit Sub
    End If
    
    If IsNumeric(txtPreAviso.Text) Then
        If CLng(txtPreAviso.Text) = 0 Then
            MsgBox TR.Trad("No puede usar valores en cero%99%")
            Exit Sub
        End If
    Else
        MsgBox TR.Trad("Use numeros!%99%")
        Exit Sub
    End If
    
    Dim T1 As Long, T2 As Long
    
    Dim Te444 As TextStream
    'Dim TeTest As TextStream
    
    Dim CM2 As New CommonDialog
    CM2.DialogTitle = TR.Trad("Indique donde se grabara el archivo desencriptado " + _
        "de respuestas%99%")
    CM2.DialogPrompt = TR.Trad("Indique donde se grabara el archivo " + _
        "desencriptado de respuestas%99%")
    
    CM2.ShowSave
    
    Dim FF2 As String
    FF2 = CM2.FileName
    If FF2 = "" Then
        MsgBox TR.Trad("No eligio archivo, no se seguirá%99%")
        Exit Sub
    End If
    
    If fso.FileExists(GPF("dalivmp2")) Then fso.DeleteFile GPF("dalivmp2"), True
    
    Dim Valores() As Long, Valores2() As Long
    Dim ValoresS() As String, ValoresS2() As String
    
    Dim v As Long, V2 As Long
    Randomize
    V2 = CLng(Rnd * 66) + 100 'cantidad de valores de validación
    
    Set Te444 = fso.CreateTextFile(GPF("dalivmp2"), True)
    
        Dim Renglon As String, R1 As String, R2 As String, R3 As String
        'escribir primero las validaciones que se van a usar y el texto recordatorio
        
        R1 = CompleteSTR(CLng(txtUSOS) * 8, 8)
        R2 = CompleteSTR(CLng(txtPreAviso) * 6, 8)
        'ahora las letras enumeradas
        Dim lR As Long, lR2 As Long
        R3 = CompleteSTR(Len(tPC.Text), 2) 'indico el largo para que sepa hasta donde leer
        For lR = 1 To Len(tPC)
            lR2 = Asc(Mid(tPC, lR, 1)) * lR 'valor ascii de la letra (lr es maximo 15 segun maxLenght
            R3 = R3 + CompleteSTR(lR2, 4) 'letra pasada a 4 digitos en string
        Next lR
        
        Renglon = Mid(R1, 4, 1) + _
                  Mid(R2, 8, 1) + _
                  Mid(R1, 1, 1) + _
                  Mid(R2, 1, 1) + _
                  Mid(R1, 8, 1) + _
                  Mid(R2, 2, 1) + _
                  Mid(R1, 6, 1) + _
                  Mid(R2, 3, 1) + _
                  Mid(R1, 3, 1) + _
                  Mid(R2, 5, 1) + _
                  Mid(R1, 7, 1) + _
                  Mid(R2, 4, 1) + _
                  Mid(R1, 5, 1) + _
                  Mid(R2, 6, 1) + _
                  Mid(R1, 2, 1) + _
                  Mid(R2, 7, 1) + R3


        Te444.Write Renglon
        
        'Set TeTest = FSO.CreateTextFile(AP + "t.t", True)
        
        For v = 1 To V2
            ReDim Preserve Valores(v): ReDim Preserve Valores2(v)
            ReDim Preserve ValoresS(v): ReDim Preserve ValoresS2(v)
            
            Randomize: Valores(v) = CLng(Rnd * 7483648)
            Randomize: Valores2(v) = CLng(Rnd * 5999999)
            
            ValoresS(v) = CompleteSTR(Valores(v), 8)
            ValoresS2(v) = CompleteSTR(Valores2(v), 8)
            
            'TeTest.WriteLine ValoresS(V) + " = " + ValoresS2(V)
            
            Renglon = Mid(ValoresS2(v), 4, 1) + _
                      Mid(ValoresS(v), 1, 1) + _
                      Mid(ValoresS2(v), 1, 1) + _
                      Mid(ValoresS(v), 6, 1) + _
                      Mid(ValoresS2(v), 2, 1) + _
                      Mid(ValoresS(v), 5, 1) + _
                      Mid(ValoresS2(v), 5, 1) + _
                      Mid(ValoresS2(v), 6, 1) + _
                      Mid(ValoresS2(v), 3, 1) + _
                      Mid(ValoresS(v), 8, 1) + _
                      Mid(ValoresS(v), 3, 1) + _
                      Mid(ValoresS2(v), 8, 1) + _
                      Mid(ValoresS2(v), 7, 1) + _
                      Mid(ValoresS(v), 7, 1) + _
                      Mid(ValoresS(v), 2, 1) + _
                      Mid(ValoresS(v), 4, 1)
            
            Te444.Write Renglon
            
        Next v
    
    Te444.Close
    'TeTest.Close
    
    
    'traducir y dejar para que el usuario quite el archivo de texto
    'estan separados porque eran dos botones ya que las funciones de desencritar deben ser
    'efectivas leyendo el archivo encriptado y no como en TeTes
    
    'archivo que eligio para grabar
    If fso.FileExists(FF2) Then fso.DeleteFile FF2, True 'archivo desencriptado
    
    Dim TX As String
    Dim TE As TextStream
    
    Set TE = fso.OpenTextFile(GPF("dalivmp2"))
        TX = TE.ReadAll
    TE.Close
    
    Dim pos As Long 'posicion del archivo que voy leyendo
    pos = 1
    'las primeras 16 son dos numeros de 8 mezclados que informan de cuantos creditos
    'se valida y con que preaviso
    Dim TP As String, TP2 As String, TP3 As String 'temporales
    
    TP = Mid(TX, pos, 16)
    TP2 = Mid(TP, 3, 1) + Mid(TP, 15, 1) + Mid(TP, 9, 1) + Mid(TP, 1, 1) + _
          Mid(TP, 13, 1) + Mid(TP, 7, 1) + Mid(TP, 11, 1) + Mid(TP, 5, 1)
          
    Dim Usos As Long
    Usos = CLng(TP2 / 8)
          
    TP2 = Mid(TP, 4, 1) + Mid(TP, 6, 1) + Mid(TP, 8, 1) + Mid(TP, 12, 1) + _
          Mid(TP, 10, 1) + Mid(TP, 14, 1) + Mid(TP, 16, 1) + Mid(TP, 2, 1)
          
    Dim PreAviso As Long
    PreAviso = CLng(TP2 / 6)
    
    pos = pos + 16
    
    'los siguentes 2 digitos especifican el largo del texto
    
    Dim LN As Long, LN2 As Long, LN3 As Long 'temporales
    
    Dim RecPC As String
    RecPC = ""
    TP = Mid(TX, pos, 2)
    LN = CLng(TP)
    pos = pos + 2
    For LN2 = 0 To LN - 1 'cuatro numeros cada letra
        LN3 = CLng(Mid(TX, pos + (LN2 * 4), 4)) / (LN2 + 1)
        TP2 = Chr(LN3)
        RecPC = RecPC + TP2
    Next LN2
      
    pos = pos + (LN * 4)
    'listo ahora solo los numeros. cada 16 hay 2 grupos de 8 encriptados
    
    List1.Clear 'minga lo voy a ordenar. Uso un listbox
    
    For LN = pos - 1 To (Len(TX) - 16) Step 16
        TP = Mid(TX, LN + 2, 1) + Mid(TX, LN + 15, 1) + Mid(TX, LN + 11, 1) + Mid(TX, LN + 16, 1) + _
             Mid(TX, LN + 6, 1) + Mid(TX, LN + 4, 1) + Mid(TX, LN + 14, 1) + Mid(TX, LN + 10, 1)
        TP2 = Mid(TX, LN + 3, 1) + Mid(TX, LN + 5, 1) + Mid(TX, LN + 9, 1) + Mid(TX, LN + 1, 1) + _
             Mid(TX, LN + 7, 1) + Mid(TX, LN + 8, 1) + Mid(TX, LN + 13, 1) + Mid(TX, LN + 12, 1)
            
        TP3 = NumToTec(TP2)
        
        'ordenar por numero!!!
        List1.AddItem TP + " = " + TP2 + " = " + TP3

    Next LN
    
    'los unicos botones que considero que existen si o si so izquierda, derecha y OK
    'ademas o existen el arriba - abajo o existe el salir
    'el ok y la insercion de moneda esta si o si tambien
    
    'los reemplazos son:
    '0 izq
    '1 der
    '2 der
    '3 izq
    '4 der
    '5 der
    '6 izq
    '7 izq
    '8 izq
    '9 der
    
    'el ok queda reservado para dar fin a la cadena
    'la insercion de moneda la dejamo

    Dim VALS As String: VALS = ""
    For LN = 0 To List1.ListCount - 1
        VALS = VALS + List1.List(LN) + vbCrLf
    Next LN

    Dim Te445 As TextStream
    Set Te445 = fso.CreateTextFile(FF2, True)
        TR.SetVars RecPC
        Te445.WriteLine TR.Trad("DETALLE DE VALIDACION DE EQUIPO: %01%%98%La variable es el nombre del equipo%99%")
        Te445.WriteLine "-----------------------------------------"
        
        Te445.WriteLine TR.Trad("Contador Historico al grabarse: %99%") + STRceros(CONTADOR2, 11)
        TR.SetVars CreditosValidar
        Te445.WriteLine TR.Trad("Ya pasaron: %01%%98%La variable son la cantidad de" + _
            "créditos que pasaron en la validación%99%")
        Te445.WriteLine "-----------------------------------------"
        TR.SetVars Usos
        Te445.WriteLine TR.Trad("Cantidad de usos: %01%%99%")
        TR.SetVars PreAviso
        Te445.WriteLine TR.Trad("Preaviso: %01%%99%")
        Te445.WriteLine "-----------------------------------------"
        Te445.WriteLine
        Te445.WriteLine TR.Trad("Valores que consultaran / respuestas%99%")
        Te445.WriteLine
        Te445.WriteLine VALS
        
    Te445.Close
    
    
    MsgBox TR.Trad("Se ha creado un archivo nuevo de seguridad sin problemas" + vbCrLf + _
           "Esto no afecta la continuidad de la validación de usos ni la reinicia." + vbCrLf + _
           "Solo cambia las claves que se van a pedir cuando corresponda.%99%")
End Sub

Private Function CompleteSTR(Num As Long, HastaDig As Long) As String
    Dim SN As String
    SN = CStr(Num)
    
    If Len(SN) < HastaDig Then
        SN = String(HastaDig - Len(SN), "0") + SN
    End If
    
    CompleteSTR = SN
End Function

Private Sub XxBoton3_Click()
    Unload Me
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    txtRegistroDiario.Text = TR.Trad("Registro diario de actividades%99%")
    chkValid.Caption = TR.Trad("Bloquear el equipo segun usos%99%")
    XxBoton1.Caption = TR.Trad("Generar un archivo de claves%98%Para que el dueño " + _
        "de la fonola pueda saber con clave responder al momento de validarlo se " + _
        "crea aquí un archivo de texto con codigos y sus correspondientes claves%99%")
    XxBoton4.Caption = TR.Trad("GRABAR CAMBIOS%99%")
    XxBoton2.Caption = TR.Trad("Reinicializar conteo%99%")
    Label3(1).Caption = TR.Trad("Valor Actual del contador histórico%99%")
    Label3(0).Caption = TR.Trad("texto para recordar el equipo%99%")
    Label2.Caption = TR.Trad("Creditos de preaviso%99%")
    Label1(0).Caption = TR.Trad("cantidad de creditos a los que se bloquera%99%")
    XxBoton3.Caption = TR.Trad("Salir%99%")
End Sub
