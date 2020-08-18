VERSION 5.00
Object = "{AC1ACB77-BE60-49F4-BE38-2F9A87F5E5E4}#2.0#0"; "tbrX_Boton II.ocx"
Begin VB.Form frmVALID 
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
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      BeginProperty Font 
         Name            =   "Trebuchet MS"
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
      Height          =   645
      Left            =   7470
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Bloquear el equipo segun usos"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   4290
      TabIndex        =   0
      Top             =   300
      Width           =   2985
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
         Left            =   210
         TabIndex        =   15
         Text            =   "0"
         Top             =   2670
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
         Left            =   780
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1950
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
         Left            =   1350
         MaxLength       =   6
         TabIndex        =   4
         Text            =   "0"
         Top             =   1290
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
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   2
         Text            =   "0"
         Top             =   630
         Width           =   1395
      End
      Begin tbrX_Boton2.XxBoton XxBoton1 
         Height          =   405
         Left            =   750
         TabIndex        =   6
         Top             =   3210
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         xFColor         =   16777215
         xBColor         =   64
         xCapt           =   "Generar un archivo de claves"
         xEnabled        =   -1  'True
      End
      Begin tbrX_Boton2.XxBoton XxBoton4 
         Height          =   405
         Left            =   750
         TabIndex        =   12
         Top             =   4050
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         xFColor         =   16777215
         xBColor         =   128
         xCapt           =   "GRABAR CAMBIOS"
         xEnabled        =   0   'False
      End
      Begin tbrX_Boton2.XxBoton XxBoton2 
         Height          =   405
         Left            =   750
         TabIndex        =   17
         Top             =   3630
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   714
         xFColor         =   16777215
         xBColor         =   64
         xCapt           =   "Reinicializar conteo"
         xEnabled        =   -1  'True
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
         Left            =   480
         TabIndex        =   16
         Top             =   2430
         Width           =   3195
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "texto para recordar el equipo"
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
         Left            =   450
         TabIndex        =   9
         Top             =   1680
         Width           =   3195
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Creditos de preaviso"
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
         Left            =   1170
         TabIndex        =   5
         Top             =   1050
         Width           =   1785
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "cantidad de creditos a los que se bloquera"
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
         Left            =   450
         TabIndex        =   3
         Top             =   390
         Width           =   3405
      End
   End
   Begin tbrX_Boton2.XxBoton XxBoton3 
      Height          =   645
      Left            =   2580
      TabIndex        =   11
      Top             =   5220
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   1138
      xFColor         =   16777215
      xBColor         =   64
      xCapt           =   "Salir"
      xEnabled        =   -1  'True
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
    XxBoton4.Enabled = txtUSOS.Enabled
End Sub

Private Sub Form_Load()
    'validar cada X Creditos
    Validar = LeerConfig("Validar", "0")
    ValidarCada = LeerConfig("ValidarCada", "3000")
    AvisarAntes = LeerConfig("AvisarAntes", "500")
    
    If Validar Then
        chkValid.Value = 1
        txtEstadoValidacion = "Estado de validación: " + vbCrLf + _
            "Creditos Usados: " + CStr(CreditosValidar) + " de " + CStr(ValidarCada) + vbCrLf + _
            " Quedan: " + CStr(ValidarCada - CreditosValidar) + vbCrLf + _
            " Codigo Actual: " + CodigoParaClaveActual
    Else
        chkValid.Value = 0
        txtEstadoValidacion = "Estado de validación: " + vbCrLf + "  * No hay bloqueos pendientes"
    End If
    
    txtUSOS = ValidarCada
    txtPreAviso = AvisarAntes
    tPC = LeerConfig("IdentPcValid", "No Identificada")
    
    tERR.Anotar "acmv"
    'mostrar el registro diario de contador
    Dim TE2 As TextStream
    Set TE2 = FSO.OpenTextFile(GPF("rdcday"), ForReading, False)
        Dim TodoTe2 As String
        TodoTe2 = TE2.ReadAll
    TE2.Close
    
    txtRegistroDiario = "REGISTRO DE ACTIVIDADES DE LOS CONTADORES CADA VEZ QUE INICIA 3PM" + vbCrLf + vbCrLf + _
        "Contador 'R' es el reiniciable y contador 'H' es el historico." + vbCrLf + vbCrLf + _
        TodoTe2
    
    tCONT.Text = STRceros(CONTADOR2, 11)
    
    Text2.Text = "¿Como proteger mi equipo al rentarlo ?" + vbCrLf + _
        "3PM cuenta con un sistema de bloqueos diferidos según cantidades de creditos cargados." + vbCrLf + _
        "En primer lugar debe activar la opción 'Bloquear el equipo segun usos'. " + _
        "De esta forma si el equipo no esta bajo su administración podrá asegurarse que deban " + _
        "contactarlo periodicamente para validar el uso de la rockola." + vbCrLf + _
        "Este sistema de seguridad funciona a base de creditos usados, la casilla " + _
        "'cantidad de creditos a los que se bloquera' indica cuantos créditos se cargaran " + _
        "antes de que el equipo se bloquee. Tenga en cuenta para esto que cada moneda puede representar " + _
        "más de un crédito. Esto se especifica en la sección precios de la configuración." + vbCrLf + _
        "Los 'Creditos de preaviso' son los de anticipación al bloqueo del equipo. Aqui le aparecerá " + _
        "al usuario una pantalla indicando que pida a usted la clave. Por ejemplo puede poner 4000 " + _
        "creditos con 400 de preaviso, de esta forma cuando pasen 3600 creditos cada vez que inicie " + _
        "aparecera un cartel solicitando clave. Esta se podrá omitir pero cuando llegue a los 4000 ya " + _
        "no podrá qudará bloqueda." + vbCrLf + _
        "El boton 'Generar un archivo de claves' creara la lista de codigos y claves correspondientes " + _
        " y le pedirá una ubicación para grabar el archivo de claves desencriptado. Un pen-drive será " + _
        "una buena opción. Es muy recomendable imprimirlo. Es un archivo de y le servirá para responder " + _
        "rápidamente cuando le pidan una clave desde esta pc. Para evitar confusiones el documento " + _
        "incluye el texto que haya escrito 'texto para recordar el equipo' de forma que sabrá que claves dar " + _
        "a cada cliente si tiene más de un equipo con diferentes usuarios."
    
    chkValid_Click
    
End Sub

Private Sub XxBoton2_Click()
    CreditosValidar = 0
    EscribirArch1Linea GPF("radliv"), "0"
    
    If Validar Then
        chkValid.Value = 1
        txtEstadoValidacion = "Estado de validación: " + vbCrLf + _
            "Creditos Usados: " + CStr(CreditosValidar) + " de " + CStr(ValidarCada) + vbCrLf + _
            " Quedan: " + CStr(ValidarCada - CreditosValidar) + vbCrLf + _
            " Codigo Actual: " + CodigoParaClaveActual
    Else
        chkValid.Value = 0
        txtEstadoValidacion = "Estado de validación: " + vbCrLf + "  * No hay bloqueos pendientes"
    End If
    
End Sub

Private Sub XxBoton4_Click()
    
    'si no se crea un archivo de claves no permitir!!!
    If FSO.FileExists(GPF("dalivmp2")) = False Then
        MsgBox "No creo el archivo de claves! No se grabara!"
        Exit Sub
    End If
    
    'ver cuando se crea el codigo para validar....
    CrearNuevoCodigoValidar
    
    tERR.Anotar "aclw"
    'validacion con clave cada x creditos
    ChangeConfig "Validar", CStr(chkValid)
    ChangeConfig "ValidarCada", txtUSOS
    ChangeConfig "AvisarAntes", txtPreAviso
    
    MsgBox "Los cambios se han guardado ok y se ha creado un numero nuevo de validacion"
    
End Sub

Private Sub XxBoton1_Click()
    
    If tPC.Text = "" Then
        MsgBox "No definio el texto para recordar la PC. No puede seguir"
        Exit Sub
    End If
    
    If IsNumeric(txtUSOS.Text) Then
        If CLng(txtUSOS.Text) = 0 Then
            MsgBox "No puede usar valores en cero"
            Exit Sub
        End If
    Else
        MsgBox "Use numeros !"
        Exit Sub
    End If
    
    If IsNumeric(txtPreAviso.Text) Then
        If CLng(txtPreAviso.Text) = 0 Then
            MsgBox "No puede usar valores en cero"
            Exit Sub
        End If
    Else
        MsgBox "Use numeros !"
        Exit Sub
    End If
    
    Dim T1 As Long, T2 As Long
    
    Dim Te444 As TextStream
    'Dim TeTest As TextStream
    
    Dim CM2 As New CommonDialog
    CM2.DialogTitle = "Indique donde se grabara el archivo desencriptado de respuestas"
    CM2.DialogPrompt = "Indique donde se grabara el archivo desencriptado de respuestas"
    
    CM2.ShowSave
    
    Dim FF2 As String
    FF2 = CM2.FileName
    If FF2 = "" Then
        MsgBox "No eligio archivo, no se seguirá"
        Exit Sub
    End If
    
    If FSO.FileExists(GPF("dalivmp2")) Then FSO.DeleteFile GPF("dalivmp2"), True
    
    Dim Valores() As Long, Valores2() As Long
    Dim ValoresS() As String, ValoresS2() As String
    
    Dim V As Long, V2 As Long
    Randomize
    V2 = CLng(Rnd * 66) + 100 'cantidad de valores de validación
    
    Set Te444 = FSO.CreateTextFile(GPF("dalivmp2"), True)
    
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
        
        For V = 1 To V2
            ReDim Preserve Valores(V): ReDim Preserve Valores2(V)
            ReDim Preserve ValoresS(V): ReDim Preserve ValoresS2(V)
            
            Randomize: Valores(V) = CLng(Rnd * 7483648)
            Randomize: Valores2(V) = CLng(Rnd * 5999999)
            
            ValoresS(V) = CompleteSTR(Valores(V), 8)
            ValoresS2(V) = CompleteSTR(Valores2(V), 8)
            
            'TeTest.WriteLine ValoresS(V) + " = " + ValoresS2(V)
            
            Renglon = Mid(ValoresS2(V), 4, 1) + _
                      Mid(ValoresS(V), 1, 1) + _
                      Mid(ValoresS2(V), 1, 1) + _
                      Mid(ValoresS(V), 6, 1) + _
                      Mid(ValoresS2(V), 2, 1) + _
                      Mid(ValoresS(V), 5, 1) + _
                      Mid(ValoresS2(V), 5, 1) + _
                      Mid(ValoresS2(V), 6, 1) + _
                      Mid(ValoresS2(V), 3, 1) + _
                      Mid(ValoresS(V), 8, 1) + _
                      Mid(ValoresS(V), 3, 1) + _
                      Mid(ValoresS2(V), 8, 1) + _
                      Mid(ValoresS2(V), 7, 1) + _
                      Mid(ValoresS(V), 7, 1) + _
                      Mid(ValoresS(V), 2, 1) + _
                      Mid(ValoresS(V), 4, 1)
            
            Te444.Write Renglon
            
        Next V
    
    Te444.Close
    'TeTest.Close
    
    
    'traducir y dejar para que el usuario quite el archivo de texto
    'estan separados porque eran dos botones ya que las funciones de desencritar deben ser
    'efectivas leyendo el archivo encriptado y no como en TeTes
    
    'archivo que eligio para grabar
    If FSO.FileExists(FF2) Then FSO.DeleteFile FF2, True 'archivo desencriptado
    
    Dim TX As String
    Dim TE As TextStream
    
    Set TE = FSO.OpenTextFile(GPF("dalivmp2"))
        TX = TE.ReadAll
    TE.Close
    
    Dim Pos As Long 'posicion del archivo que voy leyendo
    Pos = 1
    'las primeras 16 son dos numeros de 8 mezclados que informan de cuantos creditos
    'se valida y con que preaviso
    Dim TP As String, TP2 As String, TP3 As String 'temporales
    
    TP = Mid(TX, Pos, 16)
    TP2 = Mid(TP, 3, 1) + Mid(TP, 15, 1) + Mid(TP, 9, 1) + Mid(TP, 1, 1) + _
          Mid(TP, 13, 1) + Mid(TP, 7, 1) + Mid(TP, 11, 1) + Mid(TP, 5, 1)
          
    Dim Usos As Long
    Usos = CLng(TP2 / 8)
          
    TP2 = Mid(TP, 4, 1) + Mid(TP, 6, 1) + Mid(TP, 8, 1) + Mid(TP, 12, 1) + _
          Mid(TP, 10, 1) + Mid(TP, 14, 1) + Mid(TP, 16, 1) + Mid(TP, 2, 1)
          
    Dim PreAviso As Long
    PreAviso = CLng(TP2 / 6)
    
    Pos = Pos + 16
    
    'los siguentes 2 digitos especifican el largo del texto
    
    Dim LN As Long, LN2 As Long, LN3 As Long 'temporales
    
    Dim RecPC As String
    RecPC = ""
    TP = Mid(TX, Pos, 2)
    LN = CLng(TP)
    Pos = Pos + 2
    For LN2 = 0 To LN - 1 'cuatro numeros cada letra
        LN3 = CLng(Mid(TX, Pos + (LN2 * 4), 4)) / (LN2 + 1)
        TP2 = Chr(LN3)
        RecPC = RecPC + TP2
    Next LN2
      
    Pos = Pos + (LN * 4)
    'listo ahora solo los numeros. cada 16 hay 2 grupos de 8 encriptados
    
    List1.Clear 'minga lo voy a ordenar. Uso un listbox
    
    For LN = Pos - 1 To (Len(TX) - 16) Step 16
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
    Set Te445 = FSO.CreateTextFile(FF2, True)
    
        Te445.WriteLine "DETALLE DE VALIDACION DE EQUIPO: " + RecPC
        Te445.WriteLine "-----------------------------------------"
        Te445.WriteLine "Contador Historico al grabarse: " + STRceros(CONTADOR2, 11)
        Te445.WriteLine "Ya pasaron: " + CStr(CreditosValidar)
        Te445.WriteLine "-----------------------------------------"
        Te445.WriteLine "Cantidad de usos: " + CStr(Usos)
        Te445.WriteLine "Preaviso: " + CStr(PreAviso)
        Te445.WriteLine "-----------------------------------------"
        Te445.WriteLine
        Te445.WriteLine "Valores que consultaran / respuestas"
        Te445.WriteLine
        Te445.WriteLine VALS
        
    Te445.Close
    
    
    MsgBox "Se ha creado un archivo nuevo de seguridad sin problemas" + vbCrLf + _
           "Esto no afecta la continuidad de la validacion de usos ni la reinicia." + vbCrLf + _
           "Solo cambia las claves que se van a pedir cuando corresponda."
End Sub

Private Function CompleteSTR(Num As Long, HastaDig As Long) As String
    Dim sN As String
    sN = CStr(Num)
    
    If Len(sN) < HastaDig Then
        sN = String(HastaDig - Len(sN), "0") + sN
    End If
    
    CompleteSTR = sN
End Function

Private Sub XxBoton3_Click()
    Unload Me
End Sub
