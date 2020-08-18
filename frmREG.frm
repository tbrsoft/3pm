VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de 3PM"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Cancel          =   -1  'True
      Caption         =   "COMPRAR AHORA!"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7470
      Width           =   3720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir Manual de uso. RECOMENDADO si es su primer uso"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2460
      TabIndex        =   24
      Top             =   6360
      Width           =   6915
   End
   Begin VB.Frame frFull 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2385
      Left            =   120
      TabIndex        =   15
      Top             =   60
      Visible         =   0   'False
      Width           =   11565
      Begin VB.TextBox txtEmp3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4980
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   840
         Width           =   1665
      End
      Begin VB.TextBox txtEMP1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5010
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   120
         Width           =   1665
      End
      Begin VB.TextBox txtGenCod 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   810
         Width           =   4800
      End
      Begin VB.TextBox txtGenCodSL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1170
         Width           =   4800
      End
      Begin VB.TextBox txtGenCodMIN 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   450
         Width           =   4800
      End
      Begin VB.TextBox TXTcODIGOminimo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   90
         Width           =   4800
      End
      Begin VB.TextBox txtCodSL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   810
         Width           =   4800
      End
      Begin VB.TextBox txtEmp2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   5010
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Width           =   1665
      End
      Begin VB.ListBox lstArchReg 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   150
         TabIndex        =   19
         Top             =   1530
         Width           =   11235
      End
      Begin VB.TextBox txtCodToFind 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   60
         Width           =   3165
      End
      Begin VB.CommandButton cmdGENERATE 
         Caption         =   "rareneg"
         Height          =   315
         Left            =   3330
         TabIndex        =   17
         Top             =   120
         Width           =   1095
      End
      Begin VB.TextBox txtCodGenerado 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   6690
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   450
         Width           =   4800
      End
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404000&
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia de Usuario Final"
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
      Height          =   375
      Left            =   1950
      TabIndex        =   23
      Top             =   5430
      Value           =   1  'Checked
      Width           =   8655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Ver CLUF"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4770
      TabIndex        =   22
      Top             =   5820
      Width           =   1785
   End
   Begin VB.TextBox LBL 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
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
      Height          =   2295
      Left            =   1020
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "frmREG.frx":0000
      Top             =   60
      Width           =   8295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      TabIndex        =   10
      Top             =   6930
      Width           =   1785
   End
   Begin VB.ComboBox cmbCountry 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4410
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4890
      Width           =   2805
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   7
      Left            =   8970
      MaxLength       =   5
      TabIndex        =   7
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   6
      Left            =   7890
      MaxLength       =   5
      TabIndex        =   6
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   5
      Left            =   6810
      MaxLength       =   5
      TabIndex        =   5
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   4
      Left            =   5730
      MaxLength       =   5
      TabIndex        =   4
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   3
      Left            =   4650
      MaxLength       =   5
      TabIndex        =   3
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   2
      Left            =   3570
      MaxLength       =   5
      TabIndex        =   2
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   1
      Left            =   2490
      MaxLength       =   5
      TabIndex        =   1
      Top             =   3600
      Width           =   1050
   End
   Begin VB.TextBox lblGUID 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      Text            =   "Aqui va el codigo"
      Top             =   2700
      Width           =   11535
   End
   Begin VB.TextBox txtCOD 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Index           =   0
      Left            =   1410
      MaxLength       =   5
      TabIndex        =   0
      Top             =   3600
      Width           =   1050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   9
      Top             =   6930
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG.frx":0006
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   705
      Index           =   3
      Left            =   480
      TabIndex        =   20
      Top             =   3990
      Width           =   10605
   End
   Begin VB.Image Image1 
      Height          =   1635
      Left            =   10110
      Picture         =   "frmREG.frx":009C
      Stretch         =   -1  'True
      Top             =   390
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Indique su pais de residencia"
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
      Height          =   255
      Index           =   2
      Left            =   4440
      TabIndex        =   14
      Top             =   4680
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Si ya recibio el código de tbrSoft cárguelo aqui. Respete mayúsculas y minúsculas!!"
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
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   13
      Top             =   3360
      Width           =   11505
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo a enviar a tbrSoft. Valor unico para esta PC"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Index           =   0
      Left            =   1890
      TabIndex        =   12
      Top             =   2400
      Width           =   8355
   End
End
Attribute VB_Name = "frmREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Errores As Long 'Veces que se erro la contraseña
Dim UsosDemo As Long

Dim STRconCodigos As String 'variable fundamental que contiene todos los codigos

Dim A1 As String, A2 As String, A3 As String, A4 As String, A5 As String

Dim GGG As String

Dim ArchGUID As String 'archivo con el primer codigo que se le pidio al usuario
Dim SysF As Folder 'ubicacion de la carpeta de Windows
Dim CarpetaSys As String


Private Sub Check1_Click()
    Command1.Enabled = Check1
End Sub

Private Sub cmdGENERATE_Click()
    'para generar deno desencriptar y luego
    txtGenCodMIN = "DEMO2: " + MostraDeA5(GenerarCodigoDemo(txtCodToFind))
    txtGenCod = "FULL: " + MostraDeA5(GenerarCodigo(txtCodToFind))
    txtGenCodSL = "SL: " + MostraDeA5(GenerarCodigoSL(txtCodToFind))
End Sub

Private Sub Command1_Click()
    Dim CodigoUnido As String
    CodigoUnido = txtCOD(0) + txtCOD(1) + txtCOD(2) + txtCOD(3) + txtCOD(4) + txtCOD(5) + txtCOD(6) + txtCOD(7)

    'hay un codigo personal (mas corto) que habilita la funcion de creado de claves
    Dim CF As String
    'If CodigoUnido = "FULL" Then
    '    frFull.Visible = True
    '   'poner el codigo terminado en las casillas
    '    CF = GenerarCodigo(lblGUID)
    '    txtCOD(0) = Mid(CF, 1, 5): txtCOD(1) = Mid(CF, 6, 5): txtCOD(2) = Mid(CF, 11, 5)
    '    txtCOD(3) = Mid(CF, 16, 5): txtCOD(4) = Mid(CF, 21, 5): txtCOD(5) = Mid(CF, 26, 5)
    '    txtCOD(6) = Mid(CF, 31, 5): txtCOD(7) = Mid(CF, 36, 5)
    '    Exit Sub
    'End If
    ''
    'if CodigoUnido = "SL" Then
    '   frFull.Visible = True
    '    'poner el codigo terminado en las casillas
    '    CF = GenerarCodigoSL(lblGUID)
    ''    txtCOD(0) = Mid(CF, 1, 5): txtCOD(1) = Mid(CF, 6, 5): txtCOD(2) = Mid(CF, 11, 5)
    '    txtCOD(3) = Mid(CF, 16, 5): txtCOD(4) = Mid(CF, 21, 5): txtCOD(5) = Mid(CF, 26, 5)
    '    txtCOD(6) = Mid(CF, 31, 5): txtCOD(7) = Mid(CF, 36, 5)
    '    Exit Sub
    'End If
    '''
    'If CodigoUnido = "DEMO2" Then
    '    frFull.Visible = True
    '    'poner el codigo terminado en las casillas
    '    CF = GenerarCodigoDemo(lblGUID)
    '    txtCOD(0) = Mid(CF, 1, 5): txtCOD(1) = Mid(CF, 6, 5): txtCOD(2) = Mid(CF, 11, 5)
    '    txtCOD(3) = Mid(CF, 16, 5): txtCOD(4) = Mid(CF, 21, 5): txtCOD(5) = Mid(CF, 26, 5)
    '    txtCOD(6) = Mid(CF, 31, 5): txtCOD(7) = Mid(CF, 36, 5)
    '    Exit Sub
    'End If
    '
    If cmbCountry = "(SELECCIONE PAIS)" Then MsgBox "Debe cargar el pais de residencia": Exit Sub
    
    If CodigoUnido = "" Then
        Dim TXTmsg As String
        TXTmsg = "3PM en version Demo tiene limite de discos, " + _
            "y trunca los temas a los 2 minutos." + vbCrLf + _
            "¿Desea ejecutar 3PM en version demo?"
        If MsgBox(TXTmsg, vbQuestion + vbYesNo, "3PM demo") = vbNo Then End
        Errores = 1
        UsosDemo = 1
        'ver si ya se abrio antes como demo
        If FSO.FileExists(ArchREG) Then
            'leer a ver que pasa
            
            Set TE = FSO.OpenTextFile(ArchREG, ForReading, False)
            A1 = TE.ReadLine 'este es el guid
            A2 = TE.ReadLine 'este es la clave enviada por tbrSoft (hay 3 distintas)
            A3 = TE.ReadLine 'ingresos demo
            A4 = TE.ReadLine 'estado actual del registro. Puede ser
                '"DEMO" todavia no ingreso contraseña
                '"FUCK". Intentos de crak
                '"FUCK OFF"'ya jodio demasiado, esta bloqueado
                '"FULL". Ya lo puede usar, esta registrado OK
            A5 = TE.ReadLine 'veces que se erro la contraseña
            
            TE.Close
            UsosDemo = Val(A3) + 1
        Else
            A5 = "0"
        End If
        'cargar el archivo de registro como demo. De todas formas se debera volver a abrir esta pantalla
        'aqu se escribe por primera vez
        Set TE = FSO.CreateTextFile(ArchREG, True)
        Dim GGG As String
        GGG = lblGUID
        TE.WriteLine GGG
        TE.WriteLine "00000"
        TE.WriteLine CStr(UsosDemo)
        TE.WriteLine "DEMO" 'estado actual del registro. Puede ser
            '"DEMO" todavia no ingreso contraseña
            '"DEMO2" ingreso la contraseña gratuita
            '"FUCK". Intentos de crack
            '"FULL". Ya lo puede usar, esta registrado OK. Licencia comun
            '"SL" es un superlicencia
        'si ya existia el archivo guarda la cantidad de errores que hubo
        'si no lo pone en 0
        TE.WriteLine CStr(A5)
            
        TE.Close
        
        If UsosDemo > 40 Then
            'no se puede iniciar infunitas demos
            MsgBox "No se puede utilizar mas de 40 veces como demo. 3PM se cerrara. Solicite " + _
                "su contraseña gratuita para seguir yusando la demo a info@tbrsoft.com"
            End
        End If
        
        TypeVersion = "DEMO"
        frmINI.Show 1
    Else
        'ver si sirve el valor devuelto
        'lblGUID es un nuevo codigo generado o si se hizo al azar es un codigo grabado en un archivo
        Dim ClaveCargada As String
        ClaveCargada = "FUCK" 'por defecto le erro
        If CodigoUnido = GenerarCodigoDemo(lblGUID) Then
            MsgBox "Clave gratuita de 3PM. Se desbloquea el limite de usos como demo"
            ClaveCargada = "DEMO2"
        End If
        If CodigoUnido = GenerarCodigo(lblGUID) Then
            MsgBox "El codigo se ha cargado correctamente. Bienvenido a 3PM"
            ClaveCargada = "FULL"
        End If
        If CodigoUnido = GenerarCodigoSL(lblGUID) Then
            MsgBox "SUPERLICENCIA de 3PM. El codigo de SuperLicencia se ha cargado correctamente. Bienvenido a Super3PM"
            ClaveCargada = "SL"
        End If
        If ClaveCargada <> "FUCK" Then 'entro OK de alguna forma
            'cargar archivo de registro OK en esta PC
            Set TE = FSO.CreateTextFile(ArchREG, True)
            
            GGG = lblGUID
            TE.WriteLine GGG
            TE.WriteLine CodigoUnido
            TE.WriteLine CStr(UsosDemo)
            TE.WriteLine ClaveCargada
            'estado actual del registro. Puede ser
                '"DEMO" todavia no ingreso contraseña
                '"FUCK". Intentos de crak
                '"FUCK OFF". Inhabilitado
                '"FULL". Ya lo puede usar, esta registrado OK
            'ya no importa la cantidad de errores
            TE.WriteLine "0"
            
            TE.Close
            TypeVersion = ClaveCargada 'DEMO2, FULL o SL
            frmINI.Show 1
        Else
            Errores = 1
            MsgBox "El codigo que ha ingresado no es valido. Si esto se " + _
            " repite es probable que este equipo quede inhabilitado para utilizar 3PM"
            'escribir la cantidad de fallas en el archReg
            If FSO.FileExists(ArchREG) Then
                'leer a ver que pasa
                
                Set TE = FSO.OpenTextFile(ArchREG, ForReading, False)
                A1 = TE.ReadLine 'este es el guid
                A2 = TE.ReadLine 'este es la clave enviada por tbrSoft
                A3 = TE.ReadLine 'ingresos demo
                A4 = TE.ReadLine 'estado actual del registro. Puede ser
                    '"DEMO" todavia no ingreso contraseña
                    '"FUCK". Intentos de crak
                    '"FUCK OFF"'ya jodio demasiado, esta bloqueado
                    '"FULL". Ya lo puede usar, esta registrado OK
                A5 = TE.ReadLine 'cantidad de veces que se erro la contraseña
                TE.Close
                Errores = Val(A5) + 1
                
            End If
            
            'cargar el archivo de registro como demo. De todas formas se debera volver a abrir esta pantalla
            
            Set TE = FSO.CreateTextFile(ArchREG, True)
                GGG = lblGUID
                TE.WriteLine GGG
                TE.WriteLine CodigoUnido 'el codigo malo que cargo
                TE.WriteLine CStr(UsosDemo)
                If Errores > 10 Then
                    TE.WriteLine "FUCK OFF"
                Else
                    TE.WriteLine "FUCK"
                End If
                'estado actual del registro. Puede ser
                    '"DEMO" todavia no ingreso contraseña
                    '"FUCK". Intentos de crak
                    '"FULL". Ya lo puede usar, esta registrado OK
                TE.WriteLine CStr(Errores)
            TE.Close
            End
        End If
        
    End If
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Command3_Click()
    AbrirArchivo AP + "manual.doc", Me
End Sub

Private Sub Command4_Click()
    frmCLUF.Show 1
End Sub

Private Sub Command5_Click()
    frmCompraYA.Show 1
End Sub

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    AjustarFRM Me, 12000
    'se graba en win y system
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    
    If UCase(App.EXEName) <> "3PM" Then
        MsgBox "No puede cambiar el nombre del programa"
        End
    End If
    'VER SI existe el archivo con los datos de las
    'imágenes de inicio y de cierre
    Dim ArchImgIni As String
    ArchImgIni = AP + "imgini.tbr"
    If FSO.FileExists(ArchImgIni) Then
        GoTo YaEstaIMG
    Else
        Set TE = FSO.CreateTextFile(ArchImgIni, True)
        If FSO.FolderExists(WINfolder + "\img3PM") = False Then FSO.CreateFolder WINfolder + "\img3PM"
        If FSO.FolderExists(WINfolder + "\img3PM\w") = False Then FSO.CreateFolder WINfolder + "\img3pm\w"
        'ver imagen de inicio
        If FSO.FileExists("c:\logo.sys") Then
            TE.WriteLine "ImgIni=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "\img3pm\w\logo.sys") Then FSO.DeleteFile WINfolder + "\img3pm\w\logo.sys", True
            FSO.CopyFile "c:\logo.sys", WINfolder + "\img3pm\w\logo.sys", True
        Else
            TE.WriteLine "ImgIni=0"
        End If
        
        'ver imagen de cerrando
        If FSO.FileExists(WINfolder + "\logow.sys") Then
            TE.WriteLine "ImgCerrando=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "\img3pm\w\logow.sys") Then FSO.DeleteFile WINfolder + "\img3pm\w\logow.sys", True
            FSO.CopyFile WINfolder + "\logow.sys", WINfolder + "\img3pm\w\logow.sys", True
        Else
            TE.WriteLine "ImgCerrando=0"
        End If
        
        'ver imagen de apagar
        If FSO.FileExists(WINfolder + "\logos.sys") Then
            TE.WriteLine "ImgApagar=1"
            'copiar el archivo a un lugar seguro para
            'despues administrar los cambios
            If FSO.FileExists(WINfolder + "\img3pm\w\logos.sys") Then FSO.DeleteFile WINfolder + "\img3pm\w\logos.sys", True
            FSO.CopyFile WINfolder + "\logos.sys", WINfolder + "\img3pm\w\logos.sys", True
        Else
            TE.WriteLine "ImgApagar=0"
        End If
        'escribir que todas las imagenes se cargan desde windows
        TE.WriteLine "LoadImgIni=w"
        TE.WriteLine "LoadImgCerrando=w"
        TE.WriteLine "LoadImgApagar=w"
        TE.Close
    End If
    
YaEstaIMG:
    'VER si ya se pasaron las imagenes de 3pm
    'a la carpeta que corresponde
    If FSO.FolderExists(WINfolder + "\img3pm") = False Then FSO.CreateFolder (WINfolder + "\img3pm")
    If FSO.FolderExists(WINfolder + "\img3pm\3") = False Then FSO.CreateFolder (WINfolder + "\img3pm\3")
    If FSO.FileExists(AP + "logo.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "\img3pm\3\logo.sys") Then FSO.DeleteFile WINfolder + "\img3pm\3\logo.sys", True
        FSO.CopyFile AP + "logo.sys", WINfolder + "\img3pm\3\logo.sys", True
        'If FSO.FileExists(WINfolder + "\img3pm\3\logo.sys") = False Then FSO.CopyFile AP + "logo.sys", WINfolder + "\img3pm\3\logo.sys", True
    End If
    If FSO.FileExists(AP + "logow.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "\img3pm\3\logow.sys") Then FSO.DeleteFile WINfolder + "\img3pm\3\logow.sys", True
        FSO.CopyFile AP + "logow.sys", WINfolder + "\img3pm\3\logow.sys", True
    End If
    If FSO.FileExists(AP + "logos.sys") Then
        'siempre copiarlo si esta
        If FSO.FileExists(WINfolder + "\img3pm\3\logos.sys") Then FSO.DeleteFile WINfolder + "\img3pm\3\logos.sys", True
        FSO.CopyFile AP + "logos.sys", WINfolder + "\img3pm\3\logos.sys", True
    End If
    
    'SACAR los 0 y las O por la confusion (igual con las l (eles))
STRconCodigos = "87dfsdfw897564fghererh56424dfg23d4fge89r7e89rgWER7W9E8F7SD54s5d6f4sd56fw4e4w" + _
    "YUGykjgKHJBViuhyiuHHJKbkLjb567876543242324768768768jhjkhUYIULU6654654564A23E" + _
    "JF9ELFK45trF8QW78EREDF54DCDrtergdfv5asd4dfvdfgfg5sv432s4cv35sd4f88sf7d384f8s" + _
    "WUIYiuyiuyGgfgdcvxIUYPIertef54s65dfxcvxdsdf46232sd3f22WER223WE2R32WE23W2EpF32" + _
    "RE6T43D5F4G53ER4e5cvefgsdgf4v3sd5f435ssdpf3asd36r5t4354jLhjpyu534354fghd35z4s" + _
    "WE56F43SA5D4V564RT788DF4BV234RTH86RB64S35D4FG38R4BH36wer84sd54v3as54ga5er48a" + _
    "er5t4sd53fbv456564dfgwe5675s4d3fp22EG486WE68R4V3E5F4V3A36453DV436SDF4V365555" + _
    "WEFp3SAD32Vp325e454sdf3g54sd3f4553SER4G35SD4F35V43AS2pDFV32ASpCSD23A4BVNMVNM" + _
    "QWERwerqw4532sd4fpgse3g4586WEDF864AS65F4A6S5df86g4ds65fvp3s2dpf35g45rt75s4d5"
    
STRconCodigos = STRconCodigos + "yehvndis782p6655S5DLCIEUDJXLSASpXSAAAAAAQ" + _
    "UE7ENCNMSLSLPDpF8EREDSDAQWp4655SDKDICKX8p643457gfh4567ghkskdncksppwpppwwkdmx" + _
    "qwLdaqL9e9dkkLx59ejk3me9j892DJDJE899DJSDLpSD99WE23KWD992JKLWE999jqwjqw99qwdqp" + _
    "h999EDMDPWE9DLXpAASDP9DCJCM995yDaiLidi99sd9Ss9dSDFKJDF99DMCMSJS9SD9S9JXSS9S9A" + _
    "qdsjc9vIdLfyuwet2fgsh557hedfg3sgfhdfgf5yhyfhbpekefpLkefe9rjfje99e8rkefjLkd9e" + _
    "QLWLDEMDKE9ELCK9E9FE9EFA9EFK4399DUWE99JVDA9IDvbdfgS6456456456234563634568457" + _
    "ASKLJSLs9d9sd98s6s78564s34s37s8s7cyv76x8durnngic8edfgdui4jjf9dj348djk389893a" + _
    "skjhd9999S8D99CMCMGLRJEU48RJFHG8HJR8GJE89EJC9E84HJF84Jdfg345gdf84NFJKSKDKFF5"
    
STRconCodigos = STRconCodigos + "SJFU4FJKGLDpCMW949TIEMD827394A3DF939R8FJF" + _
    "W9DJDME9E9FCJChduc99d93LrLpgjkdhjd73hydbhasyxbsnvmbadfjdisdLs93dfgdnmdisd8sa" + _
    "n66d7cjkek49fmr9rfmcv9d9fnmc9dnmcv9dnc9d89eh4n4nmv9d83nf8enf83nmkLddfgaksasS9" + _
    "dfgUDMSLCLpD8894KRKF93ILJDNMMD992MDM92992MDpSpLD99299WLKCJSWLW995EJLpS95SKLpS" + _
    "ALSgfv5C9DKL3KFJLCPE953JCLLKSD95E9RJFLDLP93JDFLSLDU953KFJkLsjsLd93jfjspdjsLi" + _
    "a59dmcvf9dfgrtjcdkswLs89djd9939erjfLsd59w95sjsLpdj39rjf9535ruc95795s895sd95s" + _
    "S8D895DK3KGLFKMCNSUD5498DVHCUD7E74HDNSUW8E4HDPL539EJICJDJW59WEHJDNC59SD59JAA" + _
    "ert5545GF95394MCVJKDIW8929e95c95sLdjsLpdu95s95e893w895djsp9d589w895d895s5sss"
    
STRconCodigos = STRconCodigos + "zsmkxjkLsiLc95uu8489fndefjL349845789f89w89" + _
    "JKDHKS8989RNKDFKLFD95348953DNDNCIUIE78DHXHIASKJSDYISIYSDTFGSKJSHSJKSGCBVCG6S" + _
    "SDKDNMCNDHSUISIUDHSJKDNXCNKLSJKDHDUIWY78HDJKSDHCBNSJKSDYFKLSKLJKLSDGSDYESUDYUS" + _
    "SJKDNCN233nawesdmdjs83783678wdkusdhsikwi7w7sdghskjjkdjkSDF346LGJK48FMEKD8S83JE" + _
    "sjkdjcnsdkudi7fyeiuidjkdkLxjkchxkdustykiwLksdLkLsjcnsksuduiwiudysihcnskshjag" + _
    "asjdjcjsuyebdfnLsLduuebndndpLskLdiuiuwiueui689wiushcnkLskLjss8ILUSYS7IYDISSS" + _
    "SCCCLISWUDUW89789F5C57S76WGEBFHDGW4WTSHWH3TDBDT5CGSG5GFWHDFJSJjshstcyd63hfksbns" + _
    "sjkdnc7875Lenme89c8weienmduixhxhsyqywkdjsjkzzzzkwiwi78df6c453x4x6x7c6x5s4s5s"

    cmbCountry.AddItem "Argentina"
    cmbCountry.AddItem "(SELECCIONE PAIS)"
    cmbCountry.AddItem "España"
    cmbCountry.AddItem "Chile"
    cmbCountry.AddItem "Uruguay"
    cmbCountry.AddItem "Honduras"
    cmbCountry.AddItem "El Salvador"
    cmbCountry.AddItem "Mexico"
    cmbCountry.AddItem "Venezuela"
    cmbCountry.AddItem "Bolivia"
    cmbCountry.AddItem "Colombia"
    cmbCountry.AddItem "Perú"
    cmbCountry.AddItem "Nicaragua"
    cmbCountry.AddItem "EEUU"
    cmbCountry.AddItem "Paraguay"
    cmbCountry.AddItem "Costa Rica"
    cmbCountry.AddItem "Ecuador"
    cmbCountry.AddItem "Guatemala"
    cmbCountry.AddItem "Panama"
    cmbCountry.AddItem "Puerto Rico"
    cmbCountry.AddItem "Republica Dominicana"
    cmbCountry.ListIndex = 0
    
    
    'ver la ubicacion del archivo de registro
    Set SysF = FSO.GetSpecialFolder(SystemFolder)
    CarpetaSys = SysF.path + "\"
    ArchREG = CarpetaSys + "rmlvf.dll"
    
    'no se pone el primer componente por qyue este cambia con cada inicio
    lblGUID = GetGuidSL
    
    txtCodSL = MostraDeA5(GenerarCodigoSL(lblGUID))
    TXTcODIGOminimo = MostraDeA5(GenerarCodigoDemo(lblGUID))
    txtCodGenerado = MostraDeA5(GenerarCodigo(lblGUID))
    
    'no se registrado o se ha perdido el archivo de registro
    TXT = "Bienvenido a 3PM. Gracias por confiar en tbrSoft Argentina" + vbCrLf + vbCrLf + _
    "Puede utilizar esta version demo con algunas restricciones simplemente " + _
    "indicando su pais de residencia y presionando OK ahora" + vbCrLf + vbCrLf + _
    "El costo de 3PM con licencia para un equipo es de U$S 75" + vbCrLf + _
    "3PM no incluye en ninguna de sus licencia el derecho de venta del software. Por lo que " + _
    "solo usted tendra una copia LEGAL si compra este software a " + _
    "tbrSoft Argentina." + vbCrLf + _
    "Para adquirir la version definitiva deberá solicitarlo a tbrSoft " + _
    "via email a info@tbrsoft.com o a avazquez@cpcipc.org"
    
    LBL = TXT
    
    'si esta registrada corroborar que no sea un registro de otra PC
    If FSO.FileExists(ArchREG) = False Then
        lstArchReg.AddItem "No existe el archivo de registro (ArchReg)"
    Else
        'ver si el GUID de esta maquina coincide con el del
        'registro. Esto evita que se copie el registro de una
        'maquina a otra
        
        Dim A1 As String, A2 As String, A3 As String, A4 As String
        
        Set TE = FSO.OpenTextFile(ArchREG, ForReading, False)
        'ver si el archivo no esta vacio!!!
        If TE.AtEndOfStream Then GoTo ESDEMO
        
        A1 = TE.ReadLine 'este es el guid
        A2 = TE.ReadLine 'este es la clave enviada por tbrSoft
        A3 = TE.ReadLine 'dias de demo
        A4 = TE.ReadLine 'estado actual del registro. Puede ser
            '"DEMO" todavia no ingreso contraseña
            '"DEMO2" conraseña gratuita
            '"FUCK". Intentos de crak
            '"FUCK OFF". Inhabilitado
            '"FULL". Ya lo puede usar, esta registrado OK
            '"SL" SUPERLICENCIA
        A5 = TE.ReadLine 'veces que se erro la contraseña
            
        TE.Close
        lstArchReg.AddItem "Archivo de registro (ArchReg)"
        lstArchReg.AddItem A1
        lstArchReg.AddItem A2
        lstArchReg.AddItem A3
        lstArchReg.AddItem A4
        lstArchReg.AddItem A5
        
        'contar los usos
        Dim ArchUsos As String
        ArchUsos = WINfolder + "\slx98.dll"
        
        If FSO.FileExists(ArchUsos) Then
            'ver cuanto hay
            Set TE = FSO.OpenTextFile(ArchUsos, ForReading, False)
            Dim Usado As Long
            Usado = Val(TE.ReadLine)
            'ver si hay que parar
            If Usado > 100000 Then '1.000.000 son 550 años (5 usos por dia)
                MsgBox "Ha pasado los usos habilitados. Esta no es una version definitiva"
                End
            End If
            
            Usado = Usado + 1
            TE.Close
            
            'sumar uno
            Set TE = FSO.CreateTextFile(ArchUsos, True)
            TE.WriteLine Str(Usado)
            TE.Close
        Else
            'es el primer uso legal
            Set TE = FSO.CreateTextFile(ArchUsos, True)
            TE.WriteLine "1"
            TE.Close
        End If
        
        If A1 = lblGUID Then
        'guid actual es el numero corto para (agrandar) y solicitar a tbrSoft
            If A2 = GenerarCodigoDemo(lblGUID) Then
            'este si es corto. El sistema esta comprobando el valor directamente desde el numeo que le sirve como referencia
                If A4 = "DEMO2" Then
                'esta todo OK puede usar
                    TypeVersion = "DEMO2"
                    Unload Me
                    frmINI.Show 1
                Else
                    'crack que no sabe que debe poner "DEMO2"
                    GoTo FUCK
                End If
            End If
            
            If A2 = GenerarCodigo(lblGUID) Then
            'este si es corto. El sistema esta comprobando el valor directamente desde el numeo que le sirve como referencia
                If A4 = "FULL" Then
                'esta todo OK puede usar
                    TypeVersion = "FULL"
                    Unload Me
                    frmINI.Show 1
                Else
                    'crack que no sabe que debe poner "FULL"
                    GoTo FUCK
                End If
            End If
            
            If A2 = GenerarCodigoSL(lblGUID) Then
            'este si es corto. El sistema esta comprobando el valor directamente desde el numeo que le sirve como referencia
                If A4 = "SL" Then
                'esta todo OK puede usar
                    TypeVersion = "SL"
                    Unload Me
                    frmINI.Show 1
                Else
                    'crack que no sabe que debe poner "FULL"
                    GoTo FUCK
                End If
            End If
ESDEMO:
            'ver si esta trabajando como demo
            If A2 = "00000" Then
                If A4 = "DEMO" Then
                    'permitir ver si carga la contraseña o
                    'entra de nuevo como demo
                End If
            Else
                'no es demo y no es codigo valido
                'intento de adivinacion de codigo
                If A4 = "FUCK" Or A4 = "FULL" Then
                    GoTo FUCK
                Else
                'cualquier otro valor (pueden haber borrado el FUCK OFF"
                    MsgBox "3PM ha sido inhabilitado de este equipo, " + _
                        "no podra ser usado nuevamente ya que se ha " + _
                        "intentado usar ilegalmente"
                    End
                End If
            End If
        Else
            ' el codigo a solicitar debe ser el mismo. Ya sea al azar o no
            ' puede venir de otra PC
            GoTo FUCK
        End If
        
    End If
    Exit Sub
FUCK:
    'sumar uno en fallos
    Set TE = FSO.OpenTextFile(ArchREG, ForReading, False)
        A1 = TE.ReadLine 'este es el guid
        A2 = TE.ReadLine 'este es la clave enviada por tbrSoft
        A3 = TE.ReadLine 'ingresos demo
        A4 = TE.ReadLine 'estado actual del registro. Puede ser
            '"DEMO" todavia no ingreso contraseña
            '"FUCK". Intentos de crak
            '"FUCK OFF"'ya jodio demasiado, esta bloqueado
            '"FULL". Ya lo puede usar, esta registrado OK
        A5 = TE.ReadLine ' veces que se erro la contraseña
    TE.Close
    Errores = Val(A5) + 1
    Set TE = FSO.CreateTextFile(ArchREG, True)
        TE.WriteLine A1
        TE.WriteLine A2
        TE.WriteLine A3
        TE.WriteLine A4
        TE.WriteLine CStr(Errores)
    TE.Close
    MsgBox "Existe un archivo de registro de 3PM con datos no validos." + vbCrLf + _
    "Debe solicitar una licencia para este equipo. Solicitela a tbrSoft " + _
    "Argentina (info@tbrsoft.com / avazquez@cpcipc.org) una contraseña " + _
    "de acceso en esta PC como se indica en la página que sigue"
End Sub

Public Function GenerarCodigo(GUID As String) As String
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod)
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEmp2 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigo = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function

Public Function GenerarCodigoSL(GUID As String) As String
    'generacion de codigos SUPELICENCIA
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod) + 40
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEmp3 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigoSL = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function

Public Function GenerarCodigoDemo(GUID As String) As String
    'generacion de codigos que dan alguna validez minima adicional
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod) + 80
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEMP1 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigoDemo = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function



Public Function MostraDeA5(TXT As String)
    Dim c As Long, Letra As String, newTXT As String
    c = 0
    Do While c < Len(TXT)
        Letra = Mid(TXT, c + 1, 5)
        newTXT = newTXT + Letra
        c = c + 5
        If c < Len(TXT) Then newTXT = newTXT + "-"
    Loop
    MostraDeA5 = newTXT
End Function

Private Sub txtCOD_Change(Index As Integer)
    If Index < 7 And Len(txtCOD(Index)) = 5 Then
        'pasar a la casilla siguinete
        txtCOD(Index + 1).SetFocus
    End If
    If Index = 7 And Len(txtCOD(Index)) = 5 Then
        cmbCountry.SetFocus
    End If
End Sub

Private Sub txtCOD_GotFocus(Index As Integer)
    'pintar todo
    txtCOD(Index).SelStart = 0
    txtCOD(Index).SelLength = Len(txtCOD(Index))
End Sub
