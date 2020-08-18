VERSION 5.00
Object = "{AC1ACB77-BE60-49F4-BE38-2F9A87F5E5E4}#2.0#0"; "tbrX_Boton II.ocx"
Begin VB.Form frmREG2 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LICENCIA 3PM"
   ClientHeight    =   5820
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrX_Boton2.XxBoton XxBoton4 
      Height          =   375
      Left            =   900
      TabIndex        =   9
      Top             =   4020
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      xFColor         =   16777215
      xBColor         =   192
      xCapt           =   "Salir"
      xEnabled        =   -1  'True
   End
   Begin tbrX_Boton2.XxBoton XxBoton3 
      Height          =   405
      Left            =   630
      TabIndex        =   8
      Top             =   4830
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   714
      xFColor         =   16777215
      xBColor         =   4210816
      xCapt           =   "Ver contrato de licencia"
      xEnabled        =   -1  'True
   End
   Begin tbrX_Boton2.XxBoton XxBoton2 
      Height          =   645
      Left            =   5490
      TabIndex        =   7
      Top             =   3330
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1138
      xFColor         =   16777215
      xBColor         =   64
      xCapt           =   "Insertar licencia recibida"
      xEnabled        =   -1  'True
   End
   Begin tbrX_Boton2.XxBoton XxBoton1 
      Height          =   525
      Left            =   5190
      TabIndex        =   6
      Top             =   1110
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   926
      xFColor         =   16777215
      xBColor         =   64
      xCapt           =   "Obtener archivo para pedir licencia"
      xEnabled        =   -1  'True
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia"
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
      Height          =   465
      Left            =   540
      TabIndex        =   3
      Top             =   5220
      Value           =   1  'Checked
      Width           =   2925
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2820
      Left            =   30
      Picture         =   "frmREG2.frx":0000
      ScaleHeight     =   2820
      ScaleWidth      =   3750
      TabIndex        =   0
      Top             =   60
      Width           =   3750
   End
   Begin tbrX_Boton2.XxBoton XxBoton5 
      Height          =   585
      Left            =   5730
      TabIndex        =   11
      Top             =   4980
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1032
      xFColor         =   16777215
      xBColor         =   64
      xCapt           =   "Quitar licencia actual"
      xEnabled        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Más informacion en www.tbrsoft.com/sw/3pm o por email a info@tbrsoft.com o por msn a tbrsoft@hotmail.com"
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
      Height          =   765
      Index           =   4
      Left            =   240
      TabIndex        =   10
      Top             =   2850
      Width           =   3315
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   1
      X1              =   0
      X2              =   6540
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Si debe cambiar el tipo de licencia que tiene use esta opción"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   4050
      TabIndex        =   5
      Top             =   4680
      Width           =   6105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG2.frx":645F
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
      Index           =   0
      Left            =   4200
      TabIndex        =   4
      Top             =   2040
      Width           =   5745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   3840
      X2              =   10800
      Y1              =   4350
      Y2              =   4350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   3840
      X2              =   3840
      Y1              =   5820
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   3840
      X2              =   10380
      Y1              =   2910
      Y2              =   2910
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG2.frx":64ED
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Index           =   3
      Left            =   3930
      TabIndex        =   2
      Top             =   90
      Width           =   6165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Si ya ha recibido su archivo de licencia carguelo desde aquí"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   4020
      TabIndex        =   1
      Top             =   3030
      Width           =   6105
   End
End
Attribute VB_Name = "frmREG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    MostrarCursor True
    
'    Dim txt As String
'    txt = "Bienvenido a 3PM." + vbCrLf + "Gracias por confiar en tbrSoft Argentina" + vbCrLf + vbCrLf + _
'    "Puede utilizar esta version demo con algunas restricciones " + vbCrLf + vbCrLf + _
'    "Si desea adquirir definitivamente este software presione el boton " + _
'    "'COMPRAR AHORA' o siga los pasos indicados " + _
'    "en la herramienta creada para este fin en Inicio/Programas/tbrSoft/3PM/Licencia" + vbCrLf + vbCrLf + _
'    "Si desea quitar esta pantalla de bienvenida y otras limitaciones " + _
'    "puede obtener una clave gratuita utilizando la misma herramienta de compra" + vbCrLf + vbCrLf + _
'    "Si esta PC ya contaba con licencia de 3PM la funcion de 'COMPRAR LICENCIA' " + _
'    "lo resolvera." + vbCrLf + vbCrLf + _
'    "Si ya ha adquirido y dispone de su archivo de licencia use la opción" + _
'    "'Cargar archivo de licencia'" + vbCrLf + vbCrLf + _
'    "Cualquier duda envie un email a info@tbrsoft.com"

    Label1(4).Caption = "Más informacion en " + vbCrLf + _
                        "www.tbrsoft.com/sw/3pm" + vbCrLf + _
                        "Email: info@tbrsoft.com" + vbCrLf + _
                        " MSN: tbrsoft@hotmail.com"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MostrarCursor False
End Sub

Private Sub XxBoton1_Click()
    Dim CM As New CommonDialog
    
    CM.InitDir = ""
    
    'CM.DialogTitle = "Especifique en que carpeta se grabara"
    CM.DialogPrompt = "ESPECIFIQUE EN QUE DESTINO SE GRABARA"
    
    CM.ShowFolder
    Dim F As String
    
    F = CM.InitDir
    
    If F = "" Then Exit Sub
    If Right(F, 1) <> "\" Then F = F + "\"
    
    F = F + "CODIGO_3PM"
    If FSO.FolderExists(F) = False Then FSO.CreateFolder F
    
    Dim F2 As String
    F2 = F + "\3PM_" + CStr(Day(Date)) + "_" + CStr(Month(Date)) + "_" + CStr(Hour(time)) + "_" + CStr(Minute(time)) + ".LIC"
    
    If FSO.FileExists(F2) Then FSO.DeleteFile F2, True
    
    FSO.CopyFile GPF("cd4pm"), F2, True
    
    MsgBox "El archivo para pedir su licencia se copio en " + vbCrLf + F2 + vbCrLf + _
        "Envíelo por email a info@tbrsoft.com o utilize el software especial de envio"
    
End Sub

Private Sub XxBoton2_Click()

    'leer algun archivo de licecnia
    Dim CM As New CommonDialog
    CM.DialogTitle = "Cargar licencia de 3PM v7 ..."
    CM.Filter = "Licencia de 3PM v7 (*.*)|*.*"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'ponerlo como original ...
    FSO.CopyFile F, GPF("cd7pm"), True
    ' y como copia ...
    FSO.CopyFile F, GPF("cd8pm"), True
    
    K.IngresaClave
    
    If K.LICENCIA = BErronea Then
        MsgBox "La licencia no es correcta"
        Exit Sub
    End If
    
    If K.LICENCIA = ParaOtraPC Then
        MsgBox "El archivo es una licencia pero ha sido desarrollada para otro equipo." + vbCrLf + _
            "Es posible tambien que esto suceda por cambios que haya realizado en el hardware de su pc." + vbCrLf + _
            "Consulte a tbrSoft informando los cambios de hardware si los hubo"
        Exit Sub
    End If
    
    If K.LICENCIA = CGratuita Then
        MsgBox "El archivo es una licencia gratuita"
    End If
        
    If K.LICENCIA = DMinima Then
        MsgBox "El archivo es una licencia minima"
    End If
    
    If K.LICENCIA = EComun Then
        MsgBox "El archivo es una licencia simple"
    End If
    
    If K.LICENCIA = FPremium Then
        MsgBox "El archivo es una licencia premium"
    End If
    
    If K.LICENCIA = GFull Then
        MsgBox "El archivo es una licencia full"
    End If
    
    If K.LICENCIA = HSuperLicencia Then
        MsgBox "El archivo es una SuperLicencia"
    End If
    
    MsgBox "3PM se cerrará ahora. Al iniciarlo nuevamente se usara su archivo de licencia"
    
    Unload Me
    
    YaCerrar3PM True
    
End Sub

Private Sub XxBoton3_Click()
    frmCLUF.Show 1
End Sub

Private Sub XxBoton4_Click()
    Unload Me
End Sub

Private Sub XxBoton5_Click()
    If MsgBox("¿Desea borrar los datos de su licencia actual para volver a cargarlos?" + vbCrLf + _
        "Usese solo para cuando obtenga una nueva clave para cargar", vbCritical + vbYesNo, "NUEVA LICENCIA") = vbNo Then Exit Sub
    
    'borro el archivo de registro para que inicie preguntando clave
    
    'borrar el original...
    If FSO.FileExists(GPF("cd7pm")) Then FSO.DeleteFile GPF("cd7pm"), True
    '... y la copia
    If FSO.FileExists(GPF("cd8pm")) Then FSO.DeleteFile GPF("cd8pm"), True
    
    If FSO.FileExists(GPF("cd7pm")) Or FSO.FileExists(GPF("cd8pm")) Then
        MsgBox "No se ha podido borrar el archivo de licencia"
    Else
        MsgBox "La información de licencia se ha borrado correctamente. El sistema se cerrará " + _
            "para que cargue nuevamente su clave"
    End If
    
    End
End Sub
