VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmREG2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "LICENCIA 3PM"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton XxBoton4 
      Height          =   405
      Left            =   9720
      TabIndex        =   5
      Top             =   4530
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton3 
      Height          =   405
      Left            =   960
      TabIndex        =   4
      Top             =   4380
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ver contrato de licencia"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
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
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   3990
      Value           =   1  'Checked
      Width           =   195
   End
   Begin tbrFaroButton.fBoton XxBoton2 
      Height          =   615
      Left            =   7950
      TabIndex        =   9
      Top             =   2550
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Insertar licencia recibida"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton1 
      Height          =   615
      Left            =   6090
      TabIndex        =   10
      Top             =   990
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Obtener archivo para pedir licencia"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   615
      Left            =   7890
      TabIndex        =   11
      Top             =   3450
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Validar Plug-ins Comprados"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   8
      fECol           =   5452834
   End
   Begin VB.Label lblNP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO DE LA LICENCIA"
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
      Height          =   255
      Left            =   4560
      TabIndex        =   13
      Top             =   4620
      Width           =   2685
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "He leido y estoy de acuerdo con el Contrato de Licencia"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   300
      TabIndex        =   12
      Top             =   3990
      Width           =   4275
   End
   Begin VB.Image Image1 
      Height          =   2820
      Left            =   360
      Picture         =   "frmREG2.frx":0000
      Top             =   90
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO DE LA LICENCIA"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Index           =   6
      Left            =   4530
      TabIndex        =   8
      Top             =   4170
      Width           =   6225
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Utilice otros complementos de 3PM."
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
      Height          =   525
      Index           =   5
      Left            =   4950
      TabIndex        =   7
      Top             =   3450
      Width           =   2445
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Más información en:"
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
      Height          =   885
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   2910
      Width           =   4365
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
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG2.frx":1B35
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
      Left            =   4830
      TabIndex        =   3
      Top             =   1650
      Width           =   5745
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   3
      X1              =   4485
      X2              =   11460
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   2
      X1              =   4470
      X2              =   4470
      Y1              =   7230
      Y2              =   -30
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      Index           =   0
      X1              =   4470
      X2              =   11010
      Y1              =   2370
      Y2              =   2370
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmREG2.frx":1BCA
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
      Height          =   825
      Index           =   3
      Left            =   4560
      TabIndex        =   1
      Top             =   90
      Width           =   6165
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Si ha recibido su archivo de licencia  desde tbrDataServer Cliente o desde tbrSoft cárguelo desde aquí."
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
      Height          =   885
      Index           =   2
      Left            =   4560
      TabIndex        =   0
      Top             =   2550
      Width           =   3195
   End
End
Attribute VB_Name = "frmREG2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub fBoton1_Click()
    frmPLIN.Show 1
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
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

    TR.SetVars "tbrSoft"
    Label1(4).Caption = TR.Trad("Más informacion en " + vbCrLf + _
                        "www.%01%.com/sw/3pm" + vbCrLf + _
                        "Email: info@%01%.com" + vbCrLf + _
                        " MSN: %01%@hotmail.com%99%")
    Label1(6).Caption = "|" + _
        CStr(K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw=="))) + "|" + _
        CStr(K.sabseee(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA=="))) + "|" + _
        CStr(K.sabseee(dcr("yTSbeYe2oWp2ydIUpGyes+DYNN6qU8l9pMMGAAqH+wBg8bBgTQ+/hw=="))) + "|" + _
        CStr(K.sabseee(dcr("VZPmSDtgWIj2UthiVZfN1LsFHe7IZv/K/ue9/JPXBYNJosAztaasKg=="))) + "|" + _
        CStr(K.sabseee(dcr("OqgcJfckN8975IVShi0xrqPphoO7CJfy1bRk3zQnHno="))) + "|" + _
        CStr(K.GETnnFecha) + "|" + CStr(K.GETnnFecha2) + "|" + _
        CStr(K.GETnnVers) + "|" + CStr(K.GETnnVers2) + "|" + _
        CStr(K.GETdifXaSoporte) + "|" + CStr(K.GETdifXaVersion) + "|" + K.sabseee_STR
        
        
    lblNP.Visible = (NP > 0)
    If NP > 0 Then lblNP.Caption = CStr(NP)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    MostrarCursor False
    frmIndex.Timer3.Enabled = True
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub XxBoton1_Click()
    
    Dim Ident3 As String
    
    Ident3 = InputBox(TR.Trad("Indique un breve recordatorio para esta PC" + vbCrLf + _
        "Por ejemplo 'rockola 17' o 'celeron266' o algun texto que le " + _
        "permita diferenciar este equipo%99%"), _
        TR.Trad("Identificacion basica del equipo a licenciar%99%"), _
        TR.Trad("Rockola 0001 (no use mas de 15 caracteres)%98%Aqui va " + _
        "un ejemplo de descripción que se puede hacer a la PC, es el nombre " + _
        "predeterminado con que se denomina a la PC. A partir de esto el usuario " + _
        "debera escribir otro%99%"))
    
    Ident3 = Left(Ident3, 15)
    Ident3 = Replace(Ident3, " ", "_")
    Ident3 = Replace(Ident3, "/", "_")
    Ident3 = Replace(Ident3, "\", "_")
    Ident3 = Replace(Ident3, "|", "_")
    Ident3 = Replace(Ident3, "?", "_")
    Ident3 = Replace(Ident3, "¿", "_")
    Ident3 = Replace(Ident3, "!", "_")
    Ident3 = Replace(Ident3, "¡", "_")
    Ident3 = Replace(Ident3, "+", "_")
    Ident3 = Replace(Ident3, "*", "_")
    Ident3 = Replace(Ident3, "#", "_")
    Ident3 = Replace(Ident3, "$", "_")
    Ident3 = Replace(Ident3, "%", "_")
    Ident3 = Replace(Ident3, "&", "_")
    Ident3 = Replace(Ident3, "'", "_")
    Ident3 = Replace(Ident3, Chr(34), "_")
    
    Dim CM As New CommonDialog
    
    CM.InitDir = ""
    
    'CM.DialogTitle = "Especifique en que carpeta se grabara"
    CM.DialogPrompt = TR.Trad("ESPECIFIQUE EN QUE DESTINO SE " + _
        "GRABARA%98%Se refiere a en que carpeta se grabara%99%")
    
    CM.ShowFolder
    Dim F As String
    
    F = CM.InitDir
    
    If F = "" Then Exit Sub
    If Right(F, 1) <> "\" Then F = F + "\"
    
    F = F + "CODIGO_3PM"
    If fso.FolderExists(F) = False Then fso.CreateFolder F
    
    Dim F2 As String
    F2 = F + "\3PM_" + Ident3 + CStr(Year(Date)) + CStr(Month(Date)) + "_" + CStr(Day(Date)) + ".LIC"
    
    If fso.FileExists(F2) Then fso.DeleteFile F2, True
    
    'cd4pm
    tSTR = dcr("Q2XVfrAUAdZINJMzhtkgNyKxMYyBMkFn")
    fso.CopyFile GPF(tSTR), F2, True
    
    TR.SetVars F2, "tbrSoft"
    MsgBox TR.Trad("El archivo para pedir su licencia se copio en" + vbCrLf + _
        "%01%" + vbCrLf + _
        "Envíelo por email a info@%02%.com o utilice el software especial " + _
        "de envio%98%La variable 1 es un path a un archivo%99%")
    
End Sub

Private Sub XxBoton2_Click()

    'leer algun archivo de licecnia
    Dim CM As New CommonDialog
    
    'Cargar licencia de 3PM v7 ...
    tSTR = dcr("c3Bw2mfLayocTYOm4pbjjEea0fne6co8JH97k8xj+e6XUfsdCf0biA38iJ95g4d+")
    CM.DialogTitle = tSTR
    
    'Licencia de 3PM v7
    CM.Filter = dcr("sJuE0dWMmhSb7A9XwmkXSRC9+jLN/o0+AucAxpf3viE=") + "(*.*)|*.*"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    'ponerlo como original ...
    fso.CopyFile F, GPF("cd7pm"), True
    ' y como copia ...
    fso.CopyFile F, GPF("cd8pm"), True
    
    tERR.Anotar "IC10"
    K.IngresaClave dcr("1Vx0YVGhEoIisHPLAZMHXw=="), True 'para que lea lo que acabo de insertar
    
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) = LicenciaVencida Then
        'La licencia es valida pero esta vencida, debe solicitar actualización a tbrsoft (info@tbrsoft.com)
        MsgBox dcr("Z2osmNjKl63bFTWavaa4Ogv7KDrjIRygFqba6tbGMjCr1S51EmvYqAucwSiflyqHf273384tQaL2lgdIeZbvhDQlitckwe/lswE6AYefOL0Klb9xB9SJDCwkTwRnX9pEqNfTfx05La6pGKjGNNFW3g==")
        Exit Sub
    End If
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) = BErronea Then
        'La licencia no es correcta
        MsgBox dcr("F7mIU1B2+aN6WS5VGH0O0kEQdjUKLasXkKAjoyHXYmQO2MuvHeSixw==")
        Exit Sub
    End If
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) = ParaOtraPC Then
        'El archivo es una licencia pero ha sido desarrollada para otro equipo.
        'Es posible tambien que esto suceda por cambios que haya realizado en el hardware de su pc.
        'Consulte a tbrSoft informando los cambios de hardware si los hubo.
        'Puede enviar el arhivo reg3PM.log que esta en la carpeta de 3PM para recibir una pronta respuesta%99%
        tSTR = dcr("0C0d4AcUabUKptJC1TQhkexCsocfahhfbaQTRdy9yU5uMatnxakbAP0CE1Ct1n2+GU4ClWA2VOq0MclzAbVmCllKur7z3tVOd3t+eqFOMCyd8EBLv31w9VZITjIbcrZX46NWT4CUc6qN/O6UpYnMI68JLqYg+B7pO8Nho7DgddIv2zEAEGKfvvEDqBWFxNXdsxo+c2dFc0gm2SpAKE0NLGSsA1aDljLnXA3P/Fde/WozvPuOg55J4zpSaE5TVel3Q3kcffywQmXw9//vEol7YW4Z8rIoYXFJOjDpiMHrteJFMuDEbl4OpQmzm0SUA+f7Kv8jNNWcm9GhKIgr8cVFM/nvZUp9wBnCDUxG6rmtmi71eTY5RL+/2EKVLTxPxC8fyyMdk/AJPT52yWKiTQDuETGzILZggLoagiT9z+Q297dKiY6R8A+gE4ZBoYJwgc+dbdVjQo0N9nItTn7Yn0ATzQ==")
        MsgBox tSTR
        Exit Sub
    End If
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) = CGratuita Then
        'El archivo es una licencia gratuita%99%
        tSTR = dcr("SilvSmiUyJAHkszjvdo8ZoVojF0yASverX6C1iu+1E4LEp+bIo6rgIAdbTjcw5yKtF1QwFbChpU=")
        MsgBox tSTR
    End If
        
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) = DMinima Then
        'El archivo es una licencia minima%99%
        tSTR = dcr("ScAMW/vjNnMlGklUCrX71qcnTZ8My8+E7vuO43xUr0ppvCNCRRirXtEoS7sYNQJa9XI/OWrEKNg=")
        MsgBox tSTR
    End If
    
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) = EComun Then
        'El archivo es una licencia simple
        MsgBox dcr("N2LoA4ffuFsCAZwqJAU8kkrc6flqTAnTSzxB/92JypOxLu5RxHH7rEmqpdk5c540")
    End If
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) = FPremium Then
        'El archivo es una licencia premium%99%
        tSTR = dcr("DeVxGdX8ADYFjdibvXI9s++enRlG+cHAgj7UxeI96xfoAoJemDD5dGL7Y5EV1uNV3X4jsnEsyvo=")
        MsgBox tSTR
    End If
    
    If K.sabseee(dcr("q44KmdDBQ+IB8dTOX8F+VA==")) = GFull Then
        'El archivo es una licencia full
        MsgBox dcr("jOG0byK7hBLqiii+xUDlC6AFEm9deuveZZEvd4TpsP4aFkXEWZudZOv38gug2r4h")
    End If
    
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) = Supsabseee Then
        'El archivo es una SuperLicencia
        MsgBox dcr("HEadsdpypfdzGht8+ZzCB5Bw62EU/rhuFYigl6LUVy+pSWOwtbZ9ZPGPPVRUsNx6")
    End If
    
    If K.sabseee(dcr("1Vx0YVGhEoIisHPLAZMHXw==")) > Supsabseee Then
        'libertad al 3PM
        MsgBox dcr("BE1s+L/UQcVFZoCB2K5uK1YXnVpNO+7xC0MVlRoELwVkAUAq+Gby/Q==")
    End If
    
    '3PM se cerrará ahora. Al iniciarlo nuevamente se usará su archivo de licencia
    MsgBox dcr("n2A5NoSLR7NWfc3eas//L7FI1Yv3bMD9oGGWjqxs9jobUAeZQM5D5UoG+xDd5H0tLV1Vs9sxX8JyWXC0hd/8DhWKwJHZUWBfB2CPik1dB0kcwqhMvibBBUL/kBmojBEg")
    
    Unload Me
    YaCerrar3PM True
    
End Sub

Private Sub XxBoton3_Click()
    AbrirArchivo AP + "license.rtf", Me
    'frmCLUF.Show 1
End Sub

Private Sub XxBoton4_Click()
    Unload Me
End Sub

'eliminar licencia no sirve, ayuda a los hackers
'Private Sub XxBoton5_Click()
'    If MsgBox(TR.T r a d("¿Desea borrar los datos de su licencia actual para " + _
'        "volver a cargarlos?" + vbCrLf + _
'        "Usese solo para cuando obtenga una nueva clave para cargar%99%"), vbCritical + vbYesNo, "NUEVA LICENCIA") = vbNo Then Exit Sub
'
'    'borro el archivo de registro para que inicie preguntando clave
'
'    'borrar el original...
'    If fso.FileExists(GPF("cd7pm")) Then fso.DeleteFile GPF("cd7pm"), True
'    '... y la copia
'    If fso.FileExists(GPF("cd8pm")) Then fso.DeleteFile GPF("cd8pm"), True
'
'    If fso.FileExists(GPF("cd7pm")) Or fso.FileExists(GPF("cd8pm")) Then
'        MsgBox TR.Tr a d("No se ha podido borrar el archivo de licencia%99%")
'    Else
'        MsgBox TR.Tr a d("La información de licencia se ha borrado correctamente. " + _
'            "El sistema se cerrará para que cargue nuevamente su clave%99%")
'    End If
'
'    End
'End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    XxBoton4.Caption = TR.Trad("Salir%99%")
    XxBoton3.Caption = TR.Trad("Ver contrato de licencia%99%")
    XxBoton2.Caption = TR.Trad("Insertar licencia recibida%99%")
    XxBoton1.Caption = TR.Trad("Obtener archivo para pedir licencia%99%")
    Check1.Caption = TR.Trad("He leido y estoy de acuerdo con el Contrato de Licencia%99%")
    'XxBoton5.Caption = TR.Trad("Quitar licencia actual%99%")
    fBoton1.Caption = TR.Trad("Validar Plugins Comprados%99%")
    
    Label1(5).Caption = TR.Trad("Utilice otros complementos de 3PM%99%")
    TR.SetVars "tbrSoft"
    Label1(4).Caption = TR.Trad("Más informacion en www.%01%.com/sw/3pm o por email " + _
        "a info@%01%.com o por msn a %01%@hotmail.com%99%")
        
    Label1(0).Caption = TR.Trad("Si desea asistencia y habilitaciones de modo " + _
        "automático on-line consultenos sobre tbrDataServer que le permite mantener " + _
        "contacto directo.%99%")
    
    TR.SetVars dcr("1Vx0YVGhEoIisHPLAZMHXw=="), "tbrSoft"
    Label1(3).Caption = TR.Trad("Si desea obtener una licencia de %01% v7 para este " + _
        "equipo debe enviar un archivo a %02% con la solicitud y datos únicos de " + _
        "esta PC. Este trámite debe realizarse incluso si tenía una " + _
        "licencia de %01% v6%99%")
    Label1(2).Caption = TR.Trad("Si ya ha recibido su archivo de licencia " + _
        "carguelo desde aquí%99%")
End Sub
