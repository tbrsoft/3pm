VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmHabKar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Habilitar en esta PC uso de karaokes"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   405
      Left            =   1650
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   645
      Left            =   930
      TabIndex        =   2
      Top             =   780
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Obtener archivo para pedir licencia de karaoke."
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   645
      Left            =   930
      TabIndex        =   3
      Top             =   1620
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Insertar licencia de karaoke recibida."
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3000
      Width           =   225
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Genere el pedido de licencia de karaoke e ingrese la licencia desde aquí."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   465
      Index           =   2
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   4515
   End
End
Attribute VB_Name = "frmHabKar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fBoton1_Click()
        
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
    CM.DialogPrompt = TR.Trad("ESPECIFIQUE EN QUE DESTINO SE GRABARA%98%Se abrira un " + _
        "cuadro de dialogo para elegir una carpeta%99%")
    
    CM.ShowFolder
    Dim F As String
    
    F = CM.InitDir
    
    If F = "" Then Exit Sub
    If Right(F, 1) <> "\" Then F = F + "\"
    
    F = F + "CODIGO_3PM"
    If fso.FolderExists(F) = False Then fso.CreateFolder F
    
    Dim F2 As String
    F2 = F + "\" + Left(lstCDs, 5) + "_" + _
                   Ident3 + "_" + _
                   CStr(Year(Date)) + CStr(Month(Date)) + "_" + CStr(Day(Date)) + ".LIC_Kar"
    
    If fso.FileExists(F2) Then fso.DeleteFile F2, True
    
    'crear uno para
    Dim nFOt2 As New tbrDATA.clsTODO
    'asegurarse que vaya con el noombre que tiene que ir!!!
    nFOt2.SetLog AP + "kc2.log"
    nFOt2.SetSF "mLicenciaCD001Kar" 'nuevo agosto 2007 para no mezclar con karaokes ni con programas de artime y manu
    nFOt2.DoNow F2
    
    TR.SetVars F2
    MsgBox TR.Trad("El archivo para habilitar uso de CD se copio en " + vbCrLf + _
        "%01%" + vbCrLf + _
        "Envíelo por email a info@tbrsoft.com o utilize el software especial de envio%99%")
    
End Sub

Private Sub fBoton2_Click()
    'leer algun archivo de licecnia
    Dim CM As New CommonDialog
    CM.DialogTitle = TR.Trad("Cargar licencia de KARAOKE ...%99%")
    TR.SetVars "KARAOKE 3PM"
    CM.Filter = TR.Trad("Licencia de %01%%98%La variable dice Karaoke 3PM%99%") + " (*.*)|*.*"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    tERR.Anotar "IC10kar1"
    
    Dim PARA As String
    PARA = "mLicenciaCD001Kar"
    
    'YYYYYYYYYYYYYYYYYYYYYY
    'traigo la licencia del 2 y anda en el 1 tambien!!!!!!!!!!!!!!
    fso.CopyFile F, GPF("plin7"), True
    fso.CopyFile F, GPF("plin8"), True 'copia de seguridad
    
    tERR.Anotar "IC10kar2"
    K.IngresaClave PARA, True
    
    Dim j As TypeLic
    j = K.sabseee("mLicenciaCD001Kar")
    If j = Supsabseee Then
        MsgBox TR.Trad("Se cargo la SuperLicencia de karaoke%99%")
    Else
        If j >= EComun Then
            MsgBox TR.Trad("Se cargo la licencia de karaoke%99%")
        Else
            MsgBox TR.Trad("No se cargo la licencia contacte a tbrSoft%99%")
        End If
    End If
End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub

'cada CD de karaokes es unico y tiene su propia clave
'clave CD 01: "sadjf98sad7f980asd7f098asdfasdf87sad809f7as0d9f"
'clave CD 02: "asdf8097sad7f6sa543f54sad3f54sad3f4sadfdsasadfs6a5d"
'clave CD 03: "sdf6asd7f65sad65f4sad7f4as8df598sadf87sad6f987sad6f9"
Private Sub Form_Load()
    Pintar_fBoton Me
    Label2.Caption = K.sabseee("mLicenciaCD001Kar")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
