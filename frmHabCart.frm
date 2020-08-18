VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmHabCart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00400000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Habilitar en esta PC venta de música"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   405
      Left            =   1530
      TabIndex        =   0
      Top             =   2130
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "SALIR"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   645
      Left            =   360
      TabIndex        =   1
      Top             =   780
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Generar archivo para pedir habilitación"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton2 
      Height          =   645
      Left            =   360
      TabIndex        =   2
      Top             =   1440
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1138
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Insertar archivo recibido para habilitar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxXXX x x X x x xx x Xx x X xxx X X xx x x x x X X X X  XXXXXXX X X X X X XX X X X  x X x XX xx X x xX  XX X x xX x"
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
      Height          =   645
      Left            =   210
      TabIndex        =   3
      Top             =   30
      Width           =   3885
   End
End
Attribute VB_Name = "frmHabCart"
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
                   CStr(Year(Date)) + CStr(Month(Date)) + "_" + CStr(Day(Date)) + ".LIC_Carrito"
    
    If fso.FileExists(F2) Then fso.DeleteFile F2, True
    
    'crear uno para
    Dim nFOt2 As New tbrDATA.clsTODO
    'asegurarse que vaya con el noombre que tiene que ir!!!
    nFOt2.SetLog AP + "kc3.log"
    nFOt2.SetSF dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==") 'nuevo agosto 2007 para no mezclar con karaokes ni con programas de artime y manu
    nFOt2.DoNow F2
    
    TR.SetVars F2
    MsgBox TR.Trad("El archivo para habilitar uso de CD se copio en " + vbCrLf + _
        "%01%" + vbCrLf + _
        "Envíelo por email a info@tbrsoft.com o utilize el software especial de envio%99%")
End Sub

Private Sub fBoton2_Click()
    'leer algun archivo de licecnia
    Dim CM As New CommonDialog
    CM.DialogTitle = TR.Trad("Cargar licencia de Carro de compras ...%99%")
    CM.Filter = TR.Trad("Licencia de Carro de compras%99%") + " (*.*)|*.*"
    CM.ShowOpen
    
    Dim F As String
    F = CM.FileName
    
    If F = "" Then Exit Sub
    
    tERR.Anotar "IC10kar1"
    
    Dim PARA As String
    PARA = dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")
    
    fso.CopyFile F, GPF("plin1"), True
    fso.CopyFile F, GPF("plin2"), True
    
    tERR.Anotar "IC10kar2"
    K.IngresaClave PARA, True
    
    If K.sabseee(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA==")) >= GFull Then
        MsgBox TR.Trad("Se cargo la licencia del carro de ventas sin problemas%99%")
    Else
        MsgBox TR.Trad("No se cargo la licencia contacte a tbrSoft%99%")
    End If
End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    
    Dim RDS As TypeLic
    RDS = K.sabseee(dcr("MCuVh38359iRH+GBaAkXedz8Pl38peUqZHKs0a0SpMe+QLrW9mKdnA=="))
    If RDS <= aSinCargar Then Label1.Caption = "Sin licencia para venta de música cargada"
    If RDS = BErronea Then Label1.Caption = "Licencia para venta de música errónea o no válida"
    If RDS = CGratuita Then Label1.Caption = "Licencia gratuita para venta de música cargada"
    If RDS > CGratuita And RDS < Supsabseee Then Label1.Caption = "Licencia simple para venta de música cargada"
    If RDS = Supsabseee Then Label1.Caption = "SuperLicencia para venta de música cargada"
    
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
