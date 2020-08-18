VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmERRORES 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "SI SUCEDEN ERRORES EN 3PM"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command1 
      Height          =   615
      Left            =   3930
      TabIndex        =   10
      Top             =   4920
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Ok"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton command2 
      Height          =   555
      Left            =   5550
      TabIndex        =   9
      Top             =   4260
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   979
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar descripción de falla e instantánea del sistema"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton command3 
      Height          =   555
      Left            =   450
      TabIndex        =   8
      Top             =   4260
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   979
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar detalle de error"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3330
      Width           =   5000
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   5220
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3330
      Width           =   5000
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   5220
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   870
      Width           =   5000
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   870
      Width           =   5000
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción del error:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   3060
      Width           =   2970
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En caso de que su sistema presente fallas notorias pero sin mensaje de error siga las siguinetes instrucciones:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   795
      Left            =   5220
      TabIndex        =   5
      Top             =   30
      Width           =   4950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "En caso de que su sistema cierre bruscamente por algún mensaje de error"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   765
      Left            =   150
      TabIndex        =   4
      Top             =   60
      Width           =   4950
   End
   Begin VB.Label lblREP 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción de la Falla:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5220
      TabIndex        =   3
      Top             =   3030
      Width           =   2970
   End
End
Attribute VB_Name = "frmERRORES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
    If Text3 = "" Then
        MsgBox TR.Trad("No hay texto para grabar!" + vbCrLf + _
            "Escriba el detalle de la falla y pruebe de nuevo%99%")
        Exit Sub
    End If
    
    Dim nFSO As New Scripting.FileSystemObject
    Dim nTE As TextStream
    
    Set nTE = fso.CreateTextFile(AP + "REG_DESCR_FALLA.W15", True)
        nTE.WriteLine "DESCRIPCION FALLA:"
        nTE.WriteLine Text3.Text
        nTE.WriteLine "CAMINO"
        nTE.WriteLine tERR.LogAcumulado
    nTE.Close
    
    Set nTE = Nothing
    Set nFSO = Nothing
    
    MsgBox TR.Trad("El detalle se grabo OK con formato W15. Envíe ahora " + _
        "los archivos mencionados%99%")
End Sub

Private Sub Command3_Click()
    
    If Text4 = "" Then
        MsgBox TR.Trad("No hay texto para grabar!" + vbCrLf + _
            "Escriba el mensaje de error que cierra 3PM bruscamente " + _
            "y pruebe de nuevo%99%")
        Exit Sub
    End If
    
    Dim nFSO As New Scripting.FileSystemObject
    Dim nTE As TextStream
    
    Set nTE = fso.CreateTextFile(AP + "REG_DESCR_ERR.W15", True)
        nTE.Write Text4.Text
    nTE.Close
    
    Set nTE = Nothing
    Set nFSO = Nothing
    
    MsgBox TR.Trad("El detalle se grabo OK con formato W15. Envie " + _
        "ahora los archivos mencionados%99%")
    
    'empaquetar todo para mandarlo por el servidor
    'son todos los w15
    'el reg3pm.log
    'el archivo de configuracion (que deberá incluir la versión)
    
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    MostrarCursor True
    
    Text1 = TR.Trad("  1- Ingrese a la configuración de 3PM%99%") + vbCrLf + _
        TR.Trad("  2- Ingrese a la seccion 'OTRAS OPCIONES'%99%") + vbCrLf + _
        TR.Trad("  3- Active la casilla 'ACTIVAR REGISTRO DE ERROR PERMANENETE'%99%") + vbCrLf + _
        TR.Trad("  4- Presione 'GRABAR'%99%") + vbCrLf + _
        TR.Trad("  5- Cierre 3PM y vuelva a ejecutarlo%99%") + vbCrLf + _
        TR.Trad("  6- Ponga al sistema en las mismas circunstancias en que se " + _
            "dio el error para tratar de que se genere de nuevo%99%") + vbCrLf + _
        TR.Trad("  7- Al aparecer el error y luego de que 3PM se cierre " + _
            "bruscamente escriba el texto de error en la casilla 'Descripcion " + _
            "del error' y luego presione el boton 'Grabar Detalle de error'%99%") + vbCrLf
    
    TR.SetVars "info@tbrsoft.com", "tbrsoft@cpcipc.org"
    Text1 = Text1 + TR.Trad("  8- Envie por email a tbrsoft (%01% - %02%)" + _
            " los siguientes archivos:" + vbCrLf + _
            "  De la carpeta de 3PM" + vbCrLf + _
            "   * Todos los archivos REG****.W15" + vbCrLf + _
            "   * El archivo reg3PM.log" + vbCrLf + _
            "  De la carpeta de sistema (C:\Windows\System en W98 o " + _
            "Me ó C:\Windows\System32 en WXP):" + vbCrLf + _
            "   * El archivo 3PM.CFG" + vbCrLf + vbCrLf + _
            "Con esta información tbrSoft le enviará en un lapso de " + _
            "24-72 Hs una respuesta concreta a su problema.%99%")

    Text2 = TR.Trad("  1- Ponga al sistema en las mismas circunstancias " + _
        "en que se presenta la falla%99%") + vbCrLf + _
        TR.Trad("  2- Describa la falla en la casilla 'Descripcion de la Falla' y " + _
        "luego presione el boton 'Grabar descripcion de la falla e " + _
        "Instantánea del sistema'%99%") + vbCrLf
    
    TR.SetVars "info@tbrsoft.com", "tbrsoft@cpcipc.org"
    Text2 = Text2 + TR.Trad("  3- Envie por email a tbrsoft (%01% - %02%) " + _
        "los siguientes archivos:" + vbCrLf + _
        "  De la carpeta de 3PM" + vbCrLf + _
        "   * Todos los archivos REG****.W15" + vbCrLf + _
        "   * El archivo reg3PM.log" + vbCrLf + _
        "  De la carpeta de sistema (C:\Windows\System en W98 o Me ó " + _
        "C:\Windows\System32 en WXP):" + vbCrLf + _
        "   * El archivo 3PM.CFG" + vbCrLf + _
        "Con esta información tbrSoft le enviará en un lapso de 24-72 Hs " + _
        "una respuesta concreta a su problema.%99%")
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command3.Caption = TR.Trad("Grabar Detalle de error%99%")
    Command2.Caption = TR.Trad("Grabar descripcion de la falla e Instantánea " + _
        "del sistema%99%")
    Command1.Caption = TR.Trad("OK%99%")
    Label3.Caption = TR.Trad("Descripcion del error:%99%")
    Label2.Caption = TR.Trad("En caso de que su sistema presente fallas notorias " + _
        "pero sin mensaje de error siga las siguientes instrucciones:%99%")
    Label1.Caption = TR.Trad("En caso de que su sistema cierre bruscamente " + _
        "por algún mensaje de error%99%")
    lblREP.Caption = TR.Trad("Descripción de la Falla:%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
