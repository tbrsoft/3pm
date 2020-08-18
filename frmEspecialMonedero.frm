VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmEspecialMonedero 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Correcion señal monedero"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin tbrFaroButton.fBoton Command2 
      Height          =   585
      Left            =   8340
      TabIndex        =   14
      Top             =   5850
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   585
      Left            =   8340
      TabIndex        =   13
      Top             =   5220
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar y salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command3 
      Height          =   435
      Left            =   6210
      TabIndex        =   12
      Top             =   5580
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "-"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command4 
      Height          =   435
      Left            =   5730
      TabIndex        =   11
      Top             =   5580
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "+"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command5 
      Height          =   435
      Left            =   3780
      TabIndex        =   10
      Top             =   5580
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "-"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command6 
      Height          =   435
      Left            =   3300
      TabIndex        =   9
      Top             =   5580
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "+"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.TextBox txtMS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   5790
      TabIndex        =   8
      Text            =   "300"
      Top             =   6240
      Width           =   705
   End
   Begin VB.TextBox txtMS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3450
      TabIndex        =   7
      Text            =   "300"
      Top             =   6240
      Width           =   705
   End
   Begin VB.ListBox lstVals 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5205
      Index           =   1
      IntegralHeight  =   0   'False
      ItemData        =   "frmEspecialMonedero.frx":0000
      Left            =   5280
      List            =   "frmEspecialMonedero.frx":0007
      TabIndex        =   2
      Top             =   330
      Width           =   2175
   End
   Begin VB.ListBox lstVals 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5220
      Index           =   0
      IntegralHeight  =   0   'False
      ItemData        =   "frmEspecialMonedero.frx":0014
      Left            =   2700
      List            =   "frmEspecialMonedero.frx":001B
      TabIndex        =   0
      Top             =   330
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   3060
      X2              =   6000
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Máximo de milisegundos de separación para considerarlas separadas. No es recomendable utilizar más de 400 ms."
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
      Index           =   4
      Left            =   90
      TabIndex        =   6
      Top             =   5700
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Elija el valor que necesite de la lista correspondiente, con los botones ""+"" y ""-"" modifique hasta el valor necesario."
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
      Height          =   1545
      Index           =   3
      Left            =   8160
      TabIndex        =   5
      Top             =   1080
      Width           =   2235
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla monedero 2 (S)"
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
      Height          =   285
      Index           =   2
      Left            =   5310
      TabIndex        =   4
      Top             =   30
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla monedero 1 (Q)"
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
      Height          =   285
      Index           =   1
      Left            =   2730
      TabIndex        =   3
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmEspecialMonedero.frx":0028
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
      Height          =   5295
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   270
      Width           =   2475
   End
End
Attribute VB_Name = "frmEspecialMonedero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim te7 As TextStream, J As Long
    'lista de reemplazos
    Set te7 = fso.OpenTextFile(GPF("rempmon45"), ForWriting, True)
        te7.WriteLine "TO Q"
        For J = 1 To 20
            te7.WriteLine lstVals(0).List(J - 1)
        Next J
        te7.WriteLine "TO S"
        For J = 1 To 20
            te7.WriteLine lstVals(1).List(J - 1)
        Next J
        te7.WriteLine txtMS(0)
        te7.WriteLine txtMS(1)
    te7.Close
    
    Set te7 = Nothing
    
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    ModifiVal 1, lstVals(1).ListIndex, -1
End Sub

Private Sub Command4_Click()
    ModifiVal 1, lstVals(1).ListIndex, 1
End Sub

Private Sub Command5_Click()
    ModifiVal 0, lstVals(0).ListIndex, -1
End Sub

Private Sub Command6_Click()
    ModifiVal 0, lstVals(0).ListIndex, 1
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir 'Agregado por el complemento traductor
    lstVals(0).Clear: lstVals(1).Clear
    Dim J As Long
    For J = 1 To 20
        lstVals(0).AddItem CStr(J) + "=0"
        lstVals(1).AddItem CStr(J) + "=0"
    Next J
    
    'ver si ya existe y mostrarlo como esta
    Dim TMP As String, SP() As String
    Dim TE8 As TextStream
    If fso.FileExists(GPF("rempmon45")) Then
        Set TE8 = fso.OpenTextFile(GPF("rempmon45"), ForReading, False)
            TMP = TE8.ReadLine 'solo dice "to Q"
            For J = 1 To 20
                TMP = TE8.ReadLine
                lstVals(0).List(J - 1) = TMP
            Next J
            TMP = TE8.ReadLine 'solo dice "to S"
            For J = 1 To 20
                TMP = TE8.ReadLine
                lstVals(1).List(J - 1) = TMP
            Next J
            txtMS(0) = TE8.ReadLine
            txtMS(1) = TE8.ReadLine
        TE8.Close
    End If
    
    Set TE8 = Nothing
    With frmConfig
        lstVals(0).Enabled = (.chkCS.Value = 1)
        lstVals(1).Enabled = (.chkCS.Value = 1)
        command3.Enabled = (.chkCS.Value = 1)
        Command4.Enabled = (.chkCS.Value = 1)
        Command5.Enabled = (.chkCS.Value = 1)
        Command6.Enabled = (.chkCS.Value = 1)
        txtMS(0).Enabled = (.chkCS.Value = 1)
        txtMS(1).Enabled = (.chkCS.Value = 1)
    End With
End Sub

Private Sub ModifiVal(iLST As Long, lstIndex As Long, Var As Long)

    If lstVals(iLST).ListIndex = -1 Then Exit Sub
    
    Dim SP() As String
    SP = Split(lstVals(iLST), "=")
    
    Dim Cant As Long 'valor actual
    Cant = CLng(SP(1))
    
    Cant = Cant + Var
    If Cant < 0 Then Cant = 0
    
    lstVals(iLST).List(lstIndex) = SP(0) + "=" + CStr(Cant)
    
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command2.Caption = TR.Trad("Salir%99%")
    Command1.Caption = TR.Trad("Grabar y salir%99%")
    Label1(4).Caption = TR.Trad("Maximo de milisegundos de espacio para " + _
        "considerarlas separadas. Mas de 400 no se recomienda%99%")
    Label1(3).Caption = TR.Trad("Eliga el valor que necesite de la lista " + _
        "correspondiente y con los botones  +  y  -  modifique hasta " + _
        "el valor necesario%99%")
    Label1(2).Caption = TR.Trad("Tecla monedero 2 (S)%99%")
    Label1(1).Caption = TR.Trad("Tecla monedero 1 (Q)%99%")
    Label1(0).Caption = TR.Trad("La listas muestran los valores de señales " + _
        "que pueden ingresar en un breve lapso de tiempo. Algo 'no humano' " + _
        "que llega desde un monedero electrónico. Esto es de mucha utilidad " + _
        "si tiene problemas con su adpatador conectado desde el monedero a " + _
        "su teclado. Por ejemplo si debe recibir 5 señales en 500 milisegundos " + _
        "y llegan a veces 3 o 4 puede configurarlo para que cuando lleguen 3 o " + _
        "4 señales muy juntas interpretarlas como 5 señales. Lo mejor sería " + _
        "solucionar el problema de su adaptador, mientras tanto esta función " + _
        "es de utilidad%99%")
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub
