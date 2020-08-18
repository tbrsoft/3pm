VERSION 5.00
Begin VB.Form frmEspecialMonedero 
   BackColor       =   &H00404040&
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
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtMS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   5760
      TabIndex        =   14
      Text            =   "300"
      Top             =   6210
      Width           =   705
   End
   Begin VB.TextBox txtMS 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   3420
      TabIndex        =   13
      Text            =   "300"
      Top             =   6210
      Width           =   705
   End
   Begin VB.CommandButton Command6 
      Caption         =   "+"
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
      Height          =   405
      Left            =   3390
      TabIndex        =   11
      Top             =   5550
      Width           =   465
   End
   Begin VB.CommandButton Command5 
      Caption         =   "-"
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
      Height          =   405
      Left            =   3870
      TabIndex        =   10
      Top             =   5550
      Width           =   465
   End
   Begin VB.CommandButton Command4 
      Caption         =   "+"
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
      Height          =   405
      Left            =   5610
      TabIndex        =   9
      Top             =   5550
      Width           =   465
   End
   Begin VB.CommandButton Command3 
      Caption         =   "-"
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
      Height          =   405
      Left            =   6090
      TabIndex        =   8
      Top             =   5550
      Width           =   465
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8340
      TabIndex        =   7
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar y salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   8340
      TabIndex        =   6
      Top             =   5490
      Width           =   1365
   End
   Begin VB.ListBox lstVals 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
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
      Left            =   5070
      List            =   "frmEspecialMonedero.frx":0007
      TabIndex        =   2
      Top             =   300
      Width           =   2175
   End
   Begin VB.ListBox lstVals 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
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
      Top             =   300
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
      Caption         =   "Maximo de milisegundos de separacion para considerarlas separadas. Mas de 400 no se recomienda"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1035
      Index           =   4
      Left            =   90
      TabIndex        =   12
      Top             =   5700
      Width           =   2985
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Eliga el valor que necesite de la lista correspondiente y con los botones ""+"" y ""-"" modifique hasta el valor necesario"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1545
      Index           =   3
      Left            =   7530
      TabIndex        =   5
      Top             =   930
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla monedero 2 (S)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   5100
      TabIndex        =   4
      Top             =   60
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tecla monedero 1 (Q)"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
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
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
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
    Set te7 = FSO.OpenTextFile(GPF("rempmon45"), ForWriting, True)
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
    lstVals(0).Clear: lstVals(1).Clear
    Dim J As Long
    For J = 1 To 20
        lstVals(0).AddItem CStr(J) + "=0"
        lstVals(1).AddItem CStr(J) + "=0"
    Next J
    
    'ver si ya existe y mostrarlo como esta
    Dim TMP As String, SP() As String
    Dim TE8 As TextStream
    If FSO.FileExists(GPF("rempmon45")) Then
        Set TE8 = FSO.OpenTextFile(GPF("rempmon45"), ForReading, False)
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
        Command3.Enabled = (.chkCS.Value = 1)
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
