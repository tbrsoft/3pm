VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmConfigLedsTeclado 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seteo leds de teclado"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkActionLedOn 
      BackColor       =   &H00000000&
      Caption         =   "Activar luces del teclado"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2430
      TabIndex        =   8
      Top             =   1020
      Width           =   3495
   End
   Begin VB.ComboBox cmbHasta 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmConfigLedsTeclado.frx":0000
      Left            =   2250
      List            =   "frmConfigLedsTeclado.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4470
      Width           =   880
   End
   Begin VB.ComboBox cmbDesde 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmConfigLedsTeclado.frx":0096
      Left            =   1080
      List            =   "frmConfigLedsTeclado.frx":00AF
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   4470
      Width           =   880
   End
   Begin VB.ComboBox cmbAction 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      ItemData        =   "frmConfigLedsTeclado.frx":012C
      Left            =   7380
      List            =   "frmConfigLedsTeclado.frx":0145
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1680
      Width           =   2055
   End
   Begin VB.ListBox lstSuccess 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   7245
   End
   Begin tbrFaroButton.fBoton fBoton1 
      Height          =   675
      Left            =   8220
      TabIndex        =   3
      Top             =   4020
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1191
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Grabar todo"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton fBoton4 
      Height          =   675
      Left            =   7050
      TabIndex        =   4
      Top             =   4020
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   1191
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir sin grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora de inicio y fin (24 representa final del dia) de estas funciones."
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
      Height          =   465
      Left            =   210
      TabIndex        =   7
      Top             =   3870
      Width           =   4035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConfigLedsTeclado.frx":01C2
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   765
      Left            =   210
      TabIndex        =   1
      Top             =   60
      Width           =   9255
   End
End
Attribute VB_Name = "frmConfigLedsTeclado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkActionLedOn_Click()
    lstSuccess.Enabled = chkActionLedOn
    
    Dim A As Long
    For A = 0 To cmbAction.UBound
        cmbAction(A).Enabled = lstSuccess.Enabled
    Next A
    
    cmbDesde.Enabled = lstSuccess.Enabled
    cmbHasta.Enabled = lstSuccess.Enabled
    
End Sub

Private Sub cmbAction_Click(Index As Integer)
    lstSuccess.ListIndex = Index
End Sub

Private Sub fBoton1_Click()

    ActionLedOn = CStr(Abs(chkActionLedOn))
    ChangeConfig "ActionLedOn", CStr(ActionLedOn)

    ActionLedMuchoCredito = cmbAction(0).ListIndex
    ActionLedPocoCredito = cmbAction(1).ListIndex
    ActionLedPalying = cmbAction(2).ListIndex
    ActionLedNoPlaying = cmbAction(3).ListIndex
    ActionLedPalyingVip = cmbAction(4).ListIndex
    ActionLedNoPlayVip = cmbAction(5).ListIndex
    
    ChangeConfig "ActionLedMuchoCredito", CStr(ActionLedMuchoCredito)
    ChangeConfig "ActionLedPocoCredito", CStr(ActionLedPocoCredito)
    ChangeConfig "ActionLedPalying", CStr(ActionLedPalying)
    ChangeConfig "ActionLedNoPlaying", CStr(ActionLedNoPlaying)
    ChangeConfig "ActionLedPalyingVip", CStr(ActionLedPalyingVip)
    ChangeConfig "ActionLedNoPlayVip", CStr(ActionLedNoPlayVip)
    
    ActionLedINIhs = cmbDesde.ListIndex
    ActionLedFINhs = cmbHasta.ListIndex
    
    ChangeConfig "ActionLedINIhs", CStr(ActionLedINIhs)
    ChangeConfig "ActionLedFINhs", CStr(ActionLedFINhs)
    
    
    Unload Me
End Sub

Private Sub fBoton4_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    
    chkActionLedOn.Value = ActionLedOn
    chkActionLedOn_Click 'que active o no segun corresponda
    
    'cargo los hotrarios
    Dim A As Long
    cmbDesde.Clear
    cmbHasta.Clear
    For A = 0 To 24
        cmbDesde.AddItem CStr(A)
        cmbHasta.AddItem CStr(A)
    Next A
    
    'cargar los sucesos y las acciones diusponibles
    lstSuccess.Clear
    cmbAction(0).Clear
    
    '//////////////////////////////////////////////////////////
    '//////////sucesos que se pueden controlar/////////////////
    '//////////////////////////////////////////////////////////
    lstSuccess.AddItem "Se paso el maximo de creditos permitidos"
    lstSuccess.AddItem "Hay menos del maximo de creditos permitidos"
    lstSuccess.AddItem "Se esta reproduciendo musica"
    lstSuccess.AddItem "Na hay reproduccion de música"
    lstSuccess.AddItem "Se esta reproduciendo una cancion VIP"
    lstSuccess.AddItem "No se esta reproduciendo canción VIP"
    '//////////////////////////////////////////////////////////
    ''//////////Acciones///////////////////////////////////////
    '//////////////////////////////////////////////////////////
    cmbAction(0).AddItem "No hacer nada"                     '0
    cmbAction(0).AddItem "'NUM' ON"                          '1
    cmbAction(0).AddItem "'NUM' OFF"                         '2
    cmbAction(0).AddItem "'CAPS' ON"                         '3
    cmbAction(0).AddItem "'CAPS' OFF"                        '4
    cmbAction(0).AddItem "'SCROLL' ON"                       '5
    cmbAction(0).AddItem "'SCROLL' OFF"                      '6
    '//////////////////////////////////////////////////////////
    
    'agregar todos los combos de accion necesarios
    For A = 0 To lstSuccess.ListCount - 1
        If A > 0 Then
            Load cmbAction(A)
            cmbAction(A).Top = cmbAction(A - 1).Top + cmbAction(A - 1).Height
            duplicateCombo cmbAction(0), cmbAction(A)
        Else 'es el primero
            cmbAction(A).Top = lstSuccess.Top + 30
        End If
        cmbAction(A).Left = lstSuccess.Left + lstSuccess.Width
        cmbAction(A).Visible = True
        cmbAction(A).ZOrder
    Next A
    
    
    'leer la configuracion y acomodar segu corresponde!
    cmbAction(0).ListIndex = ActionLedMuchoCredito
    cmbAction(1).ListIndex = ActionLedPocoCredito
    cmbAction(2).ListIndex = ActionLedPalying
    cmbAction(3).ListIndex = ActionLedNoPlaying
    cmbAction(4).ListIndex = ActionLedPalyingVip
    cmbAction(5).ListIndex = ActionLedNoPlayVip
    
    cmbDesde.ListIndex = ActionLedINIhs
    cmbHasta.ListIndex = ActionLedFINhs
    
End Sub

Private Sub duplicateCombo(cmbOrig As ComboBox, cmbDest As ComboBox)
    cmbDest.Clear
    If cmbOrig.ListCount = 0 Then Exit Sub
    
    Dim A As Long
    For A = 0 To cmbOrig.ListCount - 1
        cmbDest.AddItem cmbOrig.List(A)
    Next A
End Sub

Private Sub lstSuccess_Click()
    On Local Error Resume Next
    If lstSuccess.ListIndex = -1 Then Exit Sub
    cmbAction(lstSuccess.ListIndex).SetFocus
End Sub
