VERSION 5.00
Begin VB.Form frmMantenimiento 
   BackColor       =   &H00404080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de 3PM"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404080&
      Caption         =   "Ver espacio libre en disco"
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
      Height          =   350
      Left            =   90
      TabIndex        =   3
      Top             =   990
      Width           =   5775
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00404080&
      Caption         =   "Revisar tamaño de las tapas de los discos (tapa.jpg)"
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
      Height          =   350
      Left            =   120
      TabIndex        =   1
      Top             =   630
      Width           =   5775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Realizar mantenimiento ahora"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   0
      Top             =   1920
      Width           =   5805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "3PM mantenimiento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Width           =   5985
   End
End
Attribute VB_Name = "frmMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaNewFicha
            'si ya hay 9 cargados se traga las fichas
            If CREDITOS <= MaximoFichas Then
                OnOffCAPS vbKeyScrollLock, True
                CREDITOS = CREDITOS + TemasPorCredito
                SumarContadorCreditos TemasPorCredito
                'grabar cant de creditos
                EscribirArch1Linea AP + "creditos.tbr", Trim(Str(CREDITOS))
                If CREDITOS >= 10 Then
                    frmIndex.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Else
                    frmIndex.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                End If
            Else
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, False
            End If
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            End
    End Select
End Sub

