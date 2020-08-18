VERSION 5.00
Begin VB.Form frmTemasDeDisco 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.ListBox lstTemas 
      BackColor       =   &H00404080&
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0E0FF&
      Height          =   8565
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   11895
   End
   Begin VB.Label lblIndicaciones 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "Utilize las flechas para desplazarse sobre los distintos temas, OK=escuchar. Escape=Salir"
      BeginProperty Font 
         Name            =   "HandelGotDLig"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   225
      Left            =   60
      TabIndex        =   1
      Top             =   8700
      Width           =   11955
   End
End
Attribute VB_Name = "frmTemasDeDisco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        ESTOY = 0
        Unload Me
    End If
    If KeyCode = vbKeyReturn Then
        Dim temaElegido As String
        temaElegido = UbicDiscoActual + lstTemas + ".mp3"
        'temaElegido = txtInLista(MATRIZ_TEMAS(lstTemas.ListIndex + 1), 0, ",")
        
        'si esta ejecutando pasa a la lista de reproducción
        If ESTOY_REPRODUCIENDO Then
            'pasar a la lista de reproducción
            Dim NewIndLista As Long
            NewIndLista = UBound(MATRIZ_LISTA)
            ReDim Preserve MATRIZ_LISTA(NewIndLista + 1)
            'se graba en Matriz_Listas como patah, nombre(sin ".mp3")
            MATRIZ_LISTA(NewIndLista + 1) = temaElegido + "," + lstTemas + " / " + FSO.GetBaseName(UbicDiscoActual)
            CargarProximosTemas
        Else
            'TEMA_REPRODUCIENDO y ESTOY_REPRODUCIENDO se cargan en ejecutartema
            EjecutarTema temaElegido
        End If
        'pase lo que pase me vuelvo a los discos y cierro ventana actual
        ESTOY = 0
        Unload Me
    End If
    If KeyCode = vbKeyX Then
        If lstTemas.ListIndex < lstTemas.ListCount - 1 Then lstTemas.ListIndex = lstTemas.ListIndex + 1
    End If
    If KeyCode = vbKeyZ Then
        If lstTemas.ListIndex > 0 Then lstTemas.ListIndex = lstTemas.ListIndex - 1
    End If
End Sub

Private Sub Form_Load()
    lstTemas.Left = Screen.Width / 2 - lstTemas.Width / 2
    lstTemas.Top = Screen.Height / 2 - lstTemas.Height / 2
    lblIndicaciones.Left = lstTemas.Left
    lblIndicaciones.Top = lstTemas.Top + lstTemas.Height
End Sub

