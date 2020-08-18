VERSION 5.00
Begin VB.Form frmAddRemoveMusic 
   BackColor       =   &H0000C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quitar musica de 3PM"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ver estadisticas de tema"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Visible         =   0   'False
      Width           =   2925
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver estadisticas del disco elegido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7140
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2670
      Width           =   1335
   End
   Begin VB.CommandButton cmdKillArch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Eliminar archivos elegidos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7170
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5400
      Width           =   2925
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Eliminar carpetas elegidas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6750
      Width           =   2925
   End
   Begin VB.ListBox lstTEMAS 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4890
      Left            =   6570
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.ListBox lstCarpetas 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   6015
   End
   Begin VB.Label lblInfoDisco 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informacion del disco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1365
      Left            =   7050
      TabIndex        =   9
      Top             =   6360
      Width           =   4785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Temas ciorrespondientes al disco elegido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6540
      TabIndex        =   6
      Top             =   150
      Width           =   3795
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la carpeta o los archivos que desee eliminar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   90
      Width           =   6105
   End
End
Attribute VB_Name = "frmAddRemoveMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MTXfiles() As String 'archivos en lstTEMAS , desde uno empieza

Private Sub cmdKillArch_Click()
    If lstTEMAS.SelCount = 0 Then
        MsgBox "No hay archivos seleccionados"
        Exit Sub
    End If
    
    If MsgBox("Esta seguro que desea eliminar los archivos elegidos?", _
        vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo NOBORRA
        Dim TotSel As Long, FileSel As String
        TotSel = lstTEMAS.SelCount
        For AA = 0 To lstTEMAS.ListCount - 1
            If lstTEMAS.Selected(AA) Then
                'en la matriz empieza en 1 y lst empieza en 0
                FileSel = txtInLista(MTXfiles(AA + 1), 0, ",")
                FSO.DeleteFile FileSel, True
                WriteTBRLog "Se borro el archivo " + FileSel, True
            End If
        Next
        'actualizar todo
        Call lstCarpetas_Click
        MsgBox "Los archivos se eliminaron correctamente"
    End If
    InfoDisco lblInfoDisco
    Exit Sub
NOBORRA:
    MsgBox "No se ha podido borrar uno mas temas, compruebe " + _
    "que no esten abiertos. Error numero:" + CStr(Err.Number) + _
    " Descripcion interna: " + Err.Description
End Sub

Private Sub Command1_Click()
    If MsgBox("Esta seguro que desea eliminar las carpetas elegidas?", _
        vbQuestion + vbYesNo) = vbYes Then
        On Local Error GoTo NOBORRA
        Dim TotSel As Long, CarpSel As String
        TotSel = lstCarpetas.SelCount
        For AA = 0 To lstCarpetas.ListCount - 1
            If lstCarpetas.Selected(AA) Then
                'en la matriz empieza en 1 y lst empieza en 0
                CarpSel = txtInLista(MATRIZ_DISCOS(AA + 1), 0, ",")
                If txtInLista(CarpSel, 99999, "\") = "01- Los mas escuchados" Then
                    MsgBox "No se puede borrar la carpeta del ranking"
                Else
                    FSO.DeleteFolder CarpSel, True
                    WriteTBRLog "Se borro la carpeta " + CarpSel, True
                End If
            End If
        Next
        'actualizar todo
        CargarCarpetas
        MsgBox "Las carpetas se eliminaron correctamente"
    End If
    Exit Sub
NOBORRA:
    MsgBox "No se ha podido borrar uno mas carpetas, compruebe " + _
    "que no esten abiertas. Error numero:" + CStr(Err.Number) + _
    " Descripcion interna: " + Err.Description

End Sub

Public Sub CargarCarpetas()
    lstCarpetas.Clear 'si no se duplican todos
    For a = 1 To UBound(MATRIZ_DISCOS)
        lstCarpetas.AddItem txtInLista(MATRIZ_DISCOS(a), 0, ",")
    Next
    lstCarpetas.Selected(0) = True
End Sub

Private Sub Command2_Click()
    'mostrar la cantidad de veces que se escucho este disco, cuanto se escucho el mas y cuanto el menos
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

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
                    frmINDEX.lblCreditos = "Creditos: " + Trim(Str(CREDITOS))
                Else
                    frmINDEX.lblCreditos = "Creditos: 0" + Trim(Str(CREDITOS))
                End If
            Else
                'apagar el fichero electronico
                OnOffCAPS vbKeyScrollLock, False
            End If
    End Select
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    'mostrar la lista de carpetas cargadas en 3PM
    CargarCarpetas
    InfoDisco lblInfoDisco
End Sub

Private Sub lstCarpetas_Click()
    lstTEMAS.Clear
    'mostrar los temas de esta carpeta solo si hay una sola carpeta elegida
    If lstCarpetas.SelCount > 1 Then
        Command2.Enabled = False
        lstTEMAS.AddItem "No hay vista disponible, multiples carpetas seleccionadas"
        lstTEMAS.Enabled = False
        GoTo FIN
    Else
        Command2.Enabled = True
        lstTEMAS.Enabled = True
    End If
    
    ReDim Preserve MTXfiles(0)
    
    MTXfiles = ObtenerArchMM(lstCarpetas)
    
    If UBound(MTXfiles) = 0 Then
        lstTEMAS.AddItem "No hay temas multimedia en esta carpeta"
        lstTEMAS.Enabled = False
    Else
        For a = 1 To UBound(MTXfiles)
            lstTEMAS.AddItem txtInLista(MTXfiles(a), 1, ",")
            lstTEMAS.Enabled = True
        Next
    End If
FIN:
    cmdKillArch.Enabled = lstTEMAS.Enabled
End Sub
