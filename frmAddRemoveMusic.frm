VERSION 5.00
Begin VB.Form frmAddRemoveMusic 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quitar musica de 3PM"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Rename Disco"
      Height          =   555
      Left            =   3030
      TabIndex        =   22
      Top             =   7140
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.ListBox lstTODO 
      BackColor       =   &H00000000&
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
      Height          =   7470
      Left            =   7200
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Revisar tamaño de las portadas"
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
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   3255
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Generar estadistica de discos"
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
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7920
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Agregar Música"
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
      Left            =   7230
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7560
      Width           =   3255
   End
   Begin VB.ListBox lstEstadisticas 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   7290
      MultiSelect     =   2  'Extended
      TabIndex        =   14
      Top             =   450
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Nuevo texto"
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
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3780
      Width           =   1450
   End
   Begin VB.CommandButton cmdKillTXT 
      BackColor       =   &H0080C0FF&
      Caption         =   "Quitar texto"
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
      Left            =   4170
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3780
      Width           =   1450
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nueva foto"
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
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2550
      Width           =   1450
   End
   Begin VB.CommandButton cmdKillTapa 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Quitar foto"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2550
      Width           =   1450
   End
   Begin VB.TextBox txtDataTXT 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4170
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmAddRemoveMusic.frx":0000
      Top             =   2940
      Width           =   2955
   End
   Begin VB.ListBox lstCarpetasShow 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7665
      Left            =   60
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   300
      Width           =   4035
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
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
      Height          =   1125
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7530
      Width           =   1335
   End
   Begin VB.CommandButton cmdKillArch 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Eliminar temas elegidos"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7980
      Width           =   2955
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Eliminar discos elegidos"
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
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7980
      Width           =   4005
   End
   Begin VB.ListBox lstTEMAS 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   4170
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4380
      Width           =   2985
   End
   Begin VB.ListBox lstCarpetas 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   30
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   8400
      Width           =   7095
   End
   Begin VB.Label lblKB 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2820
      TabIndex        =   20
      Top             =   60
      Width           =   1485
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddRemoveMusic.frx":0006
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2115
      Index           =   3
      Left            =   7290
      TabIndex        =   17
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estadisticas del disco elegido. Cantidad de veces que se escucho el disco"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Index           =   2
      Left            =   7290
      TabIndex        =   15
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   2505
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   30
      Width           =   2640
   End
   Begin VB.Label lblInfoDisco 
      Alignment       =   2  'Center
      BackColor       =   &H000040C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Informacion del disco"
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
      Height          =   1185
      Left            =   7200
      TabIndex        =   7
      Top             =   6150
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Temas del disco elegido"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   4170
      TabIndex        =   6
      Top             =   4170
      Width           =   2925
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "DISCOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   30
      Width           =   1575
   End
End
Attribute VB_Name = "frmAddRemoveMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MTXfiles() As String 'archivos en lstTEMAS , desde uno empieza
Dim CmdLg As New CommonDialog

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

Private Sub cmdKillTapa_Click()
    If MsgBox("¿Esta seguro que desea eliminar la imagen elegida?", vbYesNo + vbCritical) = vbNo Then Exit Sub
    FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    'refrescar la imagen
    TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
    cmdKillTapa.Enabled = False
End Sub

Private Sub cmdKillTXT_Click()
    If MsgBox("¿Esta seguro que desea eliminar el texto elegido?", vbYesNo + vbCritical) = vbNo Then Exit Sub
    FSO.DeleteFile lstCarpetas + "\data.txt", True
    cmdKillTXT.Enabled = False
    txtDataTXT = "NO EXISTE"
End Sub

Private Sub Command1_Click()
    'si borro una vez esta todo OK
    ' si entro por segunda vez ya matriz_discos tiene distintos indices
    'con respecto a las listas!!!!!!!!!!!!
    If MsgBox("Esta seguro que desea eliminar las carpetas elegidas?", _
        vbQuestion + vbYesNo) = vbYes Then
        On Local Error GoTo NOBORRA
        Dim TotSel As Long, CarpSel As String
        TotSel = lstCarpetasShow.SelCount
        For AA = 0 To lstCarpetasShow.ListCount - 1
            If lstCarpetasShow.Selected(AA) Then
                'NO USAR matriz_discos ya que si se borra una vez se separan
                'los indices de la matriz con los del listado (que se van quitando y en la matriz no)
                ''en la matriz empieza en 1 y lst empieza en 0
                'CarpSel = txtInLista(MATRIZ_DISCOS(AA + 1), 0, ",")
                CarpSel = lstCarpetas.List(AA)
                If Right(CarpSel, 22) = "01- Los mas escuchados" Then
                    MsgBox "No se puede borrar la carpeta del ranking"
                Else
                    FSO.DeleteFolder CarpSel, True
                    WriteTBRLog "Se borro la carpeta " + CarpSel, True
                End If
            End If
        Next
        MsgBox "Las carpetas se eliminaron correctamente"
        'actualizar todo
        CargarCarpetas
        InfoDisco lblInfoDisco
    End If
    Exit Sub
NOBORRA:
    MsgBox "No se ha podido borrar uno mas carpetas, compruebe " + _
    "que no esten abiertas. Error numero:" + CStr(Err.Number) + _
    " Descripcion interna: " + Err.Description

End Sub

Public Sub CargarCarpetas()
    lstCarpetas.Clear 'si no se duplican todos
    lstCarpetasShow.Clear
    For A = 1 To UBound(MATRIZ_DISCOS)
        Dim ThisFolder As String, TamTapa As Double
        ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
        'ver si existen o se borraron
        If FSO.FolderExists(ThisFolder) Then ' And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
            lstCarpetas.AddItem txtInLista(MATRIZ_DISCOS(A), 0, ",")
            lstCarpetasShow.AddItem txtInLista(MATRIZ_DISCOS(A), 1, ",")
        End If
    Next
    lstCarpetasShow.Selected(0) = True
End Sub

Private Sub Command2_Click()
    frmAddMusic.Show 1
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    CmdLg.DialogTitle = "Eliga la nueva imagen del disco elegido"
    CmdLg.Filter = "Imagen JPG o GIF|*.jpg;*.gif"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    'ver el tamaño!!!
    Dim TamTapa As Double
    TamTapa = FileLen(CmdLg.FileName)
    TamTapa = Round(TamTapa / 1024, 2) 'son KB
    If TamTapa > 20 Then
        If MsgBox("tbrSoft recomienda imagenes no mayores a 8 KB, esta " + _
            "imagen tine " + CStr(TamTapa) + " KB. ¿Esta seguro que desea usarla?" + _
            " Puede afectar el rendimiento del equipo!", vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    If TamTapa > 200 Then
        If MsgBox("Imagen demasiado pesada. Despues de la advertencia, ¿aun desea usarla?" + _
            " Puede afectar el rendimiento del equipo!!!", vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    If cmdKillTapa.Enabled Then
        If MsgBox("¿Esta seguro que desea reemplazar la imagen existente?", vbYesNo + vbCritical) = vbNo Then Exit Sub
        FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    End If
    FSO.CopyFile CmdLg.FileName, lstCarpetas + "\tapa.jpg", True
    TapaCD.Picture = LoadPicture(lstCarpetas + "\tapa.jpg")
End Sub

Private Sub Command5_Click()
    If Command5.Caption = "Confirmar" Then
        txtDataTXT.BackColor = &HE0E0E0   'color original
        txtDataTXT.Locked = True
        Command5.Caption = "Nuevo texto"
        'grabar el texto
        Set TE = FSO.CreateTextFile(lstCarpetas + "\data.txt", True)
        TE.Write txtDataTXT
        TE.Close
    Else
        txtDataTXT.BackColor = vbWhite '&H00E0E0E0&'color original
        txtDataTXT.SetFocus
        txtDataTXT.SelStart = 0
        txtDataTXT.SelLength = Len(txtDataTXT)
        txtDataTXT.Locked = False
        Command5.Caption = "Confirmar"
    End If
    
End Sub

Private Sub Command6_Click()
    If Command6.Caption = "Generar estadistica de discos" Then
        lstTODO.Clear
        lstTODO.AddItem " RANK DISCOS (primero los menos escuhados)"
        'pasar por todos los discos y medir de cada uno las escuchadas
        Dim ThisFolder As String
        Dim Carp As String
        For A = 1 To UBound(MATRIZ_DISCOS)
            Carp = txtInLista(MATRIZ_DISCOS(A), 1, ",")
            ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
            'ver si existen o se borraron
            If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
                lstTODO.AddItem STRceros(ContarLisen(Carp), 4) + ": " + Carp
            End If
            lstTODO.Visible = True
            Command6.Caption = "Quitar estadistica de discos"
        Next
    Else
        lstTODO.Visible = False
        Command6.Caption = "Generar estadistica de discos"
    
    End If
End Sub

Private Sub Command7_Click()
    Dim ThisFolder As String, TamTapa As Double
    Dim MasDe20Kb As String, MasDe300Kb As String
    Dim TapasGrandes As Long, TapasMuyGrandes As Long
    TapasGrandes = 0: TapasMuyGrandes = 0
        
    For A = 1 To UBound(MATRIZ_DISCOS)
        ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
        If FSO.FileExists(ThisFolder + "\tapa.jpg") Then
            TamTapa = FileLen(ThisFolder + "\tapa.jpg")
            TamTapa = Round(TamTapa / 1024, 2)
            If TamTapa > 200 Then
                MasDe300Kb = MasDe300Kb + FSO.GetBaseName(ThisFolder) + vbCrLf
                TapasMuyGrandes = TapasMuyGrandes + 1
            Else
                If TamTapa > 20 Then
                    MasDe20Kb = MasDe20Kb + FSO.GetBaseName(ThisFolder) + vbCrLf
                    TapasGrandes = TapasGrandes + 1
                End If
            End If
        End If
    Next
    
    Dim MsGrandes As String
    Dim MsMuyGrandes As String
    If TapasGrandes > 0 Then MsGrandes = "Hay " + CStr(TapasGrandes) + " tapas de mas de 20 Kb. Estas son:" + vbCrLf + MasDe20Kb
    If TapasMuyGrandes > 0 Then MsMuyGrandes = "Hay " + CStr(TapasMuyGrandes) + " tapas de mas de 200 Kb. Estas son:" + vbCrLf + MasDe300Kb
    Dim MSGfinal As String
    MSGfinal = MsGrandes + vbCrLf + MsMuyGrandes
    If Len(MSGfinal) < 5 Then
        MsgBox "Todas las portadas de los discos tiene tamaños correctos"
    Else
        MsgBox MSGfinal
    End If
End Sub

Private Sub Command8_Click()
    Dim NewFolder As String
    NewFolder = InputBox("Ingrese el nuevo nombre para el disco", "Nuevo nombre")
    If NewFolder <> "" Then
        'primero corregir y acomodar el ranking para no perder los votos!!
        '?????????¿¿¿¿¿¿¿¿¿¿¿¿
        
        
        Name lstCarpetas As AP + "discos\" + NewFolder
        lstCarpetasShow.List(lstCarpetas.ListIndex) = NewFolder
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        
        Case TeclaCerrarSistema
            OnOffCAPS vbKeyCapital, False
            If ApagarAlCierre Then APAGAR_PC
            'no puedo usar do stop porque lanza el evento ENDPLAY y esto produce un EMPEZARSIGUIENTE
            'que se come un tema de la lista
            frmIndex.MP3.DoClose
            End
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = TeclaNewFicha Then
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
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "\radilav.cfg", CStr(CreditosValidar)
            
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
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
        lstTEMAS.AddItem "No hay vista disponible"
        lstTEMAS.AddItem "Multiples carpetas seleccionadas"
        lstTEMAS.Enabled = False
    Else
        lstTEMAS.Enabled = True
        ReDim Preserve MTXfiles(0)
        MTXfiles = ObtenerArchMM(lstCarpetas)
        If UBound(MTXfiles) = 0 Then
            lstTEMAS.AddItem "No hay temas multimedia en esta carpeta"
            lstTEMAS.Enabled = False
        Else
            For A = 1 To UBound(MTXfiles)
                lstTEMAS.AddItem txtInLista(MTXfiles(A), 1, ",")
                lstTEMAS.Enabled = True
            Next
        End If
    End If
    cmdKillArch.Enabled = lstTEMAS.Enabled
    'mostrar la tapa si la tiene
    Dim TapaArch As String
    TapaArch = lstCarpetas + "\tapa.jpg"
    If FSO.FileExists(TapaArch) Then
        TapaCD.Picture = LoadPicture(TapaArch)
        cmdKillTapa.Enabled = True
        Dim TamTapa As Double
        TamTapa = FileLen(TapaArch)
        TamTapa = Round(TamTapa / 1024, 2)
        lblKB = CStr(TamTapa) + " KB"
    Else
        TapaCD.Picture = LoadPicture(AP + "tapa.jpg")
        cmdKillTapa.Enabled = False
        lblKB = "8 KB"
    End If
    'mostrar el texto si existe
    Dim DataTXT As String
    DataTXT = lstCarpetas + "\data.txt"
    If FSO.FileExists(DataTXT) Then
        Dim TE2 As TextStream
        Set TE2 = FSO.OpenTextFile(DataTXT, ForReading, False)
        txtDataTXT = TE2.ReadAll
        cmdKillTXT.Enabled = True
    Else
        txtDataTXT = "NO EXISTE"
        cmdKillTXT.Enabled = False
    End If
    'ver el ranking y analizar las estadísticas del disco
    'leer los discos y ver uno por uno (de mayor a menor) que temas se escucharon
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        lstEstadisticas.AddItem "No hay informacion"
        Exit Sub
    End If
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    lstEstadisticas.Clear
    Dim ThisLine As String, ThisDISCO As String, ThisTEMA As String, CantLisen As Long
    Dim TotLisen As Long, nPuesto As Long
    TotLisen = 0: nPuesto = 0
    Do While Not TE.AtEndOfStream
        nPuesto = nPuesto + 1
        ThisLine = TE.ReadLine
        'cantidad,pathfull,toShow(disco-tema),carpeta
        ThisDISCO = txtInLista(ThisLine, 3, ",")
        If ThisDISCO = lstCarpetasShow Then
            CantLisen = Val(txtInLista(ThisLine, 0, ","))
            ThisTEMA = txtInLista(FSO.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
            lstEstadisticas.AddItem ThisTEMA + ": " + CStr(CantLisen) + " (RANK: " + CStr(nPuesto) + ")"
            TotLisen = TotLisen + CantLisen
        End If
    Loop
    If TotLisen = 0 Then
        lstEstadisticas.AddItem "Este disco no se ha escuchado todavia"
    Else
        lstEstadisticas.AddItem ""
        lstEstadisticas.AddItem "   TOTAL: " + CStr(TotLisen)
    End If
    TE.Close
End Sub

Private Sub lstCarpetasShow_Click()
    lstCarpetas.ListIndex = lstCarpetasShow.ListIndex
End Sub

Public Function ContarLisen(Carpeta As String) As Long
    Set TE = FSO.OpenTextFile(AP + "ranking.tbr", ForReading, False)
    Dim ThisLine As String, ThisDISCO As String, ThisTEMA As String, CantLisen As Long
    Dim TotLisen As Long, nPuesto As Long
    TotLisen = 0: nPuesto = 0
    Do While Not TE.AtEndOfStream
        nPuesto = nPuesto + 1
        ThisLine = TE.ReadLine
        'cantidad,pathfull,toShow(disco-tema),carpeta
        ThisDISCO = txtInLista(ThisLine, 3, ",")
        If ThisDISCO = Carpeta Then
            CantLisen = Val(txtInLista(ThisLine, 0, ","))
            ThisTEMA = txtInLista(FSO.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
            TotLisen = TotLisen + CantLisen
        End If
    Loop
    TE.Close
    ContarLisen = TotLisen

End Function

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(Str(n))
    Dim DIF As Integer
    DIF = Cifras - Len(STRn)
    If DIF > 0 Then
        Dim CEROstr As String
        CEROstr = String(DIF, "0")
        STRceros = CEROstr + STRn
    Else
        STRceros = STRn
    End If
    
End Function

