VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmAddRemoveMusic 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Quitar música de 3PM"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton command7 
      Height          =   345
      Left            =   7230
      TabIndex        =   32
      Top             =   8280
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "revisar tamaño de las portadas"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command6 
      Height          =   345
      Left            =   7230
      TabIndex        =   31
      Top             =   7920
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "generar estadística de discos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command2 
      Height          =   345
      Left            =   7230
      TabIndex        =   30
      Top             =   7560
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "agregar música"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command15 
      Height          =   345
      Left            =   7230
      TabIndex        =   29
      Top             =   7200
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar estadísticas a disco"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command3 
      Height          =   675
      Left            =   10650
      TabIndex        =   28
      Top             =   7950
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1191
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command5 
      Height          =   375
      Left            =   2070
      TabIndex        =   27
      Top             =   7980
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   661
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "nuevo texto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton cmdKillTXT 
      Height          =   405
      Left            =   420
      TabIndex        =   26
      Top             =   7950
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "quitar texto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton cmdKillArch 
      Height          =   435
      Left            =   4170
      TabIndex        =   25
      Top             =   7950
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "eliminar canciones elegidas"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command1 
      Height          =   405
      Left            =   60
      TabIndex        =   24
      Top             =   6720
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   714
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "eliminar discos elegidos"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command8 
      Height          =   615
      Left            =   3540
      TabIndex        =   23
      Top             =   5580
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "rename disco"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command4 
      Height          =   435
      Left            =   5700
      TabIndex        =   16
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "nueva foto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton cmdKillTapa 
      Height          =   435
      Left            =   4350
      TabIndex        =   15
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   767
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "quitar foto"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.ListBox lstTODO 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   7140
      IntegralHeight  =   0   'False
      Left            =   7200
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   4665
   End
   Begin VB.ListBox lstEstadisticas 
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
      ForeColor       =   &H00404040&
      Height          =   3375
      Left            =   7290
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   450
      Width           =   4575
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
      ForeColor       =   &H00404040&
      Height          =   765
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   7140
      Width           =   4095
   End
   Begin VB.ListBox lstCarpetasShow 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4500
      IntegralHeight  =   0   'False
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   5
      Top             =   2190
      Width           =   4035
   End
   Begin VB.ListBox lstTEMAS 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   3030
      IntegralHeight  =   0   'False
      ItemData        =   "frmAddRemoveMusic.frx":0000
      Left            =   4170
      List            =   "frmAddRemoveMusic.frx":0040
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4860
      Width           =   2985
   End
   Begin VB.ListBox lstCarpetas 
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00404040&
      Height          =   270
      IntegralHeight  =   0   'False
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   8400
      Width           =   7095
   End
   Begin VB.ListBox lstOrigenes 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1380
      IntegralHeight  =   0   'False
      ItemData        =   "frmAddRemoveMusic.frx":008B
      Left            =   90
      List            =   "frmAddRemoveMusic.frx":00A4
      TabIndex        =   12
      Top             =   540
      Width           =   7035
   End
   Begin tbrFaroButton.fBoton Command9 
      Height          =   285
      Left            =   4380
      TabIndex        =   17
      Top             =   -30
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "agregar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command10 
      Height          =   285
      Left            =   4380
      TabIndex        =   18
      Top             =   240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "quitar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command14 
      Height          =   285
      Left            =   5280
      TabIndex        =   19
      Top             =   -30
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "subir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command13 
      Height          =   285
      Left            =   5280
      TabIndex        =   20
      Top             =   240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "bajar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton Command11 
      Height          =   285
      Left            =   6180
      TabIndex        =   21
      Top             =   -30
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "ayuda"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton command12 
      Height          =   285
      Left            =   6180
      TabIndex        =   22
      Top             =   240
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   503
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "grabar"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "La música se cargará en este orden."
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
      Height          =   225
      Index           =   5
      Left            =   150
      TabIndex        =   14
      Top             =   300
      Width           =   3915
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Ubicaciones de música (solo SL)"
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
      Height          =   225
      Index           =   4
      Left            =   150
      TabIndex        =   13
      Top             =   60
      Width           =   4125
   End
   Begin VB.Label lblKB 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8888 KB"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   195
      Left            =   6000
      TabIndex        =   10
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddRemoveMusic.frx":00BD
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
      Height          =   2115
      Index           =   3
      Left            =   7290
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   0
      Width           =   4575
   End
   Begin VB.Image TapaCD 
      BorderStyle     =   1  'Fixed Single
      Height          =   1905
      Left            =   4260
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   2850
   End
   Begin VB.Label lblInfoDisco 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00404040&
      Height          =   1005
      Left            =   7200
      TabIndex        =   4
      Top             =   6150
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Canciones del disco elegido"
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
      Left            =   4320
      TabIndex        =   3
      Top             =   4650
      Width           =   2925
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Discos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   1920
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
    On Error GoTo MiErr
    
    tERR.Anotar "acjy"
    If lstTEMAS.SelCount = 0 Then
        MsgBox TR.Trad("No hay archivos seleccionados%99%")
        Exit Sub
    End If
    
    msg = TR.Trad("Esta seguro que desea eliminar los archivos elegidos?%99%")
        
    If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
        On Error GoTo NOBORRA
        Dim TotSel As Long, FileSel As String
        TotSel = lstTEMAS.SelCount
        For AA = 0 To lstTEMAS.ListCount - 1
            
            If lstTEMAS.Selected(AA) Then
                'en la matriz empieza en 1 y lst empieza en 0
                FileSel = txtInLista(MTXfiles(AA + 1), 0, "#")
                fso.DeleteFile FileSel, True
                tERR.Anotar "acka", AA, TotSel, FileSel
            End If
        Next
        'actualizar todo
        Call lstCarpetas_Click
        MsgBox TR.Trad("Los archivos se eliminaron correctamente%99%")
        
    End If
    tERR.Anotar "ackb"
    InfoDisco lblInfoDisco
    Exit Sub
NOBORRA:
    MsgBox TR.Trad("No se ha podido borrar uno mas temas, compruebe " + _
        "que no esten abiertos.%99%") + vbCrLf + _
    "Error numero:" + CStr(Err.Number) + vbCrLf + _
    "Descripcion interna del error: " + Err.Description
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acji"
    Resume Next
End Sub

Private Sub cmdKillTapa_Click()
    On Error GoTo MiErr
    msg = TR.Trad("¿Esta seguro que desea eliminar la imagen elegida?%99%")

    If MsgBox(msg, vbYesNo + vbCritical) = vbNo Then Exit Sub
    fso.DeleteFile lstCarpetas + "\tapa.jpg", True
    tERR.Anotar "ackc", lstCarpetas
    'refrescar la imagen
    'ver si es superlicencia y usa otra tapa predeterminada
    IMF = GetTpPred
    
    TapaCD.Picture = LoadPicture(IMF)
    cmdKillTapa.Enabled = False
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjj"
    Resume Next
End Sub

Private Sub cmdKillTXT_Click()
    On Error GoTo MiErr
    msg = TR.Trad("¿Esta seguro que desea eliminar el texto elegido?%99%")
    
    If MsgBox(msg, vbYesNo + vbCritical) = vbNo Then Exit Sub
    tERR.Anotar "ackd", lstCarpetas
    fso.DeleteFile lstCarpetas + "\data.txt", True
    cmdKillTXT.Enabled = False
    txtDataTXT = "NO EXISTE"
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjk"
    Resume Next
End Sub

Private Sub Command1_Click()
    On Error GoTo MiErr
    
    'si borro una vez esta todo OK
    ' si entro por segunda vez ya matriz_discos tiene distintos indices
    'con respecto a las listas!!!!!!!!!!!!
    msg = TR.Trad("Esta seguro que desea eliminar las carpetas elegidas?%99%")
    
    tERR.Anotar "acke"
    If MsgBox(msg, vbQuestion + vbYesNo) = vbYes Then
        tERR.Anotar "ackf"
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
                tERR.Anotar "ackg", AA, CarpSel
                If Right(CarpSel, 22) = "_Los mas escuchados" Then
                    MsgBox TR.Trad("No se puede borrar la carpeta del ranking%99%")
                Else
                    fso.DeleteFolder CarpSel, True
                End If
            End If
        Next
        
        MsgBox TR.Trad("Las carpetas se eliminaron correctamente%99%")
        
        'actualizar todo
        CargarCarpetas
        InfoDisco lblInfoDisco
    End If
    Exit Sub
NOBORRA:

    MsgBox TR.Trad("No se ha podido borrar uno mas carpetas, compruebe que " + _
        "no esten abiertas.%99%") + vbCrLf + _
        "Error numero:" + CStr(Err.Number) + vbCrLf + _
        "Descripcion interna del error: " + Err.Description

    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjk"
    Resume Next
End Sub

Public Sub CargarCarpetas()
    On Error GoTo MiErr
    lstCarpetas.Clear 'si no se duplican todos
    lstCarpetasShow.Clear
    For A = 0 To UBound(MATRIZ_DISCOS)
        Dim ThisFolder As String, TamTapa As Double
        ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
        tERR.Anotar "ackh", A, UBound(MATRIZ_DISCOS), ThisFolder
        'ver si existen o se borraron
        'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
        If fso.FolderExists(ThisFolder) Then
            lstCarpetas.AddItem txtInLista(MATRIZ_DISCOS(A), 0, ",")
            lstCarpetasShow.AddItem txtInLista(MATRIZ_DISCOS(A), 1, ",")
        End If
    Next
    'arreglado 27/2 no se fijaba si había indices
    If lstCarpetasShow.ListCount > 0 Then
        lstCarpetasShow.Selected(0) = True
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjl"
    Resume Next
End Sub

Private Sub Command10_Click()
    'ver sio hay mas de uno!!!
    If lstOrigenes.ListCount = 1 Then
        MsgBox TR.Trad("No se puede quitar el ultimo origen de discos. " + _
            "Debe haber uno!%98%Un origen de discos es un " + _
            "directorio que contiene una o mas directorios que son discos%99%")
        Exit Sub
    End If
    'borrar
    lstOrigenes.RemoveItem lstOrigenes.ListIndex
    lstOrigenes.ListIndex = 0
End Sub

Private Sub Command11_Click()
    TR.SetVars "3PM"
    MsgBox TR.Trad("3PM en modo SuperLicencia permite la carga de musica " + _
        "desde diferentes ubicaciones." + vbCrLf + _
        "Esto permite utilizar mas de un disco rígido o diferentes particiones " + _
        "de un mismo disco." + vbCrLf + _
        "Además %01% permite dar un mejor orden a la " + _
        "música y videos expuestos al público." + vbCrLf + _
        "Al iniciar el sistema %01% leerá en primer lugar la musica de " + _
        "la primera ubicación hasta llegar a la ultima." + vbCrLf + _
        "De esta forma se podrá separar la musica en diferentes " + _
        "ritmos (rock, pop, folcklore, popular, etc)." + vbCrLf + _
        "Dentro de cada ubicación los discos estarán ordenados " + _
        "alfabéticamente permitiendo al usuario final un acceso mas " + _
        "sencillo a la música buscada%99%"), vbQuestion, "UBICACIONES DE MUSICA"
End Sub

Private Sub Command12_Click()
    Dim TMPs As String
    TMPs = ""
    For A = 0 To lstOrigenes.ListCount - 1
        'al ultimo no pongo asterisco para que al dividir no quede uno vacio al ultimo!
        If A < lstOrigenes.ListCount - 1 Then
            TMPs = TMPs + lstOrigenes.List(A) + "*"
        Else
            TMPs = TMPs + lstOrigenes.List(A)
        End If
    Next A
    EscribirArch1Linea GPF("origs"), TMPs
    MsgBox TR.Trad("Los cambios se han grabado satisfactoriamente%99%")
End Sub

Private Sub Command13_Click()
    On Error GoTo MiErr
    
    'poner el que esta elegido abajo -BAJAR-
    'si es el ultimo (X) o no hay elegido (-1) se va
    If lstOrigenes.ListIndex = (lstOrigenes.ListCount - 1) Then Exit Sub
    If lstOrigenes.ListIndex = -1 Then Exit Sub
    tERR.Anotar "acki", lstOrigenes.ListIndex, lstOrigenes.ListCount
    Dim TMPlst As String, NumSube As Long
    NumSube = lstOrigenes.ListIndex
    TMPlst = lstOrigenes.List(NumSube + 1)
    'el anterior se transforma en el que sube
    lstOrigenes.List(NumSube + 1) = lstOrigenes
    lstOrigenes.List(NumSube) = TMPlst
    lstOrigenes.ListIndex = NumSube + 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjm"
    Resume Next
End Sub

Private Sub Command14_Click()
    On Error GoTo MiErr
    'poner el que esta elegido arriba -SUBIR-
    'si es el primero (0) o no hay elegido (-1) se va
    If lstOrigenes.ListIndex < 1 Then Exit Sub
    
    Dim TMPlst As String, NumSube As Long
    NumSube = lstOrigenes.ListIndex
    TMPlst = lstOrigenes.List(NumSube - 1)
    tERR.Anotar "ackj", lstOrigenes.ListIndex, lstOrigenes.ListCount
    'el anterior se transforma en el que sube
    lstOrigenes.List(NumSube - 1) = lstOrigenes
    lstOrigenes.List(NumSube) = TMPlst
    lstOrigenes.ListIndex = NumSube - 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjn"
    Resume Next
End Sub

Private Sub Command15_Click()
    
    'borra si existia
    If fso.FileExists("C:\STATS.TXT") Then fso.DeleteFile "C:\STATS.TXT", True
    
    Dim TE121 As TextStream
    
    Set TE121 = fso.OpenTextFile("C:\STATS.TXT", ForAppending, True)
    
    TE121.WriteLine "-------------------------"
    TE121.WriteLine TR.Trad("ESTADISTICAS SEGUN DISCOS%99%")
    TE121.WriteLine "-------------------------"
    TE121.WriteLine TR.Trad("Cantidad Reproducciones: DISCO%99%")
    Dim ThisFolder As String
    Dim Carp As String, A As Long
    lstTODO.Clear
    For A = 0 To UBound(MATRIZ_DISCOS)
        Carp = txtInLista(MATRIZ_DISCOS(A), 1, ",")
        ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
        tERR.Anotar "ackp2", A, UBound(MATRIZ_DISCOS), Carp, ThisFolder
        'ver si existen o se borraron
        'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
        If fso.FolderExists(ThisFolder) Then
            lstTODO.AddItem STRceros(ContarLisen(Carp), 4) + ": " + Carp
        End If
    Next A
    
    For A = 0 To lstTODO.ListCount - 1
        TE121.WriteLine lstTODO.List(A)
    Next A
    
    'GRABAR RANK DE CANCION
    TE121.WriteLine "-------------------------"
    TE121.WriteLine TR.Trad("ESTADISTICAS SEGUN CANCIONES%99%")
    TE121.WriteLine "-------------------------"
    TE121.WriteLine TR.Trad("Cantidad Reproducciones: CANCION%99%")
        'grabar en un txt las canciones
        Dim TE120 As TextStream
        Set TE120 = fso.OpenTextFile(GPF("rd3_444"), ForReading, True)
            Do While Not TE120.AtEndOfStream
                TE121.WriteLine TE120.ReadLine
            Loop
        TE120.Close
        Set TE120 = Nothing
    TE121.Close
    Set TE121 = Nothing
    MsgBox TR.Trad("Las estadisticas se han grabado sin problemas en C:\STATS.TXT%99%")
End Sub

Private Sub Command2_Click()
    frmAddMusic.Show 1
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error GoTo MiErr
    
    CmdLg.DialogTitle = TR.Trad("Eliga la nueva imagen del disco elegido%99%")
    CmdLg.Filter = TR.Trad("Imagen JPG o GIF%99%") + "|*.jpg;*.gif"
    
    tERR.Anotar "ackk"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    'ver el tamaño!!!
    Dim TamTapa As Double
    TamTapa = FileLen(CmdLg.FileName)
    TamTapa = Round(TamTapa / 1024, 2) 'son KB
    tERR.Anotar "ackl", TamTapa
    If TamTapa > TamanoTapaPermitido Then
        TR.SetVars CStr(TamTapa)
        msg = TR.Trad("tbrSoft recomienda imagenes no mayores a 20 KB." + vbCrLf + _
        "Esta imagen tine %01% KB. ¿Esta seguro que desea usarla?" + vbCrLf + _
            " Puede afectar el rendimiento del equipo!%99%")
    
        If MsgBox(msg, vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    If TamTapa > 200 Then
        msg = TR.Trad("Imagen demasiado pesada. Despues de la advertencia, " + _
            "¿aun desea usarla?" + vbCrLf + _
            " Puede afectar el rendimiento del equipo!%99%")
    
        If MsgBox(msg, vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    'RRRRRRRRRRRRRRR
    
    If cmdKillTapa.Enabled Then
        msg = TR.Trad("¿Esta seguro que desea reemplazar la imagen existente?%99%")
        
        If MsgBox(msg, vbYesNo + vbCritical) = vbNo Then Exit Sub
        tERR.Anotar "ackm", lstCarpetas
        fso.DeleteFile lstCarpetas + "\tapa.jpg", True
    End If
    tERR.Anotar "ackn", CmdLg.FileName
    fso.CopyFile CmdLg.FileName, lstCarpetas + "\tapa.jpg", True
    TapaCD.Picture = LoadPicture(lstCarpetas + "\tapa.jpg")
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjp"
    Resume Next
End Sub

Private Sub Command5_Click()
    On Error GoTo MiErr
    'ojo ARTIME que aqui el programa se guia por lo que esta escrito en el boton
    'o sea fijate en mayusculas y minusculas!!!
    
    msg = TR.Trad("Confirmar%99%")
    
    If Command5.Caption = msg Then
        txtDataTXT.BackColor = &HE0E0E0   'color original
        txtDataTXT.Locked = True
        
        Command5.Caption = TR.Trad("Nuevo texto%99%")
        
        'grabar el texto
        Set TE = fso.CreateTextFile(lstCarpetas + "\data.txt", True)
            TE.Write txtDataTXT
            tERR.Anotar "acko", txtDataTXT
        TE.Close
    Else
        txtDataTXT.BackColor = vbWhite '&H00E0E0E0&'color original
        txtDataTXT.SetFocus
        txtDataTXT.SelStart = 0
        txtDataTXT.SelLength = Len(txtDataTXT)
        txtDataTXT.Locked = False
        
        Command5.Caption = TR.Trad("Confirmar%99%")
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjq"
    Resume Next
End Sub

Private Sub Command6_Click()
    On Error GoTo MiErr
    'ojo ARTIME que aqui el programa se guia por lo que esta escrito en el boton
    'o sea fijate en mayusculas y minusculas!!!
    msg = TR.Trad("Generar estadistica de discos%98%Inicia " + _
        "calculos y genera un archivo de texto con estadísticas%99%")
        
    If Command6.Caption = msg Then
        lstTODO.Clear
        
        lstTODO.AddItem TR.Trad(" RANK DISCOS (primero los menos escuhados)%99%")
    
        'pasar por todos los discos y medir de cada uno las escuchadas
        Dim ThisFolder As String
        Dim Carp As String
        For A = 0 To UBound(MATRIZ_DISCOS)
            Carp = txtInLista(MATRIZ_DISCOS(A), 1, ",")
            ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
            tERR.Anotar "ackp", A, UBound(MATRIZ_DISCOS), Carp, ThisFolder
            'ver si existen o se borraron
            'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
            If fso.FolderExists(ThisFolder) Then
                lstTODO.AddItem STRceros(ContarLisen(Carp), 4) + ": " + Carp
            End If
            lstTODO.Visible = True
            
            Command6.Caption = TR.Trad("Quitar estadistica de discos%98%La " + _
                "lista de discos mas escuchados se puede poner y sacar, " + _
                "con este se saca%99%")
        Next
    Else
        lstTODO.Visible = False
        Command6.Caption = TR.Trad("Generar estadistica de discos%99%")
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjr"
    Resume Next
End Sub

Private Sub Command7_Click()
    frmChekIMAGES.Show 1
    
'    On Error GoTo MiErr
'
'    Dim ThisFolder As String, TamTapa As Double
'    Dim MasDe20Kb As String, MasDe300Kb As String
'    Dim TapasGrandes As Long, TapasMuyGrandes As Long
'    TapasGrandes = 0: TapasMuyGrandes = 0
'
'    For a = 0 To UBound(MATRIZ_DISCOS)
'        ThisFolder = txtInLista(MATRIZ_DISCOS(a), 0, ",")
'        tERR.Anotar "ackq", a, UBound(MATRIZ_DISCOS), ThisFolder
'        If FSO.FileExists(ThisFolder + "\tapa.jpg") Then
'            TamTapa = FileLen(ThisFolder + "\tapa.jpg")
'            TamTapa = Round(TamTapa / 1024, 2)
'            tERR.Anotar "ackq", TamTapa
'            If TamTapa > 200 Then
'                MasDe300Kb = MasDe300Kb + FSO.GetBaseName(ThisFolder) + vbCrLf
'                TapasMuyGrandes = TapasMuyGrandes + 1
'            Else 'AQUI YA ESTA EN K_BYTES !!!!
'                If TamTapa > TamanoTapaPermitido Then
'                    MasDe20Kb = MasDe20Kb + FSO.GetBaseName(ThisFolder) + vbCrLf
'                    TapasGrandes = TapasGrandes + 1
'                End If
'            End If
'        End If
'    Next
'
'    Dim MsGrandes As String
'    Dim MsMuyGrandes As String
'
'    Select Case IDIOMA
'        Case "Español"
'            MSG = "Hay " + CStr(TapasGrandes) + _
'                " tapas de mas de " + CStr(TamanoTapaPermitido) + " Kb. Estas son:" + vbCrLf + MasDe20Kb
'        Case "English"
'        Case "Francois"
'        Case "Italiano"
'    End Select
'    tERR.Anotar "acks", TapasGrandes, TapasMuyGrandes
'    If TapasGrandes > 0 Then MsGrandes = MSG
'
'    Select Case IDIOMA
'        Case "Español"
'            MSG = "Hay " + CStr(TapasMuyGrandes) + _
'                " tapas de mas de 200 Kb. Estas son:" + vbCrLf + MasDe300Kb
'        Case "English"
'        Case "Francois"
'        Case "Italiano"
'    End Select
'    If TapasMuyGrandes > 0 Then MsMuyGrandes = MSG
'
'    Dim MSGfinal As String
'    MSGfinal = MsGrandes + vbCrLf + MsMuyGrandes
'    If Len(MSGfinal) < 5 Then
'        Select Case IDIOMA
'            Case "Español"
'                MsgBox "Todas las portadas de los discos tiene tamaños correctos"
'            Case "English"
'            Case "Francois"
'            Case "Italiano"
'        End Select
'    Else
'        MsgBox MSGfinal
'    End If
'
'    Exit Sub
'MiErr:
'    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjs"
'    Resume Next
End Sub

Private Sub Command8_Click()
    On Error GoTo MiErr

    Dim NewFolder As String
    
    NewFolder = InputBox(TR.Trad("Ingrese el nuevo nombre para el disco%99%"), _
                    TR.Trad("Nuevo nombre%98%Valor predeterminado del nombre " + _
                    "de un nuevo disco%99%"))
                    
    tERR.Anotar "ackt", NewFolder
    If NewFolder <> "" Then
        'primero corregir y acomodar el ranking para no perder los votos!!
        '?????????¿¿¿¿¿¿¿¿¿¿¿¿
        Name lstCarpetas As AP + "discos\" + NewFolder
        lstCarpetasShow.List(lstCarpetas.ListIndex) = NewFolder
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjt"
    Resume Next
End Sub

Private Sub Command9_Click()
    On Error GoTo MiErr
    tERR.Anotar "acku"
    Dim tNewFolder As String
    
        
    CmdLg.DialogPrompt = "Buscar Origen de disco"
    CmdLg.DialogTitle = "Buscar Origen de Disco"
    
    CmdLg.InitDir = "K:" 'si lo pongo en C: algunas PCs ponen SOLO el C: ¿?
        'Si no pongo nada da error el ShowFolder. Entonces elegi esto
    
    'CmdLg.InitDir = AP 'si lo pongo en C: algunas PCs ponen SOLO el C: ¿?
        'Si no pongo nada da error el ShowFolder. Entonces elegi esto
        
    'CmdLg.flags = cdlOFNExplorer Or cdlCCFullOpen Or Folder_COMPUTER Or Folder_INCLUDEFILES
    
    CmdLg.ShowFolder
    
    If CmdLg.InitDir = "" Or CmdLg.InitDir = "C:\" Then Exit Sub
    
    tNewFolder = CmdLg.InitDir
    tERR.Anotar "ackv", tNewFolder
    lstOrigenes.AddItem tNewFolder
    lstOrigenes.ListIndex = lstOrigenes.ListCount - 1
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acju"
    Resume Next
End Sub

Private Sub Form_Activate()
    Label1(0) = TR.Trad("DISCOS%98%Titulo de la lista de discos%99%")
    Command1.Caption = TR.Trad("Eliminar discos elegidos%99%")
    cmdKillTapa.Caption = TR.Trad("Quitar foto%99%")
    Command4.Caption = TR.Trad("Nueva foto%99%")
    cmdKillTXT.Caption = TR.Trad("Quitar texto%99%")
    Command2.Caption = TR.Trad("Nuevo texto%99%")
    cmdKillArch.Caption = TR.Trad("Eliminar selecciones elegidas%99%")
    Command2.Caption = TR.Trad("Agregar Música%99%")
    Command6.Caption = TR.Trad("Generar estadistica de discos%99%")
    command7.Caption = TR.Trad("Revisar tamaño de las portadas%99%")
    Command3.Caption = TR.Trad("SALIR%99%")
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaCerrarSistema
            Unload Me
            YaCerrar3PM
        Case vbKeyF5
            Command3_Click
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo MiErr
    tERR.Anotar "ackw", KeyCode, Shift
    If KeyCode = TeclaNewFicha Then
        LTE 1
        VarCreditos CSng(TemasPorCredito)
    End If
    
    If KeyCode = TeclaNewFicha2 Then
        LTE 2
        VarCreditos CSng(CreditosBilletes)
    End If
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjv"
    Resume Next
End Sub

Private Sub Form_Load()
    Pintar_fBoton Me
 
    Traducir 'Agregado por el complemento traductor
    
    On Error GoTo MiErr
    tERR.Anotar "ackx"
    AjustarFRM Me, 12000, 9000
    
'    'mostrar los origenes
'    Dim Origenes As String
'    Origenes = LeerArch1Linea(GPF("origs"))
'
'    PartOrigenes = Split(Origenes, "*")
    'es publico !! ya se carga en el load
    
    lstOrigenes.Clear
    Dim AAA As Long
    For AAA = 0 To UBound(PartOrigenes)
        tERR.Anotar "acky", AA, PartOrigenes(AAA)
        lstOrigenes.AddItem PartOrigenes(AAA)
    Next AAA
    'siempre uno elegido!
    lstOrigenes.ListIndex = 0
    
    'si no es SL taparlo JAJAJAJA!!!
    If K.sabseee("3pm") <> Supsabseee Then
        lstOrigenes.Enabled = False
        Command9.Enabled = False 'boton agregar
        Command10.Enabled = False 'boton quitar
        Command14.Enabled = False 'boton up
        Command13.Enabled = False 'boton down
        command12.Enabled = False 'boton grabar
    End If
    'mostrar la lista de carpetas cargadas en 3PM
    tERR.Anotar "ackz"
    CargarCarpetas
    InfoDisco lblInfoDisco
    
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjw"
    Resume Next
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub lstCarpetas_Click()
    On Error GoTo MiErr
    
    lstTEMAS.Clear
    'mostrar los temas de esta carpeta solo si hay una sola carpeta elegida
    tERR.Anotar "acla", lstCarpetas.SelCount
    If lstCarpetas.SelCount > 1 Then
        lstTEMAS.AddItem TR.Trad("No hay vista disponible%98%Cada vez que se " + _
            "hace un click en un disco se ven sus canciones y su portada" + _
            " esto aparece cuando hay seleecionados dos o más al mismo tiempo%99%")
        
        lstTEMAS.AddItem TR.Trad("Multiples carpetas seleccionadas%99%")
        
        lstTEMAS.Enabled = False
        
    Else
        lstTEMAS.Enabled = True
        ReDim Preserve MTXfiles(0)
        'OM- lista el contenido multimedia de una carpeta en el visor de la ventana agregar/quitar musica
        MTXfiles = ObtenerArchMM(lstCarpetas)
        tERR.Anotar "aclb", UBound(MTXfiles)
        If UBound(MTXfiles) = 0 Then
            lstTEMAS.AddItem TR.Trad("No hay temas multimedia en este directorio%99%")
            lstTEMAS.Enabled = False
        Else
            For A = 1 To UBound(MTXfiles)
                tERR.Anotar "aclc", A, UBound(MTXfiles)
                lstTEMAS.AddItem txtInLista(MTXfiles(A), 1, "#")
                lstTEMAS.Enabled = True
            Next
        End If
    End If
    cmdKillArch.Enabled = lstTEMAS.Enabled
    'mostrar la tapa si la tiene
    Dim TapaArch As String
    TapaArch = lstCarpetas + "\tapa.jpg"
    tERR.Anotar "acld", TapaArch
    If fso.FileExists(TapaArch) Then
        TapaCD.Picture = LoadPicture(TapaArch)
        cmdKillTapa.Enabled = True
        Dim TamTapa As Double
        tERR.Anotar "acle", TamTapa
        TamTapa = FileLen(TapaArch)
        TamTapa = Round(TamTapa / 1024, 2)
        lblKB = CStr(TamTapa) + " KB"
    Else
        tERR.Anotar "aclf"
        'ver si es superlicencia y usa otra tapa predeterminada
        IMF = GetTpPred
                
        TapaCD.Picture = LoadPicture(IMF)
        cmdKillTapa.Enabled = False
        lblKB = "8 KB"
    End If
    'mostrar el texto si existe
    Dim DataTXT As String
    DataTXT = lstCarpetas + "\data.txt"
    tERR.Anotar "aclg", DataTXT
    If fso.FileExists(DataTXT) Then
        Dim TE2 As TextStream
        Set TE2 = fso.OpenTextFile(DataTXT, ForReading, False)
        txtDataTXT = TE2.ReadAll
        cmdKillTXT.Enabled = True
    Else
        txtDataTXT = "NO EXISTE%98%Se refiere a un archivo de texto " + _
            "con datos adicionales del disco%99%"
        cmdKillTXT.Enabled = False
    End If
    'ver el ranking y analizar las estadísticas del disco
    'leer los discos y ver uno por uno (de mayor a menor) que temas se escucharon
    If fso.FileExists(GPF("rd3_444")) = False Then
        lstEstadisticas.AddItem TR.Trad("No hay informacion%99%")
        Exit Sub
    End If
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    lstEstadisticas.Clear
    Dim ThisLine As String, ThisDISCO As String, ThisTEMA As String, CantLisen As Long
    Dim TotLisen As Long, nPuesto As Long
    TotLisen = 0: nPuesto = 0
    tERR.Anotar "aclh"
    Do While Not TE.AtEndOfStream
        nPuesto = nPuesto + 1
        ThisLine = TE.ReadLine
        tERR.Anotar "acli", nPuesto, ThisLine
        'cantidad,pathfull,toShow(disco-tema),carpeta
        ThisDISCO = txtInLista(ThisLine, 3, ",")
        If ThisDISCO = lstCarpetasShow Then
            CantLisen = Val(txtInLista(ThisLine, 0, ","))
            ThisTEMA = txtInLista(fso.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
            lstEstadisticas.AddItem ThisTEMA + ": " + CStr(CantLisen) + " (RANK: " + CStr(nPuesto) + ")"
            TotLisen = TotLisen + CantLisen
            tERR.Anotar "aclj", CantLisen, ThisTEMA
        End If
    Loop
    If TotLisen = 0 Then
        lstEstadisticas.AddItem TR.Trad("Este disco no se ha escuchado todavia%99%")
    Else
        lstEstadisticas.AddItem ""
        lstEstadisticas.AddItem "   " + TR.Trad("TOTAL: %99%") + CStr(TotLisen)
    End If
    TE.Close
    tERR.Anotar "aclk"
    Exit Sub
    
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjx"
    Resume Next
End Sub

Private Sub lstCarpetasShow_Click()
    lstCarpetas.ListIndex = lstCarpetasShow.ListIndex
End Sub

Public Function ContarLisen(Carpeta As String) As Long
    On Error GoTo MiErr
    
    Set TE = fso.OpenTextFile(GPF("rd3_444"), ForReading, False)
    Dim ThisLine As String, ThisDISCO As String, ThisTEMA As String, CantLisen As Long
    Dim TotLisen As Long, nPuesto As Long
    TotLisen = 0: nPuesto = 0
    tERR.Anotar "acll"
    Do While Not TE.AtEndOfStream
        nPuesto = nPuesto + 1
        ThisLine = TE.ReadLine
        tERR.Anotar "aclm", nPuesto, ThisLine
        'cantidad,pathfull,toShow(disco-tema),carpeta
        ThisDISCO = txtInLista(ThisLine, 3, ",")
        If ThisDISCO = Carpeta Then
            CantLisen = Val(txtInLista(ThisLine, 0, ","))
            ThisTEMA = txtInLista(fso.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
            TotLisen = TotLisen + CantLisen
            tERR.Anotar "acln", CantLisen, ThisTEMA
        End If
    Loop
    TE.Close
    ContarLisen = TotLisen
    tERR.Anotar "aclo"
    Exit Function
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjy"
    Resume Next
End Function

Public Function STRceros(n As Variant, Cifras As Integer) As String
    'n es el numero y cifras es la cantidad final de cifras del str terminado
    'devuelve ej : para 232,6 = 000232 para 1902,12 = 000000001902
    'complaeta con ceroas adelante
    ' si n es mas lasgo que cifras devuelve el valor n sin ningun cero adelante
    Dim STRn As String
    STRn = Trim(CStr(n))
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
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    Command14.Caption = TR.Trad("subir%98%poner mas alto en el orden de " + _
        "prioridades a un elemento de una lista ordenada%99%")
    Command11.Caption = TR.Trad("ayuda%99%")
    Command9.Caption = TR.Trad("agregar%99%")
    Command13.Caption = TR.Trad("bajar%98%poner mas bajo en el orden de " + _
        "prioridades a un elemento de una lista ordenada%99%")
    command12.Caption = TR.Trad("grabar%99%")
    Command10.Caption = TR.Trad("quitar%99%")
    Command8.Caption = TR.Trad("Renombrar Disco%99%")
    command7.Caption = TR.Trad("Revisar tamaño de las portadas%99%")
    Command6.Caption = TR.Trad("Generar estadistica de discos%99%")
    Command2.Caption = TR.Trad("Agregar Música%99%")
    Command5.Caption = TR.Trad("Nuevo texto%99%")
    cmdKillTXT.Caption = TR.Trad("Quitar texto%99%")
    Command4.Caption = TR.Trad("Nueva foto%99%")
    cmdKillTapa.Caption = TR.Trad("Quitar foto%99%")
    Command3.Caption = TR.Trad("SALIR%99%")
    cmdKillArch.Caption = TR.Trad("Eliminar temas elegidos%99%")
    Command1.Caption = TR.Trad("Eliminar discos elegidos%99%")
    Command15.Caption = TR.Trad("Grabar estadisticas a disco%99%")
    Label1(5).Caption = TR.Trad("La música se cargará en este orden.%99%")
    Label1(4).Caption = TR.Trad("Ubicaciones de música (solo SL)%99%")
    Label1(3).Caption = TR.Trad("Utilize esta planilla para " + _
        "administrar sus discos. Puede modificar las imagenes y textos " + _
        "de cada uno. Utilice las estadísticas para borrar aquellos temas " + _
        "o discos que no se escuchen. Esto le permitirá utilizar mejor el " + _
        "espacio de su disco sin perder aquella musica que sus usarios " + _
        "estan escuchando%99%")
    Label1(2).Caption = TR.Trad("Estadisticas del disco elegido. Cantidad de " + _
        "veces que se escucho el disco%99%")
    lblInfoDisco.Caption = TR.Trad("Informacion del disco%99%")
    Label1(1).Caption = TR.Trad("Canciones del disco elegido%99%")
    Label1(0).Caption = TR.Trad("DISCOS%99%")
End Sub

Private Sub tbrFrame1_Click()

End Sub

Private Sub Picture1_Click()

End Sub
