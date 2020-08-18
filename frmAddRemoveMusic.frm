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
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "subir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ayuda"
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
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton Command9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "agregar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   30
      Width           =   900
   End
   Begin VB.CommandButton Command13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "bajar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   270
      Width           =   900
   End
   Begin VB.CommandButton Command12 
      BackColor       =   &H00C0C0C0&
      Caption         =   "grabar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   270
      Width           =   900
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "quitar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   270
      Width           =   900
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Rename Disco"
      Height          =   555
      Left            =   3270
      TabIndex        =   22
      Top             =   5760
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
      Height          =   7440
      IntegralHeight  =   0   'False
      Left            =   7200
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   30
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
      Left            =   1890
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7980
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
      Left            =   390
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7980
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
      Left            =   5610
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4230
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
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4230
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
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Text            =   "frmAddRemoveMusic.frx":0000
      Top             =   7140
      Width           =   4095
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
      Height          =   4500
      IntegralHeight  =   0   'False
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   8
      Top             =   2190
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
      Left            =   75
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6720
      Width           =   4035
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
      Height          =   3030
      IntegralHeight  =   0   'False
      ItemData        =   "frmAddRemoveMusic.frx":0006
      Left            =   4170
      List            =   "frmAddRemoveMusic.frx":0046
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   4860
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
      Height          =   300
      IntegralHeight  =   0   'False
      Left            =   60
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   8370
      Width           =   7095
   End
   Begin VB.ListBox lstOrigenes 
      BackColor       =   &H00000080&
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
      Height          =   1380
      IntegralHeight  =   0   'False
      ItemData        =   "frmAddRemoveMusic.frx":0091
      Left            =   90
      List            =   "frmAddRemoveMusic.frx":00AA
      TabIndex        =   23
      Top             =   540
      Width           =   7035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "La musica se cargara en este orden"
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
      Height          =   225
      Index           =   5
      Left            =   150
      TabIndex        =   31
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
      ForeColor       =   &H00C0FFFF&
      Height          =   225
      Index           =   4
      Left            =   150
      TabIndex        =   24
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
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6000
      TabIndex        =   20
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAddRemoveMusic.frx":00C3
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
      Height          =   1905
      Left            =   4260
      Stretch         =   -1  'True
      Top             =   2070
      Width           =   2850
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
      Left            =   4140
      TabIndex        =   6
      Top             =   4650
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
    If lstTEMAS.SelCount = 0 Then
        Select Case IDIOMA
            Case "Español"
                MsgBox "No hay archivos seleccionados"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        Exit Sub
    End If
    Select Case IDIOMA
        Case "Español"
            MSG = "Esta seguro que desea eliminar los archivos elegidos?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
        
    If MsgBox(MSG, vbQuestion + vbYesNo) = vbYes Then
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
        Select Case IDIOMA
            Case "Español"
                MsgBox "Los archivos se eliminaron correctamente"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
    End If
    InfoDisco lblInfoDisco
    Exit Sub
NOBORRA:
    Select Case IDIOMA
        Case "Español"
            MsgBox "No se ha podido borrar uno mas temas, compruebe " + _
                "que no esten abiertos. Error numero:" + CStr(Err.Number) + _
                " Descripcion interna: " + Err.Description
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
End Sub

Private Sub cmdKillTapa_Click()
    
    Select Case IDIOMA
        Case "Español"
            MSG = "¿Esta seguro que desea eliminar la imagen elegida?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select

    If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    'refrescar la imagen
    TapaCD.Picture = LoadPicture(SYSfolder + "f61.dlw")
    cmdKillTapa.Enabled = False
End Sub

Private Sub cmdKillTXT_Click()
    Select Case IDIOMA
        Case "Español"
            MSG = "¿Esta seguro que desea eliminar el texto elegido?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    FSO.DeleteFile lstCarpetas + "\data.txt", True
    cmdKillTXT.Enabled = False
    txtDataTXT = "NO EXISTE"
End Sub

Private Sub Command1_Click()
    'si borro una vez esta todo OK
    ' si entro por segunda vez ya matriz_discos tiene distintos indices
    'con respecto a las listas!!!!!!!!!!!!
    Select Case IDIOMA
        Case "Español"
            MSG = "Esta seguro que desea eliminar las carpetas elegidas?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    If MsgBox(MSG, vbQuestion + vbYesNo) = vbYes Then
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
                    Select Case IDIOMA
                        Case "Español"
                            MsgBox "No se puede borrar la carpeta del ranking"
                        Case "English"
                        Case "Francois"
                        Case "Italiano"
                    End Select
                Else
                    FSO.DeleteFolder CarpSel, True
                    WriteTBRLog "Se borro la carpeta " + CarpSel, True
                End If
            End If
        Next
        Select Case IDIOMA
            Case "Español"
                MsgBox "Las carpetas se eliminaron correctamente"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        'actualizar todo
        CargarCarpetas
        InfoDisco lblInfoDisco
    End If
    Exit Sub
NOBORRA:

    Select Case IDIOMA
        Case "Español"
            MsgBox "No se ha podido borrar uno mas carpetas, compruebe " + _
                "que no esten abiertas. Error numero:" + CStr(Err.Number) + _
                " Descripcion interna: " + Err.Description
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select

End Sub

Public Sub CargarCarpetas()
    lstCarpetas.Clear 'si no se duplican todos
    lstCarpetasShow.Clear
    On Error GoTo ErrCarp
    For A = 0 To UBound(MATRIZ_DISCOS)
        Dim ThisFolder As String, TamTapa As Double
        ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
        'ver si existen o se borraron
        If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
            lstCarpetas.AddItem txtInLista(MATRIZ_DISCOS(A), 0, ",")
            lstCarpetasShow.AddItem txtInLista(MATRIZ_DISCOS(A), 1, ",")
        End If
    Next
    lstCarpetasShow.Selected(0) = True
    
ErrCarp:
    WriteTBRLog "LINEA: " + LineaError + vbCrLf + Err.Description + " N°: " + Str(Err.Number), True
    Resume Next

End Sub

Private Sub Command10_Click()
    'ver sio hay mas de uno!!!
    If lstOrigenes.ListCount = 1 Then
        MsgBox "No se puede quitar el ultimo oprigen de discos. Debe haber uno!"
        Exit Sub
    End If
    'borrar
    lstOrigenes.RemoveItem lstOrigenes.ListIndex
    lstOrigenes.ListIndex = 0
End Sub

Private Sub Command11_Click()
    MsgBox "3PM en modo SuperLicencia permite la carga de musica desde diferentes ubicaciones." + vbCrLf + _
        "Esto permite utilizar mas de un disco rígido o diferentes particiones de un mismo disco." + vbCrLf + _
        "Además 3PM Kabalin 6.5 permite dar un mejor orden a la música y videos expuestos al público." + vbCrLf + _
        "Al iniciar el sistema 3PM leerá en primer lugar la musica de la primera ubicación hasta llegar a la ultima." + vbCrLf + _
        "De esta forma se podrá separar la musica en diferentes ritmos (rock, pop, folcklore, popular, etc)." + vbCrLf + _
        "Dentro de cada ubicación los discos estarán ordenados alfabéticamente permitiendo al usuario final un acceso mas sencillo a la musica buscada", vbQuestion, "UBICACIONES DE MUSICA"
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
    
    EscribirArch1Linea SYSfolder + "oddtb.jut", TMPs
    MsgBox "Los cambios se han grabado satisfactoriamente"
End Sub

Private Sub Command13_Click()
    'poner el que esta elegido abajo -BAJAR-
    'si es el ultimo (X) o no hay elegido (-1) se va
    If lstOrigenes.ListIndex = (lstOrigenes.ListCount - 1) Then Exit Sub
    If lstOrigenes.ListIndex = -1 Then Exit Sub
    
    Dim TMPlst As String, NumSube As Long
    NumSube = lstOrigenes.ListIndex
    TMPlst = lstOrigenes.List(NumSube + 1)
    'el anterior se transforma en el que sube
    lstOrigenes.List(NumSube + 1) = lstOrigenes
    lstOrigenes.List(NumSube) = TMPlst
    lstOrigenes.ListIndex = NumSube + 1
End Sub

Private Sub Command14_Click()
    'poner el que esta elegido arriba -SUBIR-
    'si es el primero (0) o no hay elegido (-1) se va
    If lstOrigenes.ListIndex < 1 Then Exit Sub
    
    Dim TMPlst As String, NumSube As Long
    NumSube = lstOrigenes.ListIndex
    TMPlst = lstOrigenes.List(NumSube - 1)
    'el anterior se transforma en el que sube
    lstOrigenes.List(NumSube - 1) = lstOrigenes
    lstOrigenes.List(NumSube) = TMPlst
    lstOrigenes.ListIndex = NumSube - 1
    
End Sub

Private Sub Command2_Click()
    frmAddMusic.Show 1
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    Select Case IDIOMA
        Case "Español"
            CmdLg.DialogTitle = "Eliga la nueva imagen del disco elegido"
            CmdLg.Filter = "Imagen JPG o GIF|*.jpg;*.gif"
        Case "English"
            CmdLg.DialogTitle = "Eliga la nueva imagen del disco elegido"
            CmdLg.Filter = "Image JPG o GIF|*.jpg;*.gif"
        Case "Francois"
        Case "Italiano"
    End Select
    
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    'ver el tamaño!!!
    Dim TamTapa As Double
    TamTapa = FileLen(CmdLg.FileName)
    TamTapa = Round(TamTapa / 1024, 2) 'son KB
    If TamTapa > 20 Then
    
        Select Case IDIOMA
            Case "Español"
                MSG = "tbrSoft recomienda imagenes no mayores a 8 KB, esta " + _
                    "imagen tine " + CStr(TamTapa) + " KB. ¿Esta seguro que desea usarla?" + _
                    " Puede afectar el rendimiento del equipo!"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    
        If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    If TamTapa > 200 Then
        Select Case IDIOMA
            Case "Español"
                MSG = "Imagen demasiado pesada. Despues de la advertencia, ¿aun desea usarla?" + _
                    " Puede afectar el rendimiento del equipo!!!"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    
        If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    End If
    
    If cmdKillTapa.Enabled Then
    
        Select Case IDIOMA
            Case "Español"
                MSG = "¿Esta seguro que desea reemplazar la imagen existente?"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    
        If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
        FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    End If
    FSO.CopyFile CmdLg.FileName, lstCarpetas + "\tapa.jpg", True
    TapaCD.Picture = LoadPicture(lstCarpetas + "\tapa.jpg")
End Sub

Private Sub Command5_Click()
    
    'ojo ARTIME que aqui el programa se guia por lo que esta escrito en el boton
    'o sea fijate en mayusculas y minusculas!!!
    
    Select Case IDIOMA
        Case "Español"
            MSG = "Confirmar"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    
    If Command5.Caption = MSG Then
        txtDataTXT.BackColor = &HE0E0E0   'color original
        txtDataTXT.Locked = True
        
        Select Case IDIOMA
            Case "Español"
                Command5.Caption = "Nuevo texto"
            Case "English"
                Command5.Caption = "New Text"
            Case "Francois"
            Case "Italiano"
        End Select
        
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
        
        Select Case IDIOMA
            Case "Español"
                Command5.Caption = "Confirmar"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    End If
    
End Sub

Private Sub Command6_Click()

    'ojo ARTIME que aqui el programa se guia por lo que esta escrito en el boton
    'o sea fijate en mayusculas y minusculas!!!
    Select Case IDIOMA
        Case "Español"
            MSG = "Generar estadistica de discos"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    If Command6.Caption = MSG Then
        lstTODO.Clear
        
        Select Case IDIOMA
            Case "Español"
                lstTODO.AddItem " RANK DISCOS (primero los menos escuhados)"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    
        'pasar por todos los discos y medir de cada uno las escuchadas
        Dim ThisFolder As String
        Dim Carp As String
        For A = 0 To UBound(MATRIZ_DISCOS)
            Carp = txtInLista(MATRIZ_DISCOS(A), 1, ",")
            ThisFolder = txtInLista(MATRIZ_DISCOS(A), 0, ",")
            'ver si existen o se borraron
            If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
                lstTODO.AddItem STRceros(ContarLisen(Carp), 4) + ": " + Carp
            End If
            lstTODO.Visible = True
            
            Select Case IDIOMA
                Case "Español"
                    Command6.Caption = "Quitar estadistica de discos"
                Case "English"
                Case "Francois"
                Case "Italiano"
            End Select
        Next
    Else
        lstTODO.Visible = False
        
        Select Case IDIOMA
            Case "Español"
                Command6.Caption = "Generar estadistica de discos"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    End If
End Sub

Private Sub Command7_Click()
    Dim ThisFolder As String, TamTapa As Double
    Dim MasDe20Kb As String, MasDe300Kb As String
    Dim TapasGrandes As Long, TapasMuyGrandes As Long
    TapasGrandes = 0: TapasMuyGrandes = 0
        
    For A = 0 To UBound(MATRIZ_DISCOS)
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
    
    Select Case IDIOMA
        Case "Español"
            MSG = "Hay " + CStr(TapasGrandes) + _
                " tapas de mas de 20 Kb. Estas son:" + vbCrLf + MasDe20Kb
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    
    If TapasGrandes > 0 Then MsGrandes = MSG
    
    Select Case IDIOMA
        Case "Español"
            MSG = "Hay " + CStr(TapasMuyGrandes) + _
                " tapas de mas de 200 Kb. Estas son:" + vbCrLf + MasDe300Kb
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    If TapasMuyGrandes > 0 Then MsMuyGrandes = MSG
    
    Dim MSGfinal As String
    MSGfinal = MsGrandes + vbCrLf + MsMuyGrandes
    If Len(MSGfinal) < 5 Then
        Select Case IDIOMA
            Case "Español"
                MsgBox "Todas las portadas de los discos tiene tamaños correctos"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    Else
        MsgBox MSGfinal
    End If
End Sub

Private Sub Command8_Click()
    Dim NewFolder As String
    
    Select Case IDIOMA
        Case "Español"
            NewFolder = InputBox("Ingrese el nuevo nombre para el disco", _
                "Nuevo nombre")
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    
    If NewFolder <> "" Then
        'primero corregir y acomodar el ranking para no perder los votos!!
        '?????????¿¿¿¿¿¿¿¿¿¿¿¿
        
        Name lstCarpetas As AP + "discos\" + NewFolder
        lstCarpetasShow.List(lstCarpetas.ListIndex) = NewFolder
    End If
End Sub

Private Sub Command9_Click()
    Dim tNewFolder As String
    CmdLg.InitDir = "C:\"
    CmdLg.ShowFolder
    
    If CmdLg.InitDir = "" Or CmdLg.InitDir = "C:\" Then Exit Sub
    
    tNewFolder = CmdLg.InitDir
    
    lstOrigenes.AddItem tNewFolder
    lstOrigenes.ListIndex = lstOrigenes.ListCount - 1
End Sub

Private Sub Form_Activate()
    Select Case IDIOMA
        Case "Español"
            Label1(0) = "DISCOS"
            Command1.Caption = "Eliminar discos elegidos"
            cmdKillTapa.Caption = "Quitar foto"
            Command4.Caption = "Nueva foto"
            cmdKillTXT.Caption = "Quitar texto"
            Command2.Caption = "Nuevo texto"
            cmdKillArch.Caption = "Eliminar temas elegidos"
            Command2.Caption = "Agregar Música"
            Command6.Caption = "Generar estadistica de discos"
            Command7.Caption = "Revisar tamaño de las portadas"
            Command3.Caption = "SALIR"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
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
            ShowCredits
            
            'grabar credito para validar
            'creditosValidar ya se cargo en load de frmindex
            CreditosValidar = CreditosValidar + TemasPorCredito
            EscribirArch1Linea SYSfolder + "radilav.cfg", CStr(CreditosValidar)
            
        Else
            'apagar el fichero electronico
            OnOffCAPS vbKeyScrollLock, False
        End If
    End If
End Sub

Private Sub Form_Load()
    AjustarFRM Me, 12000
    
    'mostrar los origenes
    Dim Origenes As String
    Origenes = LeerArch1Linea(SYSfolder + "oddtb.jut")
    Dim PartOrigenes() As String
    PartOrigenes = Split(Origenes, "*")
    
    lstOrigenes.Clear
    Dim AAA As Long
    For AAA = 0 To UBound(PartOrigenes)
        lstOrigenes.AddItem PartOrigenes(AAA)
    Next AAA
    'siempre uno elegido!
    lstOrigenes.ListIndex = 0
    
    'si no es SL taparlo JAJAJAJA!!!
    If K.LICENCIA <> HSuperLicencia Then
        lstOrigenes.Enabled = False
        Command9.Enabled = False 'boton agregar
        Command10.Enabled = False 'boton quitar
        Command14.Enabled = False 'boton up
        Command13.Enabled = False 'boton down
        Command12.Enabled = False 'boton grabar
    End If
    'mostrar la lista de carpetas cargadas en 3PM
    CargarCarpetas
    InfoDisco lblInfoDisco
End Sub

Private Sub lstCarpetas_Click()
    lstTEMAS.Clear
    'mostrar los temas de esta carpeta solo si hay una sola carpeta elegida
    If lstCarpetas.SelCount > 1 Then
        Select Case IDIOMA
            Case "Español"
                lstTEMAS.AddItem "No hay vista disponible"
                lstTEMAS.AddItem "Multiples carpetas seleccionadas"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        lstTEMAS.Enabled = False
        
    Else
        lstTEMAS.Enabled = True
        ReDim Preserve MTXfiles(0)
        MTXfiles = ObtenerArchMM(lstCarpetas)
        If UBound(MTXfiles) = 0 Then
            Select Case IDIOMA
                Case "Español"
                    lstTEMAS.AddItem "No hay temas multimedia en esta carpeta"
                Case "English"
                Case "Francois"
                Case "Italiano"
            End Select
            
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
        TapaCD.Picture = LoadPicture(SYSfolder + "f61.dlw")
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
        Select Case IDIOMA
            Case "Español"
                txtDataTXT = "NO EXISTE"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        cmdKillTXT.Enabled = False
    End If
    'ver el ranking y analizar las estadísticas del disco
    'leer los discos y ver uno por uno (de mayor a menor) que temas se escucharon
    If FSO.FileExists(AP + "ranking.tbr") = False Then
        Select Case IDIOMA
            Case "Español"
                lstEstadisticas.AddItem "No hay informacion"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
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
        Select Case IDIOMA
            Case "Español"
                lstEstadisticas.AddItem "Este disco no se ha escuchado todavia"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
    Else
        lstEstadisticas.AddItem ""
        Select Case IDIOMA
            Case "Español"
                lstEstadisticas.AddItem "   TOTAL: " + CStr(TotLisen)
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
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

