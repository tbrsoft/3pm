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
      Height          =   7140
      IntegralHeight  =   0   'False
      Left            =   7230
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
      Height          =   1455
      Left            =   10500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7200
      Width           =   1365
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
   Begin VB.CommandButton Command15 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Grabar estadisticas a disco"
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
      TabIndex        =   32
      Top             =   7200
      Width           =   3255
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
      Height          =   1005
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
    On Error GoTo MiErr
    
    tERR.Anotar "acjy"
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
                FileSel = txtInLista(MTXfiles(AA + 1), 0, "#")
                FSO.DeleteFile FileSel, True
                tERR.Anotar "acka", AA, TotSel, FileSel
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
    tERR.Anotar "ackb"
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
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acji"
    Resume Next
End Sub

Private Sub cmdKillTapa_Click()
    On Error GoTo MiErr
    
    Select Case IDIOMA
        Case "Español"
            MSG = "¿Esta seguro que desea eliminar la imagen elegida?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select

    If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    tERR.Anotar "ackc", lstCarpetas
    'refrescar la imagen
    'ver si es superlicencia y usa otra tapa predeterminada
    If K.LICENCIA = HSuperLicencia Then
        If FSO.FileExists(GPF("tddp322")) Then
            imF = GPF("tddp322")
        Else
            imF = ExtraData.GetImagePath("tapapredeterminada")
        End If
    Else
        imF = ExtraData.GetImagePath("tapapredeterminada")
    End If
    
    TapaCD.Picture = LoadPicture(imF)
    cmdKillTapa.Enabled = False
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjj"
    Resume Next
End Sub

Private Sub cmdKillTXT_Click()
    On Error GoTo MiErr
    Select Case IDIOMA
        Case "Español"
            MSG = "¿Esta seguro que desea eliminar el texto elegido?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    If MsgBox(MSG, vbYesNo + vbCritical) = vbNo Then Exit Sub
    tERR.Anotar "ackd", lstCarpetas
    FSO.DeleteFile lstCarpetas + "\data.txt", True
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
    Select Case IDIOMA
        Case "Español"
            MSG = "Esta seguro que desea eliminar las carpetas elegidas?"
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    tERR.Anotar "acke"
    If MsgBox(MSG, vbQuestion + vbYesNo) = vbYes Then
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
                    Select Case IDIOMA
                        Case "Español"
                            MsgBox "No se puede borrar la carpeta del ranking"
                        Case "English"
                        Case "Francois"
                        Case "Italiano"
                    End Select
                Else
                    FSO.DeleteFolder CarpSel, True
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

    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjk"
    Resume Next
End Sub

Public Sub CargarCarpetas()
    On Error GoTo MiErr
    lstCarpetas.Clear 'si no se duplican todos
    For a = 0 To UBound(MATRIZ_DISCOS)
        Dim ThisFolder As String, TamTapa As Double
        ThisFolder = txtInLista(MATRIZ_DISCOS(a), 0, ",")
        tERR.Anotar "ackh", a, UBound(MATRIZ_DISCOS), ThisFolder
        'ver si existen o se borraron
        'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
        If FSO.FolderExists(ThisFolder) Then
            lstCarpetas.AddItem txtInLista(MATRIZ_DISCOS(a), 0, ",")
            lstCarpetasShow.AddItem txtInLista(MATRIZ_DISCOS(a), 1, ",")
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
    For a = 0 To lstOrigenes.ListCount - 1
        'al ultimo no pongo asterisco para que al dividir no quede uno vacio al ultimo!
        If a < lstOrigenes.ListCount - 1 Then
            TMPs = TMPs + lstOrigenes.List(a) + "*"
        Else
            TMPs = TMPs + lstOrigenes.List(a)
        End If
    Next a
    EscribirArch1Linea GPF("origs"), TMPs
    MsgBox "Los cambios se han grabado satisfactoriamente"
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
    If FSO.FileExists("C:\STATS.TXT") Then FSO.DeleteFile "C:\STATS.TXT", True
    
    Dim TE121 As TextStream
    
    Set TE121 = FSO.OpenTextFile("C:\STATS.TXT", ForAppending, True)
    
    TE121.WriteLine "-------------------------"
    TE121.WriteLine "ESTADISTICAS SEGUN DISCOS"
    TE121.WriteLine "-------------------------"
    TE121.WriteLine "Cantidad Reproducciones: DISCO"
    Dim ThisFolder As String
    Dim Carp As String, a As Long
    lstTODO.Clear
    For a = 0 To UBound(MATRIZ_DISCOS)
        Carp = txtInLista(MATRIZ_DISCOS(a), 1, ",")
        ThisFolder = txtInLista(MATRIZ_DISCOS(a), 0, ",")
        tERR.Anotar "ackp2", a, UBound(MATRIZ_DISCOS), Carp, ThisFolder
        'ver si existen o se borraron
        'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
        If FSO.FolderExists(ThisFolder) Then
            lstTODO.AddItem STRceros(ContarLisen(Carp), 4) + ": " + Carp
        End If
    Next a
    
    For a = 0 To lstTODO.ListCount - 1
        TE121.WriteLine lstTODO.List(a)
    Next a
    
    'GRABAR RANK DE CANCION
    TE121.WriteLine "-------------------------"
    TE121.WriteLine "ESTADISTICAS SEGUN CANCIONES"
    TE121.WriteLine "-------------------------"
    TE121.WriteLine "Cantidad Reproducciones: CANCION"
        'grabar en un txt las canciones
        Dim TE120 As TextStream
        Set TE120 = FSO.OpenTextFile(GPF("rd3_444"), ForReading, True)
            Do While Not TE120.AtEndOfStream
                TE121.WriteLine TE120.ReadLine
            Loop
        TE120.Close
        Set TE120 = Nothing
    TE121.Close
    Set TE121 = Nothing
    MsgBox "Las estadisticas se han grabado sin problemas en C:\STATS.TXT"
End Sub

Private Sub Command2_Click()
    frmAddMusic.Show 1
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Command4_Click()
    On Error GoTo MiErr
    
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
    tERR.Anotar "ackk"
    CmdLg.ShowOpen
    If CmdLg.FileName = "" Then Exit Sub
    'ver el tamaño!!!
    Dim TamTapa As Double
    TamTapa = FileLen(CmdLg.FileName)
    TamTapa = Round(TamTapa / 1024, 2) 'son KB
    tERR.Anotar "ackl", TamTapa
    If TamTapa > TamanoTapaPermitido Then
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
        tERR.Anotar "ackm", lstCarpetas
        FSO.DeleteFile lstCarpetas + "\tapa.jpg", True
    End If
    tERR.Anotar "ackn", CmdLg.FileName
    FSO.CopyFile CmdLg.FileName, lstCarpetas + "\tapa.jpg", True
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
            tERR.Anotar "acko", txtDataTXT
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
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjq"
    Resume Next
End Sub

Private Sub Command6_Click()
    On Error GoTo MiErr
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
        For a = 0 To UBound(MATRIZ_DISCOS)
            Carp = txtInLista(MATRIZ_DISCOS(a), 1, ",")
            ThisFolder = txtInLista(MATRIZ_DISCOS(a), 0, ",")
            tERR.Anotar "ackp", a, UBound(MATRIZ_DISCOS), Carp, ThisFolder
            'ver si existen o se borraron
            'If FSO.FolderExists(ThisFolder) And ThisFolder <> AP + "discos\01- Los mas escuchados" Then
            If FSO.FolderExists(ThisFolder) Then
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
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjr"
    Resume Next
End Sub

Private Sub Command7_Click()
    On Error GoTo MiErr
    
    Dim ThisFolder As String, TamTapa As Double
    Dim MasDe20Kb As String, MasDe300Kb As String
    Dim TapasGrandes As Long, TapasMuyGrandes As Long
    TapasGrandes = 0: TapasMuyGrandes = 0
        
    For a = 0 To UBound(MATRIZ_DISCOS)
        ThisFolder = txtInLista(MATRIZ_DISCOS(a), 0, ",")
        tERR.Anotar "ackq", a, UBound(MATRIZ_DISCOS), ThisFolder
        If FSO.FileExists(ThisFolder + "\tapa.jpg") Then
            TamTapa = FileLen(ThisFolder + "\tapa.jpg")
            TamTapa = Round(TamTapa / 1024, 2)
            tERR.Anotar "ackq", TamTapa
            If TamTapa > 200 Then
                MasDe300Kb = MasDe300Kb + FSO.GetBaseName(ThisFolder) + vbCrLf
                TapasMuyGrandes = TapasMuyGrandes + 1
            Else 'AQUI YA ESTA EN K_BYTES !!!!
                If TamTapa > TamanoTapaPermitido Then
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
                " tapas de mas de " + CStr(TamanoTapaPermitido) + " Kb. Estas son:" + vbCrLf + MasDe20Kb
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
    tERR.Anotar "acks", TapasGrandes, TapasMuyGrandes
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
    
    Exit Sub
MiErr:
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjs"
    Resume Next
End Sub

Private Sub Command8_Click()
    On Error GoTo MiErr

    Dim NewFolder As String
    
    Select Case IDIOMA
        Case "Español"
            NewFolder = InputBox("Ingrese el nuevo nombre para el disco", _
                "Nuevo nombre")
        Case "English"
        Case "Francois"
        Case "Italiano"
    End Select
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
    CmdLg.InitDir = "K:" 'si lo pongo en C: algunas PCs ponen SOLO el C: ¿?
        'Si no pongo nada da error el ShowFolder. Entonces elegi esto
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
            YaCerrar3PM
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
    On Error GoTo MiErr
    tERR.Anotar "ackx"
    AjustarFRM Me, 12000
    
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
    If K.LICENCIA <> HSuperLicencia Then
        lstOrigenes.Enabled = False
        Command9.Enabled = False 'boton agregar
        Command10.Enabled = False 'boton quitar
        Command14.Enabled = False 'boton up
        Command13.Enabled = False 'boton down
        Command12.Enabled = False 'boton grabar
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

Private Sub lstCarpetas_Click()
    On Error GoTo MiErr
    
    lstTEMAS.Clear
    'mostrar los temas de esta carpeta solo si hay una sola carpeta elegida
    tERR.Anotar "acla", lstCarpetas.SelCount
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
        tERR.Anotar "aclb", UBound(MTXfiles)
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
            For a = 1 To UBound(MTXfiles)
                tERR.Anotar "aclc", a, UBound(MTXfiles)
                lstTEMAS.AddItem txtInLista(MTXfiles(a), 1, "#")
                lstTEMAS.Enabled = True
            Next
        End If
    End If
    cmdKillArch.Enabled = lstTEMAS.Enabled
    'mostrar la tapa si la tiene
    Dim TapaArch As String
    TapaArch = lstCarpetas + "\tapa.jpg"
    tERR.Anotar "acld", TapaArch
    If FSO.FileExists(TapaArch) Then
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
        If K.LICENCIA = HSuperLicencia Then
            If FSO.FileExists(GPF("tddp322")) Then
                imF = GPF("tddp322")
            Else
                imF = ExtraData.GetImagePath("tapapredeterminada")
            End If
        Else
            imF = ExtraData.GetImagePath("tapapredeterminada")
        End If
                
        TapaCD.Picture = LoadPicture(imF)
        cmdKillTapa.Enabled = False
        lblKB = "8 KB"
    End If
    'mostrar el texto si existe
    Dim DataTXT As String
    DataTXT = lstCarpetas + "\data.txt"
    tERR.Anotar "aclg", DataTXT
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
    If FSO.FileExists(GPF("rd3_444")) = False Then
        Select Case IDIOMA
            Case "Español"
                lstEstadisticas.AddItem "No hay informacion"
            Case "English"
            Case "Francois"
            Case "Italiano"
        End Select
        
        Exit Sub
    End If
    Set TE = FSO.OpenTextFile(GPF("rd3_444"), ForReading, False)
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
            ThisTEMA = txtInLista(FSO.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
            lstEstadisticas.AddItem ThisTEMA + ": " + CStr(CantLisen) + " (RANK: " + CStr(nPuesto) + ")"
            TotLisen = TotLisen + CantLisen
            tERR.Anotar "aclj", CantLisen, ThisTEMA
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
    
    Set TE = FSO.OpenTextFile(GPF("rd3_444"), ForReading, False)
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
            ThisTEMA = txtInLista(FSO.GetFileName(txtInLista(ThisLine, 1, ",")), 0, ".")
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
