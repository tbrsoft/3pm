VERSION 5.00
Begin VB.Form F1 
   BackColor       =   &H00000000&
   Caption         =   "Manejo de fallas"
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8910
   Icon            =   "F1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4320
      Left            =   150
      Picture         =   "F1.frx":0442
      ScaleHeight     =   4320
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   330
      Width           =   3960
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Intentar la mejor recuperación automatica"
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
      Height          =   435
      Index           =   2
      Left            =   4530
      MouseIcon       =   "F1.frx":ABBC
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   1230
      Width           =   3495
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Reparar eliminando archivos externos de 3PM (solo si la anterior no funciona)"
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
      Height          =   435
      Index           =   1
      Left            =   4650
      MouseIcon       =   "F1.frx":AEC6
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   1920
      Width           =   4155
   End
   Begin VB.Label LB 
      BackStyle       =   0  'Transparent
      Caption         =   "Generar un informe de errores para enviar a tbrSoft"
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
      Height          =   435
      Index           =   0
      Left            =   4140
      MouseIcon       =   "F1.frx":B1D0
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   150
      Width           =   3615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "tbrSoft Internacional 2001 - 2008"
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
      Left            =   60
      TabIndex        =   1
      Top             =   4710
      Width           =   8805
   End
End
Attribute VB_Name = "F1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fso As New Scripting.FileSystemObject
Dim JS As New tbrJUSE.clsJUSE
Dim BasePath As String

Dim SysFolder As String
Dim AP As String

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    BasePath = AP + "sf\"
    
    SysFolder = Fso.GetSpecialFolder(SystemFolder)
    If Right(SysFolder, 1) <> "\" Then SysFolder = SysFolder + "\"
    
    'BasePath = "D:\dev\3PM kundera 716226\sf\"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LB(0).ForeColor = vbWhite
    LB(1).ForeColor = vbWhite
    LB(2).ForeColor = vbWhite
End Sub

Private Sub LB_Click(Index As Integer)
    Select Case Index
        Case 0
            'empaquetar: el reg3pm.log + los archivos w15 + la configuración
            Dim F As String
            Randomize
            F = App.path + "\FullReg.JSA"
            DeleteFiles AP, "JSA"
            
            JS.Archivo = F
            
            CreateMyFile AP + "my.log", Get_LL
            
            'REGISTRO BASICO + REGISTRO DE MMPLAYER
            AddFiles App.path, "log"
            
            'ARCHIVOS W15
            AddFiles App.path, "w15"
            
            'CONFIGURACION DE 3PM
            AddFile BasePath + "marad.ona"
            
            'OTRAS COSAS INTERESANTES
            AddFile BasePath + "pindo.nga" 'lista de origenes de discos 'EX: sf+ "oddtb.jut"
            AddFile BasePath + "cd3.pm" 'Copia clave sf + "c2LK.dll"
            AddFile BasePath + "cccd3.pm" 'Copia clave sf + "c2LK.dll"
            AddFile BasePath + "cd4.pm" 'Archivo de licencia 3pm 7.0 (GENERADO)
            AddFile BasePath + "cd7.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 COREGIDO Y EN USO
            AddFile BasePath + "rdc.day" 'registro diario del contador sf + "daily.cfg"
            AddFile BasePath + "daliv.mp2" 'archivo con las claves para validar
            AddFile BasePath + "rempmon.45" 'archivos de registro de correcion de monedero
            
            'CRACK MALDITO
            AddFile BasePath + "jqs2323.dat" 'archivos de registro de correcion de monedero
            
            
            'LISTO UNIR TODO
            JS.Unir False
            
            Dim Cm As New CommonDialog
            Cm.DefaultExt = ".JSA"
            Cm.Filter = "Paquete de errores (*.JSA)|JSA"
            Cm.ShowSave
            Dim F2 As String
            F2 = Cm.FileName
            If F2 <> "" Then
                If LCase(Right(F2, 4)) <> ".JSA" Then F2 = F2 + ".JSA"
                Fso.MoveFile F, F2
                MsgBox "Se ha grabado el registro completo para enviar a tbrSoft en" + vbCrLf + F2
            Else
                MsgBox "Se ha grabado el registro completo para enviar a tbrSoft en" + vbCrLf + F
            End If
        
        Case 1 'borrar todo y volver a cero
            Reparar True
            MsgBox "Reparación finalizada"
            
        Case 2 'mejor recuperacion
            Reparar False
            MsgBox "Reparación finalizada"
            
    End Select
    
    Unload Me
    End
End Sub

Private Function AddFile(F As String) As Long
    If Fso.FileExists(F) Then
        AddFile = 0
        JS.AddFile F
    Else
        AddFile = 1
    End If
End Function

Private Function AddFiles(sFolder As String, Extension As String) As Long
    'devuleve la cantidad de agregados
    Dim Fl As Scripting.Folder
    Set Fl = Fso.GetFolder(sFolder)
    Dim F2 As Scripting.File
    Dim E1 As String, E2 As String
    For Each F2 In Fl.Files
        E1 = LCase(Right(F2.Name, Len(Extension)))
        E2 = LCase(Extension)
        If E1 = E2 Then
            AddFile F2.path
        End If
    Next
End Function

Private Function DeleteFiles(sFolder As String, Extension As String) As Long
    'devuleve la cantidad de agregados
    Dim F As Scripting.Folder
    Set F = Fso.GetFolder(sFolder)
    Dim F2 As Scripting.File
    For Each F2 In F.Files
        If LCase(Right(F2.Name, Len(Extension))) = LCase(Extension) Then
            F2.Delete True
        End If
    Next
End Function

Private Sub LB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Index <> 0 And LB(0).ForeColor <> vbWhite Then LB(0).ForeColor = vbWhite
    If Index <> 1 And LB(1).ForeColor <> vbWhite Then LB(1).ForeColor = vbWhite
    If Index <> 2 And LB(2).ForeColor <> vbWhite Then LB(2).ForeColor = vbWhite
    
    If LB(Index).ForeColor <> vbYellow Then LB(Index).ForeColor = vbYellow
End Sub

Private Function DoRepair(F1 As String, F2 As String, Optional KillAll As Boolean = False)
    'el 1 es el original, el 2 es la copia de seguridad si existe
    
    If KillAll Then
        If Fso.FileExists(F1) Then Fso.DeleteFile F1, True
        If Fso.FileExists(F2) Then Fso.DeleteFile F2, True
        Exit Function
    Else
        If Fso.FileExists(F1) = False Then
            If F2 <> "" Then
                If Fso.FileExists(F2) Then
                    Fso.CopyFile F2, F1, True
                Else
                    'no hacer nada el sistema regenerara el archivo original con sus valores predeterminados
                End If
            Else 'no hay definida copia de seguridad
                'no hacer nada el sistema regenerara el archivo original con sus valores predeterminados
            End If
        Else 'si existe el original
            If F2 <> "" Then
                If Fso.FileExists(F2) Then
                    Fso.DeleteFile F1, True
                    Fso.CopyFile F2, F1, True
                Else
                    Fso.DeleteFile F1, True
                    'no hacer nada el sistema regenerara el archivo original con sus valores predeterminados
                End If
            Else 'no hay definida copia de seguridad
                Fso.DeleteFile F1, True
                'no hacer nada el sistema regenerara el archivo original con sus valores predeterminados
            End If
        End If
    End If
End Function

Private Sub Reparar(BorrarTodo As Boolean)
    Dim FF As String 'cada archivo
    Dim FF2 As String 'copia de seguridad
    
    FF = BasePath + "pindo.nga" 'lista de origenes de discos 'EX: sf+ "oddtb.jut"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "kund.era" 'Creditos actuales para usar 'EX: AP + "creditos.tbr"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "marad.ona" 'sf + "3pmcfg.tbr"
    FF2 = BasePath + "eber.lud" 'copia de seguridad de la config sf + "autoSave3PM.cfg"
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "cpd.dor" 'Archivo con la clave de validacion sf +"codped.cfg"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "atak.e77" 'Codigos contados para validacion sf + "radilav.cfg"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c01" 'contadores de creditos
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c02" 'contadores de creditos
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c03" 'contadores de creditos
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c04" 'contadores de creditos
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c05" 'contadores de carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c06" 'contadores de carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c07" 'contadores de carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "chd.c08" 'contadores de carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "cd3.pm" 'Clave sf + "dciLib22.dll"
    FF2 = BasePath + "cccd3.pm" 'Copia clave sf + "c2LK.dll"
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin1.ppm" 'mLicencia3PMVtaMusica
    FF2 = BasePath + "plin2.ppm" 'mLicencia3PMVtaMusica BUP
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin3.ppm" 'mLicencia3PMOrigMusicaFTP
    FF2 = BasePath + "plin4.ppm" 'mLicencia3PMOrigMusicaFTP BUP
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin5.ppm" 'mLicencia3PMConfigOnline
    FF2 = BasePath + "plin6.ppm" 'mLicencia3PMConfigOnline BUP
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin7.ppm" 'mLicenciaCD001Kar
    FF2 = BasePath + "plin8.ppm" 'mLicenciaCD001Kar
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin9.ppm" 'mLicenciaCD002Kar
    FF2 = BasePath + "plin10.ppm" 'mLicenciaCD002Kar BUP
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin11.ppm" 'mLicenciaCD003Kar
    FF2 = BasePath + "plin12.ppm" 'mLicenciaCD003Kar
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin13.ppm" 'mLicenciaCD004Kar
    FF2 = BasePath + "plin14.ppm" 'mLicenciaCD004Kar
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin15.ppm" 'mLicenciaCD005Kar
    FF2 = BasePath + "plin16.ppm" 'mLicenciaCD005Kar
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "plin17.ppm" 'mLicenciaCD006Kar
    FF2 = BasePath + "plin18.ppm" 'mLicenciaCD006Kar
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "rdc.day" 'registro diario del contador sf + "daily.cfg"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "233.56b" '56 modificado por SL sf + "f5yaSL.nam"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "233.58b" '58 modificado por SL sf + "f7yaSL.nam"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "233.57b" '57 modif sf + "f6yaSL.nam"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "rempmon.45" 'reemplazo para señales de monedero sf + "teclaesp.fas"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "telc.not" 'texto en config No Tbr Wf + "SL\txtCFG.tbr"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "sequed.a32" 'Claves de uso desde afuera de la fonola
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "iis.l67" 'imagen del inicio de la SL Wf+ "SL\imgbig.tbr"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "iis.chu" 'imagen de fondo de las portadas
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "tddp.322" 'tapa predeterminada
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "tddp.323" 'tapa ranking
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "tslpri.112" 'txtSL principal Wf + "SL\txtIDX.tbr"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "rd3.444" 'AP + "ranking.tbr"
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "iit17.222" 'info sobre las imagenes de inicio de win
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "adpdp2.323" 'algo del protector
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "casc1.001" 'canciones a seguir cantando
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "daliv.mp2" 'archivo con las claves para validar
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "cart.987" 'contenido del carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
    FF = BasePath + "pcraor.mto" 'precios del carrito
    FF2 = ""
    DoRepair FF, FF2, BorrarTodo
    '************************************************
        
        'NO SE ELIMINANA
    '"cd4.pm" 'Archivo de licencia 3pm 7.0 (GENERADO)
    '"cd5.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 EN DESUSO
    '"cd6.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 (Copia SEG) EN DESUSO
    '"cd7.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 COREGIDO Y EN USO
    '"cd8.pm" 'Archivo RECIBIDO de licencia 3pm 7.0 (Copia SEG) COREGIDO Y EN USO
'    Case "pdis233"
'            TMP = "pdis.233" 'paquete de imagenes sf + "nev.man"
'        Case "extr233": TMP = "" 'sf sola para extraer el paquete de imagenes
'            Case "extr233_56": TMP = "f56.dlw" 'logo.sys
'            Case "extr233_57": TMP = "f57.dlw" 'logos.sys
'            Case "extr233_58": TMP = "f58.dlw" 'logow.sys
'Case "rempres44": TMP = "rempres.44" 'reemplazo del reserved cuando no hay sf + "razaGUID.dll"
'Case "ildw9m": TMP = "ildw9m.811" 'imagen logo.sys del win98/me wf + "img3pm\w\logo.sys"
'        Case "ildw9m2": TMP = "ildw9m.812" 'imagen logos.sys del win98/me Wf+ "img3pm\w\logos.sys"
'        Case "ildw9m3": TMP = "ildw9m.813" 'imagen logow.sys del win98/me Wf+ "img3pm\w\logow.sys"
'
'        Case "ild3pm": TMP = "ild3pm.811" 'imagen logo.sys del 3pm wf + "img3pm\3\logo.sys"
'        Case "ild3pm2": TMP = "ild3pm.812" 'imagen logos.sys del 3pm Wf+ "img3pm\3\logos.sys"
'        Case "ild3pm3": TMP = "ild3pm.813" 'imagen logow.sys del 3pm Wf+ "img3pm\3\logow.sys"
End Sub

Private Function Get_LL() As String
    'ver las versiones de todos las dlls
    Dim FLL As String, VLL As String
    Dim ACUM_LL As String
    
    FLL = SysFolder + "tbrerr.dll":            VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrreg.dll":            VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrtimer.dll":          VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrfocus.dll":          VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrplayer02.dll":       VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrSoftVumetro.dll":    VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrListaRep.dll":       VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrSKS3.dll":           VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrjuse.dll":           VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrnfo.dll":            VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrFullPak.dll":        VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "caescrypto.dll":        VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrcaescrypto.dll":     VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrprogress.dll":       VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrFaroButton.ocx":     VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrEncr.dll":           VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrPaths.dll":          VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrDrives.dll":         VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrFrame.ocx":          VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrGraficos.dll":       VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrALotOfPictures.dll": VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "ijl11.dll":             VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    FLL = SysFolder + "tbrJPG.ocx":            VLL = Fso.GetFileVersion(FLL): ACUM_LL = ACUM_LL + FLL + " " + VLL + vbCrLf
    
    Get_LL = ACUM_LL
End Function

Private Sub CreateMyFile(PT As String, TX As String)
    Dim TE As TextStream
    Set TE = Fso.CreateTextFile(PT, True)
        TE.Write TX
    TE.Close
End Sub


