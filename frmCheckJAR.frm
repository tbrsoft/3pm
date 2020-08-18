VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCheckJAR 
   BackColor       =   &H00000000&
   Caption         =   "Buscar descripcion JAR"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11205
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkREINIJAR 
      BackColor       =   &H00000000&
      Caption         =   "Ignorar descripciones existentes"
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
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   3660
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   510
      Width           =   10815
   End
   Begin tbrFaroButton.fBoton XxBoton1 
      Height          =   585
      Left            =   5190
      TabIndex        =   1
      Top             =   3690
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Comenzar Busqueda"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton2 
      Height          =   345
      Left            =   10140
      TabIndex        =   2
      Top             =   3660
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   609
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(no desmarque si desconoce su funcionamiento). Puede demorar."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   420
      TabIndex        =   5
      Top             =   3930
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Extraer información de los juegos y aplicaciones java contenido en los ficheros JAR"
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
      Height          =   465
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   8475
   End
End
Attribute VB_Name = "frmCheckJAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MyLog(T As String)
    Text1.tExt = Text1.tExt + CStr(Timer) + "  " + T + vbCrLf
    Text1.SelStart = Len(Text1.tExt) - 1
End Sub

Private Sub XxBoton1_Click()
    
    On Local Error GoTo MiErr
    
    Dim J As Long
    'entrar a cada disco y revisar que el tamaño este ok
    Dim ArchJaR As String
    Dim CarpJAR As String
    
    Dim Reparados As Long, EstabanOk As Long, Fallas As Long
    Reparados = 0: EstabanOk = 0: Fallas = 0
    
    For J = 1 To UBound(MATRIZ_DISCOS)
        'ver si la imagen existe y esta dentro de rangos permitidos
        
        CarpJAR = txtInLista(MATRIZ_DISCOS(J), 0, ",")
        
        If CarpJAR <> "_RANK_" Then
            If Right(CarpJAR, 1) <> "\" Then CarpJAR = CarpJAR + "\"
            
            TR.SetVars CarpJAR
            MyLog TR.Trad("Comenzando con %01% %98%La variable es el path de un " + _
                "directorio donde se buscan imagenes para corregir%99%")
            
            'BUSCAR TODOS LOS .jar
            ArchJaR = Dir(CarpJAR + "*.JAR")
            Do While ArchJaR <> ""
                Dim JJ As New tbrJAR, A As Long
                A = JJ.Open_Jar(CarpJAR + ArchJaR, chkREINIJAR.Value)
                If A = 0 Then
                    Reparados = Reparados + 1
                    MyLog "Reparado ok: " + ArchJaR
                End If
                
                If A = 3 Then
                    EstabanOk = EstabanOk + 1
                    MyLog "Verificado ok: " + ArchJaR
                End If
                
                If A = 2 Then
                    Fallas = Fallas + 1
                    MyLog "Falla!! : " + ArchJaR
                End If
                
                ArchJaR = Dir
            Loop
        End If
        
        Me.Caption = "Buscando descripciones " + CStr(Round(J / UBound(MATRIZ_DISCOS), 4) * 100) + " %"
    Next J
    
    MyLog TR.Trad("Proceso Terminado%99%")
    TR.SetVars Reparados
    MsgBox "Proceso Terminado" + _
        vbCrLf + _
        "Se agregaron " + CStr(Reparados) + " descripciones de ficheros JAR" + vbCrLf + _
        "Se verificaron " + CStr(EstabanOk) + " descripciones preexistentes" + vbCrLf + _
        "Fallaron " + CStr(Fallas) + vbCrLf + vbCrLf + _
        "Si hay 100% de fallas revise que el comando JAR este en su PATH o consulte a tbrSoft"
        
    Exit Sub
    
MiErr:
    MyLog "ERROR N°" + CStr(Err.Number) + " - " + Err.Description
    tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjs"
    Resume Next
End Sub

Private Function getBestImg(sFolder As String) As String
    'me dan una carpeta y busco entre todas las imagenes que tiene para suponer cual puede ser la tapa

    'devuelvo "" si no hay nada que sirva
    
    If Right(sFolder, 1) <> "\" Then sFolder = sFolder + "\"
    
    Dim ACH As String
    Dim EXTS(4) As String
    EXTS(0) = "jpg"
    EXTS(1) = "jpeg"
    EXTS(2) = "gif"
    EXTS(3) = "bmp"
    EXTS(4) = "tiff"
    
    Dim J As Long
    Dim res() As String 'lista de todas las imagenes disponibles
    ReDim res(0)
    For J = 0 To UBound(EXTS)
        ACH = Dir(sFolder + "*." + EXTS(J))
        Do While ACH <> ""
            ReDim Preserve res(UBound(res) + 1)
            res(UBound(res)) = sFolder + ACH
            ACH = Dir
        Loop
    Next J
    
    If UBound(res) = 0 Then
        getBestImg = ""
        Exit Function
    Else
        If UBound(res) = 1 Then 'si es solo una devuelvo esa!
            getBestImg = res(1)
            Exit Function
        End If
        
        'hay mas de una imagen
        Dim ThisPtos As Long 'puntos de la imagen actual
        Dim MaxPTOS As Long 'puntaje de cada imagen y maximo al mismo tiempo para elegir de una
        Dim IndMaxPtos As Long 'indice del elemento con mas puntos
        
        MaxPTOS = 0
        IndMaxPtos = 1 'predeterminada
        For J = 1 To UBound(res)
            ThisPtos = 0
            If InStr(FSO.GetBaseName(res(J)), "tapa") > 0 Then ThisPtos = ThisPtos + 500
            If InStr(FSO.GetBaseName(res(J)), "frente") > 0 Then ThisPtos = ThisPtos + 500
            ThisPtos = ThisPtos + CLng(FileLen(res(J)) / 1024) 'mas tamaño deberia ser mas calidad!
            
            If ThisPtos > MaxPTOS Then
                MaxPTOS = ThisPtos
                IndMaxPtos = J
            End If
            
        Next J
        
        'la mejor elegida
        getBestImg = res(IndMaxPtos)
        
    End If
    
End Function

Private Sub XxBoton2_Click()
    Unload Me
End Sub
