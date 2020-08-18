VERSION 5.00
Object = "{50A3BB94-15CB-4B69-8357-277A2E4CC1AB}#2.0#0"; "tbrJPG.ocx"
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmChekIMAGES 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reparación y búsqueda de imágenes"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   4125
      Left            =   2670
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   240
      Width           =   5235
   End
   Begin tbrJPG.tbrToJPG TJ 
      Height          =   450
      Left            =   1110
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   794
   End
   Begin tbrFaroButton.fBoton XxBoton1 
      Height          =   585
      Left            =   210
      TabIndex        =   0
      Top             =   2340
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Revisar Imagenes"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin tbrFaroButton.fBoton XxBoton2 
      Height          =   585
      Left            =   210
      TabIndex        =   3
      Top             =   3720
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   1032
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "Salir"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   5452834
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmChekIMAGES.frx":0000
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
      Height          =   2025
      Left            =   180
      TabIndex        =   4
      Top             =   390
      Width           =   2205
   End
End
Attribute VB_Name = "frmChekIMAGES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MyLog(T As String)
    If Len(Text1.tExt) > 300 Then Text1.tExt = Right(Text1.tExt, 300)
    Text1.tExt = Text1.tExt + CStr(Timer) + "  " + T + vbCrLf
    Text1.SelStart = Len(Text1.tExt) - 1
End Sub

Private Sub Form_Resize()
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
End Sub

Private Sub XxBoton1_Click()
    
    On Local Error GoTo MiErr
    
    Dim J As Long
    'entrar a cada disco y revisar que el tamaño este ok
    Dim ArchTapa As String
    Dim CarpTapa As String
    
    Dim Reparados As Long
    Reparados = 0
    For J = 1 To UBound(MATRIZ_DISCOS)
        'ver si la imagen existe y esta dentro de rangos permitidos
        Me.Caption = "Reparación y búsqueda de imágenes " + CStr(J)
        Me.Refresh
        CarpTapa = txtInLista(MATRIZ_DISCOS(J), 0, ",")
        tERR.Anotar "acnc66", CarpTapa
        If CarpTapa <> "_RANK_" Then
            If Right(CarpTapa, 1) <> "\" Then CarpTapa = CarpTapa + "\"
            
            TR.SetVars CarpTapa
            MyLog TR.Trad("Comenzando con %01% %98%La variable es el path de un " + _
                "directorio donde se buscan imagenes para corregir%99%")
            
            If fso.FileExists(CarpTapa + "tapa.jpg") = False Then
                'buscar otros en esta carpeta!
                TR.SetVars "Tapa.jpg"
                MyLog TR.Trad("No existe como %01%%98%La variable es 'tapa.jpg' como" + _
                    "no quiero que se cambie con la traducción lo pongo como variable%99%")
                
                ArchTapa = getBestImg(CarpTapa)
                
                If ArchTapa = "" Then
                    MyLog TR.Trad("Además no hay ninguna imagen! Se abandona%99%")
                    GoTo sig
                Else
                    TR.SetVars ArchTapa
                    MyLog TR.Trad("Se usará %01%%98%La variable es el nombre del " + _
                        "archivo de imagén que se eligio como mejor opción%99%")
                End If
                
                'sigue no mas con la mejor imagen que haya traido
                
            Else
                'si esta la uso sin fijarme si hay mas
                ArchTapa = CarpTapa + "tapa.jpg"
                MyLog TR.Trad("Si existe %99%") + ArchTapa
            End If
        Else 'el tapa es el rank!
            MyLog TR.Trad("Es la imagen del ranking! No se modifica%99%")
            GoTo sig
        End If
        
        TR.SetVars FileLen(ArchTapa), TamanoTapaPermitido * 1024
        MyLog TR.Trad("La imagen pesa %01% Bytes y se permite %02% bytes %99%")
              
        If FileLen(ArchTapa) > TamanoTapaPermitido * 1024 Then
            tERR.Anotar "acnc67"
            fso.CopyFile ArchTapa, CarpTapa + "ex-tapa.jpg", True
            
            Dim res As Long, sCompres As Long
            res = FileLen(ArchTapa) 'para que entre al do!!
            sCompres = 90
            'achicar en compresion hasta que quede dentro de niveles permitidos
            Do While res > (TamanoTapaPermitido * 1024)
                
                res = TJ.BMPtoJPG_CalcularBites(ArchTapa, sCompres)
                TR.SetVars sCompres, res
                MyLog TR.Trad("Pruebo compresion %01% queda en %02% bytes%99%")
                
                sCompres = sCompres - 10
                
                If sCompres < 10 Then
                    'ahora pruebo con el tamaño
                    Dim WiHe As Long
                    sCompres = 90
                    WiHe = 150
                    Do While res > (TamanoTapaPermitido * 1024)
                        res = TJ.BMPtoJPG_CalcularBites(ArchTapa, sCompres, WiHe, WiHe)
                        TR.SetVars sCompres, WiHe, res
                        MyLog TR.Trad("Pruebo compresion %01% y tamaño en %02% " + _
                            "queda en %03% bytes%99%")
                            
                        sCompres = sCompres - 10
                        WiHe = WiHe - 10
                    Loop
                    
                    'listo, lo mejor que se pueda o alguno que dio ok
                    TJ.BMPtoJPG CarpTapa + "ex-tapa.jpg", CarpTapa + "tapa.jpg", sCompres, WiHe, WiHe
                    
                    TR.SetVars res
                    MyLog TR.Trad("DEFINIDO EN %01% bytes%99%")
                    
                    GoTo sig
                End If
            Loop
            
            TR.SetVars res
            MyLog TR.Trad("DEFINIDO EN %01% bytes%99%")
            'listo, lo mejor que se pueda o alguno que dio ok
            TJ.BMPtoJPG CarpTapa + "ex-tapa.jpg", CarpTapa + "tapa.jpg", sCompres
            Reparados = Reparados + 1
            
            GoTo sig
        Else
            'si era otro lo dejo con el nombre que debe ser!
            If LCase(ArchTapa) <> CarpTapa + "tapa.jpg" Then
                fso.CopyFile ArchTapa, CarpTapa + "tapa.jpg", True
                Reparados = Reparados + 1
            End If
            MyLog TR.Trad("El tamaño esta ok!%99%")
        End If
sig:
        MyLog "--------- SIG ----------"
    
    Next J
    
    MyLog TR.Trad("Proceso Terminado%99%")
    TR.SetVars Reparados
    MsgBox TR.Trad("Proceso Terminado%97%Se corrigieron %01% imágenes%99%")
    Exit Sub
    
MiErr:
    MyLog "ERROR N°" + CStr(Err.Number) + " - " + Err.Description
    If Err.Number = 481 Then 'ivalid picture!!
        MyLog "Imagen No valida! en " + CarpTapa
        tERR.AppendSinHist "Imagen No valida! en " + CarpTapa
        GoTo sig
    Else
        tERR.AppendLog tERR.ErrToTXT(Err), Me.Name + ".acjs"
        Resume Next
    End If
    
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
            If InStr(fso.GetBaseName(res(J)), "tapa") > 0 Then ThisPtos = ThisPtos + 500
            If InStr(fso.GetBaseName(res(J)), "frente") > 0 Then ThisPtos = ThisPtos + 500
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
'-------Agregado por el complemento traductor------------
Private Sub Form_Load()
    Pintar_fBoton Me
    Traducir
End Sub
'-------Agregado por el complemento traductor------------
Private Sub Traducir()
    XxBoton1.Caption = TR.Trad("Revisar Imagenes%99%")
    XxBoton2.Caption = TR.Trad("Salir%99%")
    Label1.Caption = TR.Trad("Desde aquí podrá revisar tamaño y disponibilidad de las portadas. " + _
        "De esta forma el sistema automáticamente elegirá y ajustara " + _
        "a las necesidades de 3PM.%99%")
End Sub
