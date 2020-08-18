VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarritoDelete 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton ButOK 
      Height          =   705
      Left            =   7830
      TabIndex        =   5
      Top             =   2940
      Visible         =   0   'False
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1244
      fFColor         =   16777215
      fBColor         =   14737632
      fCapt           =   "OK"
      fEnabled        =   -1  'True
      fFontN          =   "Verdana"
      fFontS          =   10
      fECol           =   5452834
   End
   Begin tbr3pm.tbrFullProc SW 
      Height          =   375
      Left            =   4020
      TabIndex        =   4
      Top             =   5550
      Visible         =   0   'False
      Width           =   705
      _ExtentX        =   1244
      _ExtentY        =   661
   End
   Begin VB.ListBox lstPTHS 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1110
      IntegralHeight  =   0   'False
      Left            =   270
      TabIndex        =   0
      Top             =   2820
      Width           =   8685
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "x x x x xx  x  x x x  x xx x xx xx x x x x x x x x xxxxxx x x x xx x x x x x x  x x x x x  xxXXXX XXx X xX xx x XX"
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
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   5610
      Width           =   9435
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x x x x xx  x  x x x  x xx x xx xx x x x x x x x x xxxxxx x x x xx x x x x x x  x x x x x  xxXXXX XXx X xX xx x XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   180
      TabIndex        =   3
      Top             =   1980
      Width           =   8955
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x x x x xx  x  x x x  x xx x xx xx x x x x x x x x xxxxxx x x x xx x x x x x x  x x x x x  xxXXXX XXx X xX xx x XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   645
      Left            =   150
      TabIndex        =   2
      Top             =   1170
      Width           =   8955
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x x x x xx  x  x x x  x xx x xx xx x x x x x x x x xxxxxx x x x xx x x x x x x  x x x x x  xxXXXX XXx X xX xx x XX"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   975
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   8955
   End
End
Attribute VB_Name = "frmCarritoDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private S() As String
Private lstELIMINA() As String

Private EspNecesario As Single
Private EspDisponible As Single

Private Sub TermineDeBorrar()
    SW.ShowWait "Leyendo dispositivo ..."
    frmCarritoInDev.ShowDEV UB.GetLetterUSB(-1)
    Unload Me
End Sub

Private Sub PressOk()

    On Local Error GoTo MER

    If lstPTHS.ListIndex = -1 Then
        tERR.Anotar "dabf2"
        Exit Sub
    End If
    
    tERR.Anotar "dabf", lstPTHS.ListIndex, lstPTHS.List(lstPTHS.ListIndex)
    
    'no quiere borrar nada
    If lstPTHS.ListIndex = 1 Then
        SW.ShowWait ""
        Unload Me
        Exit Sub
    End If
    
    'dice que termino
    If lstPTHS.ListIndex = 0 Then
        'ver si de verdad termino
        If EspDisponible > EspNecesario Then
            
            TermineDeBorrar
            Exit Sub
        End If
    End If
    
    If lstPTHS.ListIndex > 1 Then
        'eliminar la carpeta ....
        Dim F As String
        'Dim SP() As String
        'SP = Split(lstPTHS, "|")
        'F = Trim(SP(0)) 'carpeta a elimiar
        F = lstELIMINA(lstPTHS.ListIndex)
        F = Mid(F, 1, Len(F) - 1)
        
        'hay una serie de cuestiones a analizar. Las carpetas que necesitan los celulares
        ' o que tienen como valores predeterminados
        'por ejemplo: Music, Musica, Video, VideoClips
        'en caso de quieran borrar esas solo deberíua borrar su contenido
        'o lo que es mas facil boorrar todo y crearla al final
        'simpre y cuando no este bloquedo para eliminar esas carpetas (hasta ahora no me paso)
        
        Dim NoDeleteThisFolder(7) As String
        NoDeleteThisFolder(0) = "Music"
        NoDeleteThisFolder(1) = "Musica"
        NoDeleteThisFolder(2) = "Video"
        NoDeleteThisFolder(3) = "VideoClip"
        NoDeleteThisFolder(4) = "VideoClips"
        NoDeleteThisFolder(5) = "Images"
        NoDeleteThisFolder(6) = "Imágenes"
        NoDeleteThisFolder(7) = "Imagenes"
        
        Dim SP() As String
        SP = Split(F, "\")
        
        Dim Z As Long, ReGenerar As String 'indica si hay que escribir algo y que es!
        ReGenerar = ""
        
        'solo si se borra esa carpeta
        If UBound(SP) = 1 Then
            For Z = 0 To UBound(NoDeleteThisFolder)
                If LCase(SP(1)) = LCase(NoDeleteThisFolder(Z)) Then
                    ReGenerar = SP(0) + "\" + SP(1)
                    Exit For
                End If
            Next Z
        End If
        
        tERR.Anotar "dabg", F
        SW.ShowWait "Eliminando " + vbCrLf + F
        tERR.Anotar "dabh"
        fso.DeleteFolder F, True
        SW.ShowWait ""
        
        'ver si hay que regrabar algo
        If ReGenerar <> "" Then
            fso.CreateFolder ReGenerar
        End If
        
        'volver a contar el espacio libre
        UB.RefreshValues -1
        UpdateDisp
        
        'puede seguir borrando
        'mostrar todo de nuevo segun lo que ya se elimino
        ShowDEV UB.GetLetterUSB(-1)
    End If

    Exit Sub
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC7"
    
    If Err.Number = 70 Then
        'PERMISO DENEGADO
        'ES SOLO LECTURA!
        SW.ShowWait "No se puede eliminar. El dispositivo puede estar bloqueado para escritura", 4000
    Else
        'ni idea que es!!!
        Resume Next
    End If

End Sub

Private Sub UpdateDisp()
    tERR.Anotar "dabj", EspNecesario, EspDisponible
    
    tERR.Anotar "dabk"
    EspNecesario = Carrito.GetTotalMB
    EspDisponible = UB.GetFreeMB(-1) 'el que sea que este elegido
    
    If EspNecesario > EspDisponible Then
        Label2.Caption = "Necesita liberar " + CStr(Round(EspNecesario - EspDisponible, 2)) + " MBs" + vbCrLf + _
            "Elija uno o más directorios y elimínelos con la tecla OK o de carrito." + vbCrLf + _
            "Para finalizar elija alguna de las dos primeras opciones."
        Label2.ForeColor = &HFF&
    Else
        Label2.Caption = "Ya hay suficiente espacio disponible, puede seguir " + _
            "eliminando si lo desea o presional la segunda opción para continuar " + _
            "con su compra"
        Label2.ForeColor = &HFF00&
        lstPTHS.ListIndex = 0
    End If
    
    Label3.Caption = "Espacio necesario " + CStr(EspNecesario) + " MB " + _
        "/ Espacio Disponible " + CStr(EspDisponible) + " MB"
    
End Sub

Private Sub ButOK_Click()
    PressOk
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    tERR.Anotar "dabf", KeyCode, Chr(KeyCode)
    Select Case KeyCode
        Case TeclaDER: MoveLS 1 ' SendKeys "{DOWN}"
        Case TeclaPagAt: MoveLS -1 ' SendKeys "+{UP}"
        Case TeclaIZQ: MoveLS -1 ' SendKeys "+{UP}"
        Case TeclaPagAd: MoveLS 1 '  SendKeys "{DOWN}"
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    tERR.Anotar "dabf", KeyCode, Chr(KeyCode)
    Select Case KeyCode
        Case TeclaESC: Unload Me
        Case TeclaCarrito: PressOk
        Case TeclaOK: PressOk
        Case Else 'puto enter numerico
            If KeyCode = 13 And TeclaOK = 108 Then PressOk
    End Select
End Sub

Private Sub MoveLS(mov As Long)
    If mov > 0 Then
        If lstPTHS.ListIndex < lstPTHS.ListCount - mov Then
            lstPTHS.ListIndex = lstPTHS.ListIndex + mov
        End If
    End If
    
    If mov < 0 Then
        If lstPTHS.ListIndex >= -mov Then
            lstPTHS.ListIndex = lstPTHS.ListIndex + mov 'mov es negativo!
        End If
    End If
End Sub

Private Sub Form_Load()
           
    EsSaving = True 'para que no se lance ni el protector ni temas al azar!
    
    'buscar todas las carpetas con multimedia
    'mostrarlas con su peso en MB
    
    'cuidado de no dejarle borrar cosas que puedan ser útiles.
    
    Label1.Caption = "Al no haber suficiente espacio en su dispositivo deberá " + _
        "elegir uno o más directorios para eliminar"
        
    ButOK.Visible = MostrarTouch

End Sub

Public Sub ShowDEV(PTh As String)
    
    On Local Error GoTo MER
    UpdateDisp
    SW.ShowWait "Leyendo dispositivo ..."

    Dim pt As New tbrPaths.clsPATHS
    'hacer la base con tamaños incluidos!!!!
    pt.LeerTodo PTh + ":\", False, False
    pt.UpdateFolderSize 'pone los tamaños en las carpatas!
    
    tERR.Anotar "dabl"
    S = pt.GetLista
    
    Dim s2() As String
    lstPTHS.Visible = False
    lstPTHS.Clear
    
    lstPTHS.AddItem "Listo, he terminado de borrar. Grabar mi compra ahora"
    lstPTHS.AddItem "Salir. No quiero borrar nada"
    
    ReDim lstELIMINA(1) 'asi empieza del 2 como en el list box!
    
    For H = 1 To UBound(S)
        tERR.Anotar "dabm", UBound(S)
        'asegurarse que sea carpeta con multimedia
        SW.ShowWait "Leyendo dispositivo ...", 0, (H / UBound(S)) * 100
        Dim TamFol As Single
        
        'si elige solo carpetas con multimedia puede evitar borrar elementos que puedan ser
        'importantes
        's2 = ObtenerArchMM(S(H))
        
        'If UBound(s2) > 1 Then
        Dim Tam As String
        If Right(S(H), 1) = "\" Then
            'ver cuanto pesa tambien
            TamFol = pt.GetTamanoDirectorioMB(S(H))
            tERR.Anotar "dabn", TamFol, S(H)
            If Len(S(H)) > 70 Then S(H) = Left(S(H), 65)
            'nuevo. los chip de celular pueden tener muchas carpetas vacias
            If TamFol > 5 Then
                ReDim Preserve lstELIMINA(UBound(lstELIMINA) + 1)
                lstELIMINA(UBound(lstELIMINA)) = S(H)
                
                Tam = CStr(TamFol)
                lstPTHS.AddItem Space(8 - Len(Tam)) + Tam + " MB - " + fso.GetBaseName(S(H))
            End If
        End If
    Next H
    
    SW.ShowWait ""
    
    lstPTHS.ListIndex = 0
    lstPTHS.Visible = True
    
    Exit Sub
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC8"
    Resume Next
End Sub

Private Sub Form_Resize()
    Label1.Left = 60
    Label1.Width = Me.Width - 120
    
    Label2.Left = Label1.Left
    Label2.Width = Label1.Width
    
    Label3.Left = Label1.Left
    Label3.Width = Label1.Width
    
    lstPTHS.Left = 200
    lstPTHS.Width = Me.Width - 400
    
    lstPTHS.Top = Label3.Top + Label3.Height + 120
    lstPTHS.Height = Me.Height - lstPTHS.Top - Label4.Height - 120
    
    Label4.Top = lstPTHS.Top + lstPTHS.Height
    Label4.Width = Label3.Width
    Label4.Left = Label3.Left
    
    'SOLO UNA VEZ DESPUES DEL LOAD
    '*****************************
    ShowDEV UB.GetLetterUSB(-1)
    '*****************************
    
    ButOK.Left = lstPTHS.Width - ButOK.Width - 160
    ButOK.Top = lstPTHS.Top + 60
    
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15
    
End Sub

Private Sub lstPTHS_Click()
    Label4.Caption = "ELEGIDO: " + lstELIMINA(lstPTHS.ListIndex)
End Sub
