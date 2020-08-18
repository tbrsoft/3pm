VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form frmCarritoInDev 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9540
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
   ScaleHeight     =   4455
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin tbrFaroButton.fBoton butOK 
      Height          =   705
      Left            =   7800
      TabIndex        =   3
      Top             =   1500
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
      Height          =   2220
      IntegralHeight  =   0   'False
      Left            =   390
      TabIndex        =   0
      Top             =   1350
      Width           =   8685
   End
   Begin tbr3pm.tbrFullProc SW 
      Height          =   345
      Left            =   8430
      TabIndex        =   2
      Top             =   4020
      Visible         =   0   'False
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   609
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
      Height          =   1065
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8955
   End
End
Attribute VB_Name = "frmCarritoInDev"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim S() As String 'para usar el itemdata

Public Sub ShowDEV(PTh As String)
        
    On Local Error GoTo MER
    
    tERR.Anotar "dabo", PTh
    Dim pt As New tbrPaths.clsPATHS
    pt.LeerTodo PTh + ":\", True, False
    
    tERR.Anotar "dabp"
    S = pt.GetLista
    S(0) = PTh + ":\3PMusic\"
    tERR.Anotar "dabq", S(0)
    'XXXXXX
    'detectar si es un celular y poner el predeterminado donde corresponda
    'xxxxxx
    lstPTHS.AddItem " Predeterminado y recomendado: " + PTh + ":\3PMusic"
    
    For H = 1 To UBound(S)
        lstPTHS.AddItem " Guardar en: " + S(H)
    Next H

    lstPTHS.ListIndex = 0

    Me.Show 1
    
    Exit Sub
MER:
    tERR.AppendLog tERR.ErrToTXT(Err), "cpCC6"
    Resume Next
End Sub

Public Sub Save(I As Long)

    Dim EE As Long 'cantidad de veces que entro en error
    EE = 0
    
    On Local Error GoTo ErrCopy

    'el parametro es el numero de orden de la lista S de paths de destino posible
    
    tERR.Anotar "dabs"
    lstPTHS.Visible = False
    SW.ShowWait "Iniciando copia" ', 1500
    
    'ver los precios para ir descontando segun corresponda
    Dim DEST As String
    
    'a veces no elige el usuario va todo automatico
    'para esto uso el indice -1
    If I = -1 Then
        'si entro asi es por que es el unico dispositivo
        ReDim S(0)
        S(0) = UB.GetLetterUSB(1) + ":\3PMusic\"
        tERR.Anotar "dabt", S(0)
        DEST = S(0)
        'DEST = S(0)
    Else
        DEST = S(I)
    End If
    
    'si eligió la predeterminada puede ser que no exista
    If I <= 0 Then
        If fso.FolderExists(S(0)) = False Then fso.CreateFolder (S(0))
    End If
        
    SW.ShowWait "Creando directorios..."
    Dim H As Long
    'primero paso por las selecciones para ver si eligio un disco completo, de esta forma
    'se lo grabo en la carpeta que eligio y no todo suelto que queda feo
    tERR.Anotar "dabu", Carrito.GetFileCant
    For H = 1 To Carrito.GetFileCant
        'ver si es una cancion o una carpeta
        Dim T As String
        If Right(Carrito.GetElement(H), 1) = "\" Then
            'ES UNA CARPETA QUE ELIGIO !!
            T = DEST + fso.GetBaseName(Carrito.GetElement(H))
            tERR.Anotar "dabv", T
            SW.ShowWait "Creando directorio " + fso.GetBaseName(Carrito.GetElement(H))
            If Right(T, 1) <> "\" Then T = T + "\"
            If fso.FolderExists(T) = False Then
                tERR.Anotar "dabw"
                fso.CreateFolder T
            End If
        End If
    Next H
    
    'copiar archivos, la estructura ya esta
    Dim InFolder As String
    SW.ShowWait "Copiando ..."
    
    'revisar si puede !!!
    Randomize
    Dim mxGra As Long
    mxGra = Int(Rnd * 5) + 5
    If Carrito.GetFileCantFull > mxGra Then
        Dim RDS As TypeLic
        RDS = K.sabseee("mLicencia3PMVtaMusica")
        If RDS < DMinima Then
            SW.ShowWait TR.Trad("Sin Licencia de carro de compras!%99%"), 3500
            SW.ShowWait ""
            Unload Me
            Exit Sub
        End If
    End If
    
    'medir la velocidad
    Dim Copiado As Single 'cantidad copiada
    Dim sTimeCopyINI As Single 'tiempo en que la copio
    Dim sTimeCopy As Single 'tiempo en que la copio
    Copiado = 0
    sTimeCopyINI = Timer
    
    Dim MBxSec As Single
    Dim Falta As Single 'segundos que faltan
    Dim FaltaTXT As String
    Dim totCart As Long
    totCart = Carrito.GetTotalMB
    
    For H = 1 To Carrito.GetFileCantFull
        InFolder = fso.GetBaseName(fso.GetParentFolderName(Carrito.GetElementFull(H)))
        tERR.Anotar "dabx3", InFolder
        'EN LOS CELULARES O PENDRIVES PUEDE APARECER EL ERROR
        '-2147024784
        
        If InFolder = "" Then GoTo SIG444
        
        If fso.FolderExists(DEST + InFolder) Then
            tERR.Anotar "daby3", Carrito.GetElementFull(H), DEST + InFolder + "\"
            fso.CopyFile Carrito.GetElementFull(H), DEST + InFolder + "\", True
        Else
            tERR.Anotar "dabz"
            fso.CopyFile Carrito.GetElementFull(H), DEST, True
        End If
        
        Copiado = Copiado + (FileLen(Carrito.GetElementFull(H)) / 1048576)
        sTimeCopy = Timer - sTimeCopyINI
        
        'descontar el credito correspondiente a los que grabo
        'XXXXXXXXXXXXXX
        'No es un numero entero.... quilombo parecido al de las canciones
        'para nmo hacer lio saco todo lo que hay que sacar si se copio el primero ok
        
        'no lo saco al final por que si no se van a avivar y sacar el pendrive antes de
        'terminar y les va a costar cero!
        If H = 1 Then
            VarCreditos -Carrito.CalculateTotalPrice
            'sumo al contador de creditos de carrito lo que se gasto
            SumarContadorCarrito Carrito.CalculateTotalPrice
            'indicar cuanta plata entro en esta fonola en concepto de compra de música
            Dim YU As Long, DTaa As String
            DTaa = CStr(Year(Date)) + STRceros(Month(Date), 2) + STRceros(Day(Date), 2) + STRceros(Hour(time), 2) + STRceros(Minute(time), 2)
            
            'grabar un registro de todo lo que se compro para control.
            Dim PrecioCU As Single 'precio de cada cancion
            PrecioCU = (Carrito.CalculateTotalPrice * (PrecioBase / TemasPorCredito))
            PrecioCU = Round(PrecioCU / Carrito.GetFileCantFull, 2)
            For YU = 1 To Carrito.GetFileCantFull
                'tERR.Anotar "A198|B" + Carrito.GetElementFull(YU)
                'grabar en un registro de aca
                dwqu "U" + Carrito.GetElementFull(YU) + "*" + CStr(PrecioCU), dwQU_See, DTaa
            Next
            
        End If
        
        MBxSec = Round(Copiado / sTimeCopy, 6)
        Falta = CLng(CSng((totCart - Copiado) / MBxSec))
        
        If Falta > 59 Then
            Dim M As Long
            M = (Falta \ 60)
            Falta = Falta - (M * 60)
            If Falta < 10 Then
                FaltaTXT = CStr(M) + ":0" + CStr(Falta)
            Else
                FaltaTXT = CStr(M) + ":" + CStr(Falta)
            End If
            
        Else
            If Falta = 1 Then
                s2 = ""
            Else
                s2 = "s"
            End If
            If Falta < 10 Then
                FaltaTXT = "0:0" + CStr(Falta)
            Else
                FaltaTXT = "0:" + CStr(Falta)
            End If
        End If
        
        SW.ShowWait "Copiando " + vbCrLf + _
            fso.GetBaseName(Carrito.GetElementFull(H)) + " (" + _
            fso.GetExtensionName(Carrito.GetElementFull(H)) + ")" + vbCrLf + _
            "(" + CStr(Round(MBxSec, 2)) + " MB/S falta: " + FaltaTXT + ")", _
            0, ((H / Carrito.GetFileCantFull) * 100)
        
        Carrito.CleanFileSoloMarca H  'por las dudas que se corte por falla y no copie ni
        'descuente de nuevo
        
SIG444:
    Next H
    
    'SACAR DEL CARRITO los ya copaidos
    'Carrito.CleanMarcados YA SE VACIA ....
    
    'vaciarlo!
    tERR.Anotar "taca"
    Carrito.ClearCart
    SW.ShowWait "Proceso terminado con exito", 3300
    
    Unload Me
    Exit Sub
    
ErrCopy:

    tERR.Anotar "ASQQ", EE
    tERR.AppendLog tERR.ErrToTXT(Err), "CopingUSB"
    
    If Err.Number = -2147024784 Then
        '-2147024784 (80070070)
        'error en el metodo copyfile del objeto iFileSystem3
        SW.ShowWait "El dispositivo aparenta tener espacio suficiente " + _
            "pero ha fallado la copia." + vbCrLf + _
            "Pruebe luego de una defragmentación o formateo del dispositivo" + vbCrLf + _
            "El carrito se vaciará", 8500
    Else
        
        EE = EE + 1

        If EE < 3 Then
            SW.ShowWait "Han ocurrido errores" + vbCrLf + _
                "Se reintentará", 4500
            Resume Next
        Else
            SW.ShowWait "Han ocurrido errores" + vbCrLf + _
                "El carrito se vaciará", 4500
        End If
    End If
    
    'vaciarlo!
    Carrito.ClearCart
    SW.ShowWait "Proceso terminado con fallas", 3300
    SW.ShowWait ""
    Unload Me
    
End Sub

Private Sub ButOK_Click()
    PressOk
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaDER: SendKeys "{DOWN}"
        Case TeclaPagAt: SendKeys "+{UP}"
        Case TeclaIZQ: SendKeys "+{UP}"
        Case TeclaPagAd: SendKeys "{DOWN}"
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case TeclaESC: Unload Me
        Case TeclaCarrito: PressOk
        Case TeclaOK: Save lstPTHS.ListIndex
        Case Else 'puto enter numerico
            If KeyCode = 13 And TeclaOK = 108 Then Save lstPTHS.ListIndex
    End Select
End Sub

Private Sub PressOk()
    Save lstPTHS.ListIndex
End Sub

Private Sub Form_Load()
    tERR.Anotar "tacb"
    
    Label1.Caption = "Defina el directorio en que se copiara" + vbCrLf + _
        "Una vez que este elegido presion 'OK' o el boton de carrito"
    
    lstPTHS.Clear

    ButOK.Visible = MostrarTouch
    
End Sub

Private Sub Form_Resize()
    Label1.Left = 60
    Label1.Width = Me.Width - 120
    
    lstPTHS.Left = 200
    lstPTHS.Width = Me.Width - 400
    
    lstPTHS.Top = Label1.Top + Label1.Height + 120
    lstPTHS.Height = Me.Height - lstPTHS.Top - 120
    
    tbrPintar frmIndex.Fondoxxx, Me, 0, 0, Me.Width / 15, Me.Height / 15

    ButOK.Left = lstPTHS.Width - ButOK.Width - 160
    ButOK.Top = lstPTHS.Top + 60

End Sub

