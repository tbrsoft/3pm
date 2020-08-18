VERSION 5.00
Object = "{181111E6-07C8-4D47-8611-3BF038099354}#5.2#0"; "tbrFaroButton.ocx"
Begin VB.Form fNEW 
   BackColor       =   &H00000000&
   Caption         =   "Test de comunicaciones tbrSoft"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9360
   Icon            =   "fNEW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   5085
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin tbrFaroButton.fBoton XxBoton1 
      Height          =   435
      Index           =   0
      Left            =   4110
      TabIndex        =   12
      Top             =   2640
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   767
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "Boton"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton xPORC2 
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   2340
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   556
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "0"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton xPORC 
      Height          =   285
      Left            =   150
      TabIndex        =   10
      Top             =   2040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   503
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "0"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin tbrFaroButton.fBoton XxBoton2 
      Height          =   465
      Left            =   3810
      TabIndex        =   8
      Top             =   1800
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   820
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   "Grabar registro"
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
   Begin VB.Frame frINI 
      BackColor       =   &H00000000&
      Caption         =   "Parametros e inicio"
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
      Height          =   1725
      Left            =   150
      TabIndex        =   2
      Top             =   60
      Width           =   2985
      Begin tbrFaroButton.fBoton XIni 
         Height          =   465
         Left            =   120
         TabIndex        =   7
         Top             =   1170
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   820
         fFColor         =   6553600
         fBColor         =   16761024
         fCapt           =   "Iniciar pruebas"
         fEnabled        =   -1  'True
         fFontN          =   ""
         fFontS          =   0
         fECol           =   16777215
      End
      Begin VB.TextBox txVUELTAS 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Text            =   "10"
         Top             =   810
         Width           =   795
      End
      Begin VB.TextBox txTimeOUT 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Text            =   "1000"
         Top             =   330
         Width           =   795
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vueltas de prueba"
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
         Height          =   465
         Left            =   1020
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TimeOut en milisegundos"
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
         Height          =   465
         Left            =   960
         TabIndex        =   3
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   3690
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox tLOG 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1395
      Left            =   3810
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   390
      Width           =   2895
   End
   Begin tbrFaroButton.fBoton xBASE 
      Height          =   735
      Left            =   90
      TabIndex        =   9
      Top             =   1980
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1296
      fFColor         =   6553600
      fBColor         =   16761024
      fCapt           =   ""
      fEnabled        =   -1  'True
      fFontN          =   ""
      fFontS          =   0
      fECol           =   16777215
   End
End
Attribute VB_Name = "fNEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private FileLog As String
Private FSO As New Scripting.FileSystemObject
Private TE As TextStream

Private Sub mLog(T As String, Optional bShow As Boolean = True)
    If bShow Then
        tLOG.Text = tLOG.Text + CStr(Timer * 100) + ": " + T + vbCrLf
        tLOG.SelStart = Len(tLOG) - 1
        tLOG.Refresh
    End If
    
    On Local Error Resume Next 'por si empeiza antes de iniciar pruebas
    'sea como sea se escribe en un archivo de texto
    TE.WriteLine CStr(Timer * 100) + " " + T
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Dim Azar As Integer
    'Randomize
    'Azar = Int(Rnd * 22) + 1
    'If KeyCode = vbKeyF9 Then Text1.Text = "sD:" + CStr(Azar) 'no suma al counter de la placa
    'If KeyCode = vbKeyF8 Then Text1.Text = "sD:14"
End Sub

Private Sub Form_Load()
    Dim FSO As New Scripting.FileSystemObject
    SF = FSO.GetSpecialFolder(SystemFolder)
    If Right(SF, 1) <> "\" Then SF = SF + "\"
    
    Me.Caption = "Test H2k               Versión " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Local Error Resume Next
    TE.Close
End Sub

Private Sub Form_Resize()
    'acomodar los 24 botones
    XxBoton1(0).Left = 30
    XxBoton1(0).Width = (Me.Width - (30 * 9)) / 4
    XxBoton1(0).Height = Me.Height / 16
    XxBoton1(0).Top = Me.Height - ((XxBoton1(0).Height + 30) * 7)
    
    'lineas de porcentaje
    xBASE.Left = 30
    xBASE.Width = Me.Width - 160
    xBASE.Top = XxBoton1(0).Top - 60 - xBASE.Height
    
    xPORC.Left = 60: xPORC2.Left = 60
    xPORC.Top = xBASE.Top + 30: xPORC2.Top = xBASE.Top + xBASE.Height - xPORC2.Height - 30
    xPORC.Width = 240: xPORC2.Width = 240
    xPORC.Caption = "0": xPORC2.Caption = "0"
    
    'log
    XxBoton2.Left = Me.Width - XxBoton2.Width - 60
    XxBoton2.Top = xBASE.Top - XxBoton2.Height - 60
    
    tLOG.Width = Me.Width / 2
    tLOG.Left = Me.Width - tLOG.Width - 150
    tLOG.Top = 30
    tLOG.Height = XxBoton2.Top - 60
    
    
    'xINI
    frINI.Top = xBASE.Top / 2 - frINI.Height / 2
    frINI.Left = tLOG.Left / 2 - frINI.Width / 2
    
    On Local Error Resume Next
    
    Dim J As Long
    For J = 1 To 23
        Unload XxBoton1(J)
    Next J
    
    For J = 0 To 22
        If J > 0 Then Load XxBoton1(J)
        
        XxBoton1(J).Width = XxBoton1(0).Width
        XxBoton1(J).Height = XxBoton1(0).Height
        
        If J / 4 = J \ 4 Then
            XxBoton1(J).Left = 30
            XxBoton1(J).Top = XxBoton1(J - 1).Top + XxBoton1(J - 1).Height + 30
        Else
            XxBoton1(J).Left = XxBoton1(J - 1).Left + XxBoton1(J - 1).Width + 30
            XxBoton1(J).Top = XxBoton1(J - 1).Top
        End If
        
        If J < 21 Then
            XxBoton1(J).Caption = "Boton " + CStr(J + 1)
        Else
            XxBoton1(J).Caption = "Monedero " + CStr(J - 20)
        End If
        
        XxBoton1(J).Visible = True
    Next J
    
End Sub

Private Sub Text1_Change()
    
    On Local Error GoTo TR
    
    Dim P As String
    P = Text1.Text
    
    If P = "" Then Exit Sub
    
    mLog Text1.Text, False
    
    Dim SP() As String
    SP = Split(P, ":")
    
    If SP(0) = "sD" Then
        
        'actualizar los contadores
        Dim I As Long
        Dim GC As Long
        I = CLng(SP(1))
        
        mLog "Se apreto " + CStr(I), True
        
        'hasta los botones que tengo!
        If I > 23 Then
            mLog " **** SIGNALIN: " + CStr(I), True
            Exit Sub
        End If
        
        'solo el contador que escuche
        GC = S3.GetCounter(I)
        If I = 0 Then Exit Sub 'solucionado abril 2008 "I" NO PUEDE SER CERO!!
        If I < 22 Then
            XxBoton1(I - 1).Caption = "Boton " + CStr(I) + ": " + CStr(GC)
            
        Else
            XxBoton1(I - 1).Caption = "Monedero " + CStr(I - 21) + ": " + CStr(GC)
        End If
        
'        For I = 0 To 22
'            GC = S3.GetCounter(I + 1)
'
'            If I < 21 Then
'                XxBoton1(I).Caption = "Boton " + String(2 - Len(CStr(I + 1)), "0") + CStr(I + 1) + ": " + _
'                    String(4 - Len(GC), "0") + CStr(GC)
'
'            Else
'                XxBoton1(I).Caption = "Monedero " + String(2 - Len(CStr(I - 20)), "0") + CStr(I - 20) + ": " + _
'                    String(3 - Len(CStr(GC)), "0") + CStr(GC)
'            End If
'        Next I
       
    End If
    
    Exit Sub
    
TR:
    mLog "ERROR: " + CStr(Err.Number) + " (" + Err.Description + ")" + vbCrLf + "..." + CStr(I) + "/" + CStr(GC)
    Resume Next
End Sub

Private Sub xINI_Click()
    
    On Local Error GoTo errTEST
    
    If Not IsNumeric(txTimeOUT) Then
        MsgBox "Parametro TimeOut mal definido !!"
        Exit Sub
    End If
    
    Dim TimOut As Single  'TimeOut
    TimOut = CLng(txTimeOUT.Text) / 1000
    
    If Not IsNumeric(txVUELTAS.Text) Then
        MsgBox "Parametro vueltas mal definido !!"
        Exit Sub
    End If
    
    Dim Vueltas As Long
    Vueltas = CLng(txVUELTAS)
    
    'probas X cantidad de licencias
    ' y pedirle que presione algunas teclas
    
    FileLog = App.Path
    If Right(FileLog, 1) <> "\" Then FileLog = FileLog + "\"
    FileLog = FileLog + "Log" + CStr(Year(Now)) + "." + _
                                CStr(Month(Now)) + "." + _
                                CStr(Day(Now)) + "." + _
                                CStr(Hour(Now)) + "." + _
                                CStr(Minute(Now)) + ".txt"
    
    
    'si lo usa 2 veces en el mismo minuto puede joder
    If FSO.FileExists(FileLog) Then FSO.DeleteFile FileLog
    
    Set TE = FSO.OpenTextFile(FileLog, ForAppending, True)
    S3.INIT
    S3.HwndMsg = Text1.hWnd
    S3.ReIniCounters
    S3.Prender
    mLog "Conectado"
    mLog "HWND: " + CStr(Text1.hWnd), False
    
    esperar 1
    
    S3.ReIniContLuis
    
    esperar 1
    
    'obtener el numero de placa
    Dim NP As Long
    NP = CLng(S3.GetnPlaca(SF + "prec.dll"))
    If NP = -1 Then
        mLog "No se podido comenzar la prueba. Quizas la interfase no este conectada o sea una versión solo botones"
        Exit Sub
    End If
    
    mLog "placa id: " + CStr(NP)
    
    mLog "Recarga..."
    esperar 2
    
    On Local Error Resume Next
    
    Dim J As Long, RET As Long, cRet As Long
    cRet = 0
    For J = 1 To Vueltas
        mLog "********* " + CStr(J) + " *****************"
        RET = S3.AddCont(J Mod 4, TimOut)
        
        If RET = 2 Then 'time out
            mLog "***** TIME OUT - CONT:" + CStr(J) + " (timeout!) " + S3.GetResLicSTR
        End If
        
        If RET = 1 Then 'llego mal!!! poner en cero
            mLog "***** MAL CONT:" + CStr(J) + " (bad!) " + S3.GetResLicSTR
            'reinicio todo!!!
            S3.ReIniContLuis
            esperar 1
        End If
        
        If RET = 0 Then
            mLog "CONT:" + CStr(J) + " (ok!) " + S3.GetResLicSTR
            cRet = cRet + 1
        End If
        
        xPORC.Width = xBASE.Width * (J / Vueltas)
        xPORC.Caption = CStr(J)
        
        Dim W2 As Single
        W2 = xBASE.Width * CSng(cRet / J)
        xPORC2.Width = CLng(W2)
        xPORC2.Caption = CStr(cRet) + " (" + CStr(Round(CSng(cRet / J), 2) * 100) + " %)"
        
    Next J
    
    Exit Sub
errTEST:
    'si el tipo pone iniciar pruebas dos veces TE nunca se cerro
    If Err.Number = 70 Then
        TE.Close
        Resume
    Else
        MsgBox "Error, reinicie esta aplicación (" + CStr(Err.Number) + ")" + vbCrLf + Err.Description
        
    End If
End Sub

Private Sub esperar(N As Single)
    N = Timer + N
    Do While Timer < N
        DoEvents
    Loop
End Sub

Private Sub XxBoton2_Click()
    On Local Error Resume Next
    TE.Close
    
    'grabarlo!
End Sub
