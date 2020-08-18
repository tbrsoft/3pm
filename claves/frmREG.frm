VERSION 5.00
Begin VB.Form frmREG 
   BackColor       =   &H00404000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de 3PM"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox lblGUID 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   570
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3960
      Width           =   3705
   End
   Begin VB.Frame frFull 
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3765
      Left            =   30
      TabIndex        =   0
      Top             =   150
      Width           =   5115
      Begin VB.ListBox lstArchReg 
         Columns         =   2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   90
         TabIndex        =   7
         Top             =   1590
         Width           =   4845
      End
      Begin VB.TextBox txtGenCod 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   810
         Width           =   4800
      End
      Begin VB.TextBox txtGenCodSL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   4800
      End
      Begin VB.TextBox txtGenCodMIN 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   330
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   450
         Width           =   4800
      End
      Begin VB.TextBox txtCodToFind 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   60
         Width           =   3165
      End
      Begin VB.CommandButton cmdGENERATE 
         Caption         =   "rareneg"
         Height          =   315
         Left            =   3330
         TabIndex        =   1
         Top             =   120
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Errores As Long 'Veces que se erro la contraseña
Dim UsosDemo As Long

Dim STRconCodigos As String 'variable fundamental que contiene todos los codigos

Dim A1 As String, A2 As String, A3 As String, A4 As String, A5 As String

Dim GGG As String

Dim ArchGUID As String 'archivo con el primer codigo que se le pidio al usuario
Dim SysF As Folder 'ubicacion de la carpeta de Windows
Dim CarpetaSys As String



Private Sub cmdGENERATE_Click()
    'para generar deno desencriptar y luego
    txtGenCodMIN = "DEMO2: " + MostraDeA5(GenerarCodigoDemo(txtCodToFind))
    txtGenCod = "FULL: " + MostraDeA5(GenerarCodigo(txtCodToFind))
    txtGenCodSL = "SL: " + MostraDeA5(GenerarCodigoSL(txtCodToFind))
End Sub

Private Sub Form_Load()
    AP = App.path
    If Right(AP, 1) <> "\" Then AP = AP + "\"
    
    'se graba en win y system
    SYSfolder = FSO.GetSpecialFolder(SystemFolder)
    WINfolder = FSO.GetSpecialFolder(WindowsFolder)
    
    If LCase(App.EXEName) <> "3pm" Then
        MsgBox "No puede cambiar el nombre del programa"
        End
    End If
    
    'SACAR los 0 y las O por la confusion (igual con las l (eles))
STRconCodigos = "87dfsdfw897564fghererh56424dfg23d4fge89r7e89rgWER7W9E8F7SD54s5d6f4sd56fw4e4w" + _
    "YUGykjgKHJBViuhyiuHHJKbkLjb567876543242324768768768jhjkhUYIULU6654654564A23E" + _
    "JF9ELFK45trF8QW78EREDF54DCDrtergdfv5asd4dfvdfgfg5sv432s4cv35sd4f88sf7d384f8s" + _
    "WUIYiuyiuyGgfgdcvxIUYPIertef54s65dfxcvxdsdf46232sd3f22WER223WE2R32WE23W2EpF32" + _
    "RE6T43D5F4G53ER4e5cvefgsdgf4v3sd5f435ssdpf3asd36r5t4354jLhjpyu534354fghd35z4s" + _
    "WE56F43SA5D4V564RT788DF4BV234RTH86RB64S35D4FG38R4BH36wer84sd54v3as54ga5er48a" + _
    "er5t4sd53fbv456564dfgwe5675s4d3fp22EG486WE68R4V3E5F4V3A36453DV436SDF4V365555" + _
    "WEFp3SAD32Vp325e454sdf3g54sd3f4553SER4G35SD4F35V43AS2pDFV32ASpCSD23A4BVNMVNM" + _
    "QWERwerqw4532sd4fpgse3g4586WEDF864AS65F4A6S5df86g4ds65fvp3s2dpf35g45rt75s4d5"
    
STRconCodigos = STRconCodigos + "yehvndis782p6655S5DLCIEUDJXLSASpXSAAAAAAQ" + _
    "UE7ENCNMSLSLPDpF8EREDSDAQWp4655SDKDICKX8p643457gfh4567ghkskdncksppwpppwwkdmx" + _
    "qwLdaqL9e9dkkLx59ejk3me9j892DJDJE899DJSDLpSD99WE23KWD992JKLWE999jqwjqw99qwdqp" + _
    "h999EDMDPWE9DLXpAASDP9DCJCM995yDaiLidi99sd9Ss9dSDFKJDF99DMCMSJS9SD9S9JXSS9S9A" + _
    "qdsjc9vIdLfyuwet2fgsh557hedfg3sgfhdfgf5yhyfhbpekefpLkefe9rjfje99e8rkefjLkd9e" + _
    "QLWLDEMDKE9ELCK9E9FE9EFA9EFK4399DUWE99JVDA9IDvbdfgS6456456456234563634568457" + _
    "ASKLJSLs9d9sd98s6s78564s34s37s8s7cyv76x8durnngic8edfgdui4jjf9dj348djk389893a" + _
    "skjhd9999S8D99CMCMGLRJEU48RJFHG8HJR8GJE89EJC9E84HJF84Jdfg345gdf84NFJKSKDKFF5"
    
STRconCodigos = STRconCodigos + "SJFU4FJKGLDpCMW949TIEMD827394A3DF939R8FJF" + _
    "W9DJDME9E9FCJChduc99d93LrLpgjkdhjd73hydbhasyxbsnvmbadfjdisdLs93dfgdnmdisd8sa" + _
    "n66d7cjkek49fmr9rfmcv9d9fnmc9dnmcv9dnc9d89eh4n4nmv9d83nf8enf83nmkLddfgaksasS9" + _
    "dfgUDMSLCLpD8894KRKF93ILJDNMMD992MDM92992MDpSpLD99299WLKCJSWLW995EJLpS95SKLpS" + _
    "ALSgfv5C9DKL3KFJLCPE953JCLLKSD95E9RJFLDLP93JDFLSLDU953KFJkLsjsLd93jfjspdjsLi" + _
    "a59dmcvf9dfgrtjcdkswLs89djd9939erjfLsd59w95sjsLpdj39rjf9535ruc95795s895sd95s" + _
    "S8D895DK3KGLFKMCNSUD5498DVHCUD7E74HDNSUW8E4HDPL539EJICJDJW59WEHJDNC59SD59JAA" + _
    "ert5545GF95394MCVJKDIW8929e95c95sLdjsLpdu95s95e893w895djsp9d589w895d895s5sss"
    
STRconCodigos = STRconCodigos + "zsmkxjkLsiLc95uu8489fndefjL349845789f89w89" + _
    "JKDHKS8989RNKDFKLFD95348953DNDNCIUIE78DHXHIASKJSDYISIYSDTFGSKJSHSJKSGCBVCG6S" + _
    "SDKDNMCNDHSUISIUDHSJKDNXCNKLSJKDHDUIWY78HDJKSDHCBNSJKSDYFKLSKLJKLSDGSDYESUDYUS" + _
    "SJKDNCN233nawesdmdjs83783678wdkusdhsikwi7w7sdghskjjkdjkSDF346LGJK48FMEKD8S83JE" + _
    "sjkdjcnsdkudi7fyeiuidjkdkLxjkchxkdustykiwLksdLkLsjcnsksuduiwiudysihcnskshjag" + _
    "asjdjcjsuyebdfnLsLduuebndndpLskLdiuiuwiueui689wiushcnkLskLjss8ILUSYS7IYDISSS" + _
    "SCCCLISWUDUW89789F5C57S76WGEBFHDGW4WTSHWH3TDBDT5CGSG5GFWHDFJSJjshstcyd63hfksbns" + _
    "sjkdnc7875Lenme89c8weienmduixhxhsyqywkdjsjkzzzzkwiwi78df6c453x4x6x7c6x5s4s5s"

      
    'ver la ubicacion del archivo de registro
    Set SysF = FSO.GetSpecialFolder(SystemFolder)
    CarpetaSys = SysF.path + "\"
    ArchREG = CarpetaSys + "rmlvf.dll"
    
    'no se pone el primer componente por qyue este cambia con cada inicio
    lblGUID = GetGuidSL
    
    txtCodSL = MostraDeA5(GenerarCodigoSL(lblGUID))
    TXTcODIGOminimo = MostraDeA5(GenerarCodigoDemo(lblGUID))
    txtCodGenerado = MostraDeA5(GenerarCodigo(lblGUID))
    
    'si esta registrada corroborar que no sea un registro de otra PC
    If FSO.FileExists(ArchREG) = False Then
        lstArchReg.AddItem "No existe el archivo de registro (ArchReg)"
    Else
        'ver si el GUID de esta maquina coincide con el del
        'registro. Esto evita que se copie el registro de una
        'maquina a otra
        
        Dim A1 As String, A2 As String, A3 As String, A4 As String
        
        Set TE = FSO.OpenTextFile(ArchREG, ForReading, False)
        'ver si el archivo no esta vacio!!!
        If TE.AtEndOfStream Then Exit Sub
        
        A1 = TE.ReadLine 'este es el guid
        A2 = TE.ReadLine 'este es la clave enviada por tbrSoft
        A3 = TE.ReadLine 'dias de demo
        A4 = TE.ReadLine 'estado actual del registro. Puede ser
            '"DEMO" todavia no ingreso contraseña
            '"DEMO2" conraseña gratuita
            '"FUCK". Intentos de crak
            '"FUCK OFF". Inhabilitado
            '"FULL". Ya lo puede usar, esta registrado OK
            '"SL" SUPERLICENCIA
        A5 = TE.ReadLine 'veces que se erro la contraseña
            
        TE.Close
        lstArchReg.AddItem "Archivo de registro (ArchReg)"
        lstArchReg.AddItem A1
        lstArchReg.AddItem A2
        lstArchReg.AddItem A3
        lstArchReg.AddItem A4
        lstArchReg.AddItem A5
    End If
End Sub

Public Function GenerarCodigo(GUID As String) As String
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod)
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEmp2 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigo = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function

Public Function GenerarCodigoSL(GUID As String) As String
    'generacion de codigos SUPELICENCIA
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod) + 40
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEmp3 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigoSL = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function

Public Function GenerarCodigoDemo(GUID As String) As String
    'generacion de codigos que dan alguna validez minima adicional
    Dim LargoCadena As Long
    
    LargoCadena = Len(STRconCodigos)
    'el largo es 2417
    Dim Parte1Cod As String, Parte2Cod As String
    Parte1Cod = txtInLista(GUID, 0, "-")
    Parte2Cod = txtInLista(GUID, 1, "-")
    
    Dim ValORIG As Long
    ValORIG = Val(Parte1Cod) + Val(Parte2Cod) + 80
    
    Dim EmpezarEnCaracter As Long
    EmpezarEnCaracter = ValORIG - (ValORIG \ LargoCadena) * LargoCadena
    If EmpezarEnCaracter <= 0 Then EmpezarEnCaracter = -EmpezarEnCaracter + 1
    If EmpezarEnCaracter > LargoCadena - 120 Then EmpezarEnCaracter = EmpezarEnCaracter - 120
    txtEMP1 = "nClave: " + CStr(EmpezarEnCaracter)
    
    GenerarCodigoDemo = Mid(STRconCodigos, EmpezarEnCaracter, 40)
End Function

Public Function MostraDeA5(TXT As String)
    Dim c As Long, Letra As String, newTXT As String
    c = 0
    Do While c < Len(TXT)
        Letra = Mid(TXT, c + 1, 5)
        newTXT = newTXT + Letra
        c = c + 5
        If c < Len(TXT) Then newTXT = newTXT + "-"
    Loop
    MostraDeA5 = newTXT
End Function

Private Sub txtCodGenerado_Change()

End Sub

Private Sub txtCodSL_Change()

End Sub
