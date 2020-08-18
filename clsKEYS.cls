VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsKEYS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private FSO As New Scripting.FileSystemObject
Private nFOt As New tbrDATA.clsTODO

'tipos de licencia que se pueden cargar
Public Enum TypeLic
    ArchivoNoValidoComoLic = -3
    ParaOtraPC = -2 'no fue hecha para este equipo
    a_SinLeer = -1
    aSinCargar = 0 'todavia no puso nada
    BErronea = 1 'cuando ya erro 3 o mas veces
    CGratuita = 2 'primera que me puede pedir
    DMinima = 3
    EComun = 4
    FPremium = 5
    GFull = 6
    HSuperLicencia = 7
End Enum

Private mLicencia As TypeLic

Private Sub Class_Initialize()
    'GPF("cd3pm") 'DatosClaveIngresada.DLL
    'asegurarme que exista "dciLib22.dll"
    mLicencia = a_SinLeer
    
    'generar el archivo de esta PC !
    nFOt.DoNow GPF("cd4pm")
    
    'validarlo con la clave si hubiera !
    LICENCIA
End Sub

Public Sub SetFileLic(pARCH As String)
    'recibo el archivo de respuesta
    'ponerlo en la ubicacion deseada!
    FSO.CopyFile pARCH, GPF("cd7pm")
    'le hago copia de seguridad
    FSO.CopyFile GPF("cd7pm"), GPF("cd8pm")
End Sub

Public Function LICENCIA() As TypeLic
    'lee lo grabado para que sea mas rapido,
    'si es la primera vez leo todo
    If mLicencia = a_SinLeer Then IngresaClave
    LICENCIA = mLicencia
End Function

Public Sub IngresaClave()
    'ingresa archivo de licencia y devuelve el tipo esperado
    tERR.Anotar "IC01"
    Dim TmpLic As TypeLic
    mLicencia = aSinCargar
    
    'si no cargo el archivo no leer nada
    If FSO.FileExists(GPF("cd7pm")) = False Then
        tERR.Anotar "IC02"
        'a ver si esta la copia!!
        If FSO.FileExists(GPF("cd8pm")) = False Then
            tERR.Anotar "IC03"
            Exit Sub
        Else 'restaurar al original la copia
            tERR.Anotar "IC04"
            FSO.CopyFile GPF("cd8pm"), GPF("cd7pm"), True
        End If
    End If
    Dim TEZ As String 'lista "FINAL:" original
    Dim mNR As Long
    mNR = nFOt.GetNR(GPF("cd7pm"), TEZ)  'el tez queda en ni bosta!!!
    tERR.Anotar "IC05", mNR
    'si mnr es <0 no sirve el archivo !!!
    If mNR < 0 Then
        TmpLic = ArchivoNoValidoComoLic
        tERR.AppendLog "IC09"
        GoTo FIM
    End If

    TEZ = Trim(UCase(TEZ))
    tERR.Anotar "IC06", TEZ
    
    Dim FFdeLaClave As String
    FFdeLaClave = nFOt.GetRF 'da ok el de esta misma PC
    
    FFdeLaClave = Trim(UCase(FFdeLaClave))
    tERR.Anotar "IC07", FFdeLaClave
    
'    'lo escribo por las dudas
'    Dim POP4 As TextStream
'    Set POP4 = FSO.OpenTextFile(AP + "sf\fur.ia", ForWriting, True)
'        POP4.WriteLine TEZ
'        POP4.WriteLine FFdeLaClave
'    POP4.Close
    
    If TEZ <> FFdeLaClave Then
        TmpLic = ParaOtraPC
        tERR.AppendLog "IC08"
    Else 'el archivo fue hecho para esta PC
        tERR.Anotar "IC09", mNR
        Select Case mNR
            Case Is < 0
                TmpLic = ArchivoNoValidoComoLic
                tERR.AppendLog "IC09"
            Case 10, 19, 81
                TmpLic = CGratuita
            Case 30, 69, 82
                TmpLic = DMinima
            Case 40, 44, 55, 22
                TmpLic = EComun
            Case 50, 2, 9, 65
                TmpLic = FPremium
            Case 60, 67, 66, 98, 11
                TmpLic = GFull
            Case 70, 71, 72, 3, 5, 7, 8, 9, 42
                TmpLic = HSuperLicencia
            Case Else
                TmpLic = BErronea
                tERR.AppendLog "IC10"
        End Select
        
    End If
    
FIM:
    '**********************
    mLicencia = TmpLic
    '**********************
End Sub
