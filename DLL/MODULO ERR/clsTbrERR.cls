VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTbrERR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private FSO As New Scripting.FileSystemObject
Private mLargoAcumula As String 'cantidad de caracteres que acumula
Private mFileLog As String 'archivo donde se registra todo
Private mFileLogGrabaTodo As String 'archivo donde se registra todo
Private mLogAcumulado As String 'acumulado de los lugares por los que paso
Private mMODOGRABATODO As Boolean
Private TextStreamGrabaTodo As TextStream

Private ContadorRegTodo As Long 'contador de lineas cuando registra todo
'al llegar 50 se borra


Public Sub Anotar(NuevoTexto As String, Optional ByVal Var1, Optional ByVal Var2, Optional ByVal Var3, Optional ByVal Var4)
    
    On Local Error GoTo AlgunaVarMal
    
    If Not IsMissing(Var1) Then
        NuevoTexto = NuevoTexto + "." + CStr(Var1)
        If Not IsMissing(Var2) Then
            NuevoTexto = NuevoTexto + "." + CStr(Var2)
            If Not IsMissing(Var3) Then
                NuevoTexto = NuevoTexto + "." + CStr(Var3)
                If Not IsMissing(Var4) Then
                    NuevoTexto = NuevoTexto + "." + CStr(Var4)
                End If
            End If
        End If
    End If
    
    mLogAcumulado = mLogAcumulado + vbCrLf + NuevoTexto
    If Len(mLogAcumulado) > mLargoAcumula Then
        mLogAcumulado = Right(mLogAcumulado, mLargoAcumula)
    End If
    
    'si esta en modo que graba todo grabar!
    If mMODOGRABATODO Then
        ContadorRegTodo = ContadorRegTodo + 1
        If ContadorRegTodo > 250 Then
            'solo los ultimos 50
            'cerrar el archivo en curso
            TextStreamGrabaTodo.Close
            'borrarlo
            FSO.DeleteFile mFileLogGrabaTodo, True
            'empezar a grabr de nuevo
            StartGrabaTodo
            ContadorRegTodo = 0
        Else
            TextStreamGrabaTodo.WriteLine NuevoTexto
        End If
        
    End If
       
    Exit Sub
    
AlgunaVarMal:
    'anotarlo como error para que no de error de automatizacion o error en la DLL
    mLogAcumulado = mLogAcumulado + vbCrLf + _
        "****" + vbCrLf + "DLL_ERR_101" + vbCrLf + _
        ErrToTXT(Err)
        
    If Len(mLogAcumulado) > mLargoAcumula Then
        mLogAcumulado = Right(mLogAcumulado, mLargoAcumula)
    End If
    Resume Next
End Sub

Public Sub StartGrabaTodo()
    Set TextStreamGrabaTodo = FSO.OpenTextFile(mFileLogGrabaTodo, ForAppending, True)
End Sub

Public Sub StopGrabaTodo()
    'se cerro correctamente, cerrar el archivo y borrarlo
    TextStreamGrabaTodo.Close
    FSO.DeleteFile mFileLogGrabaTodo, True
End Sub

Public Function ErrToTXT(sERR As ErrObject) As String
    ErrToTXT = CStr(sERR.Number) + ":" + sERR.Description + vbCrLf + sERR.Source
End Function

Public Sub AppendLog(ByVal txtAdicional As String, _
    Optional ByVal txtAdicional2 As String)
    
    'VER EL TAMAÑO DEL ARCHIVO!!!
    If FSO.FileExists(mFileLog) Then
        If FileLen(mFileLog) > 200000 Then
            'BORRARLO!!!
            FSO.DeleteFile mFileLog, True
        End If
    End If
    
    'ya tengo el log acumulado, solo necesito el error. Esto puede ser desde
    'un objeto err (con la funcion ErrToTXT o alguna descripciopn mia
    'txtadicional2 se usa cuando hay un err y ademas aclaracion mía
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(mFileLog, ForAppending, True)
        TE.WriteLine CStr(Date) + "*" + CStr(Time) + vbCrLf + _
            mLogAcumulado + vbCrLf + txtAdicional + vbCrLf + txtAdicional2 + vbCrLf
            'si esta dentro de 3PM o mprock lo muestro!!
            If LCase(mFileLog) = "d:\dev\3pm\reg3pm.log" Then
                F1.Inicializar mLogAcumulado, txtAdicional + vbCrLf + txtAdicional2
            End If
    TE.Close
    Set TE = Nothing
    
End Sub

Public Sub AppendSinHist(txt As String)
    Dim TE As TextStream
    Set TE = FSO.OpenTextFile(mFileLog, ForAppending, True)
        TE.WriteLine CStr(Timer) + ":" + txt
            'si esta dentro de 3PM o mprock lo muestro!!
            If LCase(mFileLog) = "d:\dev\3pm\reg3pm.log" Then
                F1.Inicializar mLogAcumulado, txtAdicional + vbCrLf + txtAdicional2
            End If
    TE.Close
    Set TE = Nothing
End Sub

Public Property Let FileLog(NewFile As String)
    mFileLog = NewFile
End Property

Public Property Get FileLog() As String
    FileLog = mFileLog
End Property

Public Property Let FileLogGrabaTodo(NewFile As String)
    mFileLogGrabaTodo = NewFile
End Property

Public Property Get FileLogGrabaTodo() As String
    FileLogGrabaTodo = mFileLogGrabaTodo
End Property

Public Property Let LargoAcumula(Largo As Long)
    mLargoAcumula = Largo
End Property

Public Property Get LargoAcumula() As Long
    LargoAcumula = mLargoAcumula
End Property

Private Sub Class_Initialize()
    mFileLog = "c:\log.txt"
    mLargoAcumula = 130
    mMODOGRABATODO = False
    ContadorRegTodo = 0
End Sub

Public Property Let ModoGrabaTodo(Graba As Boolean)
    mMODOGRABATODO = Graba
End Property

Public Property Get ModoGrabaTodo() As Boolean
    mMODOGRABATODO = ModoGrabaTodo
End Property

Public Property Get LogAcumulado() As String
    LogAcumulado = mLogAcumulado
End Property
