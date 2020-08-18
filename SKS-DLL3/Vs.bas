Attribute VB_Name = "Vs"
Dim FSO As New Scripting.FileSystemObject

Public Sub GetVal(Codigo As String, f As String, _
    ByRef M1, _
    ByRef M2, _
    ByRef M3)
    
    'f es el archivo con la lista!
    
    If FSO.FileExists(f) = False Then GoTo NOARCH
    
    Dim Te As TextStream
    Set Te = FSO.OpenTextFile(f, ForReading, False)
    Dim tOK As Boolean: tOK = False
    Dim Linea As String, SP() As String
    Do While Not Te.AtEndOfStream
        Linea = Te.ReadLine
        SP = Split(Linea, "|")
        If CLng(Codigo) = CLng(HEXtoLONG(SP(1))) Then
            tOK = True
            Exit Do
        End If
    Loop
    
    Dim ArrTMP() As Byte
    
    If tOK Then
        'leer los demas datos!
        'todavia estoy posicionado donde va =(sp(2))
'        i0 = CByte(HEXtoLONG(Mid(SP(2), 1, 2)))
'        i1 = CByte(HEXtoLONG(Mid(SP(2), 3, 2)))
'        i2 = CByte(HEXtoLONG(Mid(SP(2), 5, 2)))
'        i3 = CByte(HEXtoLONG(Mid(SP(2), 7, 2)))
        Dim Q As Long, Q2 As Long, POS As Long
        POS = 9
        
        For Q = 0 To 2
            For Q2 = 1 To 64
                ReDim Preserve ArrTMP(Q2 - 1)
                ArrTMP(Q2 - 1) = CByte(HEXtoLONG(Mid(SP(2), POS, 2)))
                POS = POS + 2
            Next Q2
            If Q = 0 Then M1 = ArrTMP
            If Q = 1 Then M2 = ArrTMP
            If Q = 2 Then M3 = ArrTMP
        Next Q
    Else 'no es un numero valido!
NOARCH:
        For Q = 0 To 2
            For Q2 = 1 To 64
                ReDim Preserve ArrTMP(Q2 - 1)
                ArrTMP(Q2 - 1) = 0 'CByte(HEXtoLONG(Mid(SP(2), POS, 2)))
            Next Q2
            If Q = 0 Then M1 = ArrTMP
            If Q = 1 Then M2 = ArrTMP
            If Q = 2 Then M3 = ArrTMP
        Next Q
    End If
End Sub

Private Function HEXtoLONG(N As String)
    'recibe el hex en str y devuelve un numero en str
    
    Dim Letra As String
    Dim C As Long
    Dim NumeroActual As Long
    Dim ACUM ' As Double
    For C = 1 To Len(N)
        Letra = Mid(N, C, 1)
        Select Case Letra
            Case "0", "1", "2", "3", "4", "5", "6", "7", "8", "9"
                NumeroActual = Val(Letra)
            Case "A"
                NumeroActual = 10
            Case "B"
                NumeroActual = 11
            Case "C"
                NumeroActual = 12
            Case "D"
                NumeroActual = 13
            Case "E"
                NumeroActual = 14
            Case "F"
                NumeroActual = 15
        End Select
        Dim ToSum ' As Double
        ToSum = NumeroActual * (16 ^ (Len(N) - C))
        ACUM = ACUM + ToSum
        'Label10 = Label10 + "LETRA: " + Letra + "=" + CStr(ToSum) + vbCrLf
        
    Next
    
    HEXtoLONG = CStr(ACUM)
End Function


