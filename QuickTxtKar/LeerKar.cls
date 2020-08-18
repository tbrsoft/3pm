VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LeerKar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mvarIsLyric As Boolean
Private mvarTextLyric As String

Dim TxtString As String, Bpm As Single, ppqn, Clk, Den, Quarter As Double
Dim Xini As Long, Ret As Long, IsPlaying As Boolean
Dim Lyr() As Lyric, Frase() As String

Private Type Lyric
    TempoTotal As Double
    TempoAtual As Double
    TxtString As String
    TxtStringLen As Integer
    TextStart As Long
    FraseIndex As Integer
End Type

Private Property Let TextLyric(ByVal vData As String)
    mvarTextLyric = vData
End Property

Public Property Get TextLyric() As String
    GetTextLyric
    TextLyric = mvarTextLyric
End Property

Private Property Let IsLyric(ByVal vData As Boolean)
    mvarIsLyric = vData
End Property


Public Property Get IsLyric() As Boolean
    IsLyric = mvarIsLyric
End Property

Private Sub CreateType()
    If Not IsLyric Then Exit Sub
    Dim N As Integer, OldPause As Double, C As Integer, K As Integer, W As Integer
    K = 1: W = 0: Nx = 0: Ni = 0
    Do
        Nx = Nx + 1
        NewText = ""
        K = InStr(K, TxtString, Chr(255))
        If K = 0 Then Exit Sub
        CompText = Val(Asc(Mid(TxtString, K + 2, 1)))
        Select Case Val(Asc(Mid(TxtString, K + 1, 1)))
            Case 1
                If Mid(TxtString, K + 3, 1) <> "@" Then
                    NewText = Mid(TxtString, K + 3, CompText)
                Else
                    If Mid(TxtString, K + 3, 2) = "@T" Then
                        Titles.Add StrConv(Mid(TxtString, K + 5, CompText - 2), vbProperCase)
                    End If
                End If
            Case 47
                    Exit Sub
        End Select
        Pausa1 = Val(Asc(Mid(TxtString, (K + 3) + CompText, 1)))
        Pausa2 = Val(Asc(Mid(TxtString, (K + 4) + CompText, 1)))
        Pausa3 = Val(Asc(Mid(TxtString, (K + 5) + CompText, 1)))
        Pausa4 = Val(Asc(Mid(TxtString, (K + 6) + CompText, 1)))
        If Pausa1 = 255 Then Pausa1 = 0
        If Pausa2 = 255 Then Pausa2 = 0
        If Pausa3 = 255 Then Pausa3 = -1
        If Pausa4 = 255 Then Pausa4 = -1
        K = K + 1
        Select Case Left(NewText, 1)
            Case Is = "/", "\"
                C = 0
                NewText = Right(NewText, Len(NewText) - 1)
                CompText = CompText - 1
                W = W + 1
            Case "@"
                NewText = ""
            Case Else
        End Select
        ReDim Preserve Lyr(Nx)

        Valor = 0
        If Pausa1 >= 127 Then
            If Pausa3 >= 0 Then
                If Pausa4 >= 0 Then
                    Valor = Valor And &H7F
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa1 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa2 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa3 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa4 And &H7F)
                    Pausa = Valor * Quarter
                Else
                    Valor = Valor And &H7F
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa1 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa2 And &H7F)
                    Valor = Valor * 128
                    Valor = Valor Or (Pausa3 And &H7F)
                    Pausa = Valor * Quarter
                End If
            Else
                Pausa = (((Pausa1 And &H7F) * 128) Or (Pausa2 And &H7F)) * Quarter
            End If
        Else
            Pausa = Pausa1 * Quarter
        End If
        OldPause = OldPause + Pausa
        Lyr(Nx).TxtString = ""
        Lyr(Nx).TxtStringLen = 0
        Lyr(Nx).TextStart = 0
        Lyr(Nx).TxtString = NewText
        Lyr(Nx).TxtStringLen = Len(NewText)
        Lyr(Nx).TextStart = C
        Lyr(Nx).TempoAtual = Pausa
        Lyr(Nx).TempoTotal = OldPause
        Lyr(Nx).FraseIndex = W
        C = C + CompText
    Loop

End Sub
Private Sub GetTextLyric()
    If Not IsLyric Then Exit Sub
    j = 1
    N = 0
    ReDim Frase(0)
    Do
        K = InStr(j, TxtString, Chr(255) & Chr(1))
        If K = 0 Then Exit Do
        j = K + 1
        CompText = Val(Asc(Mid(TxtString, K + 2, 1)))
        NewText = Mid(TxtString, K + 3, CompText)
        NovaFrase = False
        Select Case Left(NewText, 1)
            Case Is = "/", "\"
                NewText = Chr(13) & Chr(10) & Right(NewText, Len(NewText) - 1)
                N = N + 1
                ReDim Preserve Frase(N)
            Case "@"
                NewText = " "
        End Select
        Tx = Tx & NewText
        If Left(NewText, 2) = Chr(13) & Chr(10) Then
            NewText = Right(NewText, Len(NewText) - 2)
        End If
        If Trim(NewText) <> "" Then Frase(N) = Frase(N) & NewText
    Loop
    TextLyric = Tx
End Sub
Private Sub OpenFile(FileName As String)
        
    On Error Resume Next
    Close #1
    Open FileName For Binary As #1
    TxtString = Space(FileLen(FileName))
    Get #1, , TxtString
    Track = TxtString
    
    K = InStr(1, TxtString, Chr(255) & Chr(88))
    If K = 0 Then Exit Sub
    Num = Asc(Mid(TxtString, K + 3, 1))
    Den = Asc(Mid(TxtString, K + 4, 1))
    Select Case Den
        Case 1: Den = 2
        Case 2: Den = 4
        Case 3: Den = 8
        Case 4: Den = 16
        Case 5: Den = 32
        Case 6: Den = 64
        Case 7: Den = 128
        Case 8: Den = 256
        Case 9: Den = 512
        Case 10: Den = 1024
    End Select
    Clk = Asc(Mid(TxtString, K + 5, 1)) / Den
    
    K = InStr(1, TxtString, Chr(255) & Chr(81))
    If K = 0 Then Exit Sub
    
    n1 = Format(Hex(Asc(Mid(TxtString, K + 3, 1))), "00")
    n2 = Format(Hex(Asc(Mid(TxtString, K + 4, 1))), "00")
    N3 = Format(Hex(Asc(Mid(TxtString, K + 5, 1))), "00")
    If Len(n1) = 1 Then n1 = "0" & n1
    If Len(n2) = 1 Then n2 = "0" & n2
    If Len(N3) = 1 Then N3 = "0" & N3
    
    
    T1 = Format(Hex(Asc(Mid(TxtString, 13, 1))), "00")
    T2 = Format(Hex(Asc(Mid(TxtString, 14, 1))), "00")
    If Len(T1) = 1 Then T1 = "0" & T1
    If Len(T2) = 1 Then T2 = "0" & T2
    
    TimeBase = CDec("&H" & T1 & T2) / 4
    
    Bpm = Format(60000000 / CDec("&H" & n1 & n2 & N3), "0.00")
    ppqn = CDec("&H" & n1 & n2 & N3)
    Quarter = (ppqn / TimeBase) / 4000
    QuarterNote = Num & "/" & Den
    
    K = InStr(1, TxtString, Chr(255) & Chr(3) & Chr(5) & "W")
    If K = 0 Then
        K = InStr(1, TxtString, "@T")
        If K = 0 Then
            IsLyric = False
            Exit Sub
        End If
    End If
    IsLyric = True
    
    For T = K To 1 Step -1
        If UCase(Mid(TxtString, T, 4)) = "MTRK" Then
            IniTrack = T
            Exit For
        End If
    Next

    V1 = Format(Hex(Asc(Mid(TxtString, IniTrack + 4, 1))), "00")
    V2 = Format(Hex(Asc(Mid(TxtString, IniTrack + 5, 1))), "00")
    V3 = Format(Hex(Asc(Mid(TxtString, IniTrack + 6, 1))), "00")
    V4 = Format(Hex(Asc(Mid(TxtString, IniTrack + 7, 1))), "00")
    If Len(V1) = 1 Then V1 = "0" & V1
    If Len(V2) = 1 Then V2 = "0" & V2
    If Len(V3) = 1 Then V3 = "0" & V3
    If Len(V4) = 1 Then V4 = "0" & V4


    TxtString = Space(CDec("&H" & Va & V2 & V3 & V4) + 100)
    Track = Right(Track, Len(Track) - (IniTrack + CDec("&H" & Va & V2 & V3 & V4)))
    
    Seek #1, IniTrack + 8
    Get #1, , TxtString
    TxtString = Right(TxtString, Len(TxtString) - 0)
    Close #1
    CreateType
    GetTextLyric

End Sub

Public Function OpenDevice(FileName As String) As Boolean
    ReDim Lyr(0)
    ReDim Frase(0)
    OpenFile FileName
    DoEvents
End Function
